from __future__ import annotations
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from docx import Document
from docx.oxml.ns import qn
from .model import SourceLine, digest


def libreoffice():
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for env in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
        p = Path(os.environ.get(env, "C:/Program Files")) / "LibreOffice/program/soffice.exe"
        if p.is_file():
            return str(p)
    return None


def convert_doc(path: Path, cache: Path) -> Path:
    """Converte uma cópia. O documento original nunca é salvo pelo conversor."""
    key = digest(path)[:16]
    destination = cache / key
    destination.mkdir(parents=True, exist_ok=True)
    result = destination / "entrada.docx"
    if result.exists():
        return result
    errors = []
    if sys.platform == "win32":
        word = doc = None
        initialized = False
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            initialized = True
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            word.Options.UpdateLinksAtOpen = False
            doc = word.Documents.Open(str(path.resolve()), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
            doc.SaveAs2(str(result.resolve()), FileFormat=16, AddToRecentFiles=False)
            return result
        except Exception as exc:
            errors.append("Word: " + str(exc))
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            if initialized:
                pythoncom.CoUninitialize()
    binary = libreoffice()
    if binary:
        # Um perfil isolado evita interferir com o LibreOffice aberto pelo usuário.
        with tempfile.TemporaryDirectory(prefix="sefa_office_") as profile:
            source = destination / "entrada.doc"
            shutil.copy2(path, source)
            proc = subprocess.run([binary, "-env:UserInstallation=" + Path(profile).as_uri(), "--headless",
                "--convert-to", "docx", "--outdir", str(destination.resolve()), str(source.resolve())],
                capture_output=True, text=True, timeout=150,
                creationflags=0x08000000 if os.name == "nt" else 0)
            if result.is_file():
                return result
            errors.append(proc.stdout + proc.stderr)
    raise ValueError("Não foi possível abrir o arquivo .doc. Abra-o no Word e use Salvar como > Documento do Word (.docx), "
                     "ou instale o LibreOffice gratuito. " + " ".join(errors)[:700])


def prepare_base(path, cache):
    p = Path(path).resolve()
    if not p.is_file():
        raise ValueError("O arquivo da legislação base não foi encontrado.")
    if p.suffix.lower() == ".doc":
        return convert_doc(p, Path(cache))
    if p.suffix.lower() != ".docx":
        raise ValueError("A legislação base deve estar em .docx ou .doc.")
    return p


def docx_source(path):
    doc = Document(path)
    result = []
    page = 1
    # Paginação armazenada no DOCX não é a paginação calculada do Word.
    # Para não inventar páginas, relatamos os parágrafos para arquivos Word.
    for i, el in enumerate(doc.element.body):
        if el.tag == qn("w:p"):
            from docx.text.paragraph import Paragraph
            p = Paragraph(el, doc._body)
            text = p.text
            for line in text.replace("\r", "\n").split("\n"):
                if line.strip():
                    result.append(SourceLine(line.strip(), i + 1, None, p.style.name))
        elif el.tag == qn("w:tbl"):
            text = " | ".join(t.text or "" for t in el.iter(qn("w:t")))
            result.append(SourceLine("[TABELA] " + text, i + 1, None, "", "table"))
    return result


TERMINATOR = re.compile(r'(?:[.;:!?”"\)]|\(\s*N\.?\s?R\.?\s*\)|\(\s*A\.?\s?C\.?\s*\))\s*(?:e|ou)?\s*$', re.I)


def join_wrapped_labels(lines):
    """No PDF, uma linha quebrada pode começar por "§ 2º do art. 198" ou "art. 5º".
    Se a linha anterior não terminou a frase, é continuação e não novo dispositivo."""
    from .index import label
    out = []
    for line in lines:
        prev = out[-1] if out else None
        if (prev is not None and prev.page == line.page and label(line.text)
                and not TERMINATOR.search(prev.text) and not re.match(r"^[“\"]", line.text)):
            out[-1] = SourceLine(prev.text + " " + line.text, prev.index, prev.page, prev.style, prev.kind)
            continue
        out.append(line)
    return out


def read_amender(path, cache):
    p = Path(path).resolve()
    ext = p.suffix.lower()
    warnings = []
    if ext in (".docx", ".doc"):
        p = prepare_base(p, cache)
        d = Document(p)
        if d.element.xpath("//w:ins|//w:del"):
            raise ValueError("O ato alterador contém alterações controladas do Word. Aceite ou rejeite essas alterações em uma cópia antes de importar.")
        return docx_source(p), warnings
    if ext in (".txt", ".html", ".htm"):
        raw = p.read_bytes()
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("cp1252")
        if ext != ".txt":
            from html.parser import HTMLParser
            class Text(HTMLParser):
                def __init__(self):
                    super().__init__(); self.out = []; self.hidden = 0
                def handle_starttag(self, tag, attrs):
                    if tag in ("script", "style"):
                        self.hidden += 1
                    if tag in ("p", "div", "br", "tr", "h1", "h2", "h3", "li"):
                        self.out.append("\n")
                def handle_endtag(self, tag):
                    if tag in ("script", "style"):
                        self.hidden = max(0, self.hidden - 1)
                    if tag in ("p", "div", "tr", "li"):
                        self.out.append("\n")
                def handle_data(self, data):
                    if not self.hidden:
                        self.out.append(data)
            parser = Text(); parser.feed(text); text = "".join(parser.out)
            warnings.append("HTML importado como texto. Tabelas e disposição visual precisam ser conferidas no arquivo de origem.")
        return [SourceLine(s.strip(), i + 1) for i, s in enumerate(text.splitlines()) if s.strip()], warnings
    if ext == ".pdf":
        import fitz
        result = []
        with fitz.open(p) as doc:
            if doc.needs_pass:
                raise ValueError("O PDF está protegido por senha. Use uma cópia desbloqueada.")
            for n, page in enumerate(doc, 1):
                text = page.get_text("text", sort=True)
                if len(text.strip()) < 35:
                    warnings.append(f"Página {n} do PDF sem texto suficiente. Faça OCR antes de usar os comandos dessa página.")
                # Linhas de PDF são preservadas; o analisador reúne continuações.
                for i, line in enumerate(text.splitlines(), 1):
                    if line.strip():
                        result.append(SourceLine(line.strip(), i, n))
        if not result:
            raise ValueError("O PDF não contém texto selecionável. Faça OCR ou importe o ato em Word ou TXT.")
        result = join_wrapped_labels(result)
        warnings.append("PDF importado por extração textual. A ordem de colunas, quebras e tabelas deve ser conferida antes da geração.")
        return result, warnings
    raise ValueError("Use ato alterador em DOCX, DOC, PDF com texto, TXT ou HTML.")
