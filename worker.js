/* Roda o Formatador SEFA (Python) dentro do navegador, num Web Worker.
   A página (app.js) manda pedidos {id, acao, dados} e recebe {id, ok, resultado | erro}.

   Ao atualizar o programa: copie os .py novos para a pasta py/ e aumente VERSAO,
   para que os navegadores não usem os arquivos antigos guardados em cache. */

const VERSAO = "3.18-web11";

// Módulos do programa desktop (sem alteração) + a ponte com a página
const ARQUIVOS = [
  "web_ponte.py", "extrator_pdf.py", "importar_pdf.py", "formatador_sefa.py",
  "links_no_word.py", "links_legislacao.py", "modelo_sefa.docx", "tabela_links.tsv",
];

// Bibliotecas: as primeiras vêm da pasta pyodide/, as .whl da pasta wheels/
const PACOTES = [
  "lxml", "pillow", "cryptography", "charset-normalizer", "typing-extensions",
  "wheels/python_docx-1.2.0-py3-none-any.whl",
  "wheels/pdfminer_six-20251107-py3-none-any.whl",
  "wheels/pdfplumber-0.11.8-py3-none-any.whl",
];

// PyMuPDF (18 MB): usa a cópia da pasta wheels/ se ela existir; senão, baixa do PyPI
// (files.pythonhosted.org, o repositório oficial de bibliotecas Python), que libera o arquivo
// para qualquer site. O endereço do PyPI é fixo para esta versão e não muda.
const PYMUPDF = "pymupdf-1.28.2-cp313-abi3-pyemscripten_2025_0_wasm32.whl";
const PYMUPDF_PYPI = "https://files.pythonhosted.org/packages/58/8c/" +
  "d897dcd32a25b58186c968b15ce4324ca029e9d96460de12325314e390be/" + PYMUPDF;

async function origemPyMuPDF() {
  try {
    const r = await fetch(`${base}wheels/${PYMUPDF}`, { method: "HEAD" });
    if (r.ok) return { url: `${base}wheels/${PYMUPDF}`, local: true };
  } catch { /* sem a cópia local */ }
  return { url: PYMUPDF_PYPI, local: false };
}

importScripts("pyodide/pyodide.js");

let py = null;
let ponte = null;
const base = new URL("./", self.location).href;

function progresso(texto) {
  self.postMessage({ tipo: "progresso", texto });
}

async function iniciar() {
  progresso("Carregando o Python");
  py = await loadPyodide({ indexURL: base + "pyodide/" });
  const pymupdf = await origemPyMuPDF();
  progresso(pymupdf.local ? "Carregando as bibliotecas de PDF e Word"
    : "Carregando as bibliotecas de PDF e Word (o PyMuPDF vem do PyPI; na primeira vez, são 18 MB)");
  await py.loadPackage(
    [...PACOTES.map((p) => (p.endsWith(".whl") ? base + p : p)), pymupdf.url],
    { messageCallback: () => {}, errorCallback: (m) => console.warn(m) },
  );
  if (!py.runPython("import importlib.util; importlib.util.find_spec('fitz') is not None")) {
    throw new Error(pymupdf.local
      ? `não foi possível carregar wheels/${PYMUPDF}.`
      : "não foi possível baixar o PyMuPDF do PyPI (files.pythonhosted.org). Se a rede bloqueia esse endereço, " +
        `coloque o arquivo ${PYMUPDF} na pasta wheels do site e publique de novo.`);
  }
  progresso("Carregando o formatador");
  py.FS.mkdirTree("/formatador");
  for (const nome of ARQUIVOS) {
    const resposta = await fetch(`${base}py/${nome}?v=${VERSAO}`);
    if (!resposta.ok) throw new Error(`não encontrei py/${nome} no site (erro ${resposta.status})`);
    py.FS.writeFile(`/formatador/${nome}`, new Uint8Array(await resposta.arrayBuffer()));
  }
  py.runPython("import sys; sys.path.insert(0, '/formatador')");
  ponte = py.pyimport("web_ponte");
  py.FS.mkdirTree("/work/saida");
  py.FS.mkdirTree("/work/previa");
  return JSON.parse(ponte.versao());
}

const pronto = iniciar();

function json(texto) {
  return JSON.parse(texto);
}

function lerSeExistir(caminho) {
  try {
    return py.FS.readFile(caminho);
  } catch {
    return null;
  }
}

const acoes = {
  async versao() {
    return await pronto;
  },
  analisar({ pdf, consulta, maxPaginas }) {
    py.FS.writeFile("/work/entrada.pdf", new Uint8Array(pdf));
    return json(ponte.analisar("/work/entrada.pdf", consulta, maxPaginas));
  },
  estado() {
    return json(ponte.estado());
  },
  municipios({ valor }) {
    return json(ponte.definir_municipios(valor));
  },
  editarParagrafo({ indice, texto, estilo }) {
    return json(ponte.editar_paragrafo(indice, texto, estilo));
  },
  editarCelula({ tabela, linha, coluna, texto }) {
    return json(ponte.editar_celula(tabela, linha, coluna, texto));
  },
  conferirTabela({ tabela }) {
    return json(ponte.conferir_tabela(tabela));
  },
  trecho({ indice }) {
    return json(ponte.trecho_formatado(indice));
  },
  aplicarTrecho({ indice, runs, estilo, formatou }) {
    return json(ponte.aplicar_trecho(indice, runs, estilo, formatou));
  },
  apagar({ indice }) {
    return json(ponte.apagar_paragrafo(indice));
  },
  juntar({ indice }) {
    return json(ponte.juntar_paragrafos(indice));
  },
  dividir({ indice, antes, depois, estilo, formatou }) {
    return json(ponte.dividir_paragrafo(indice, antes, depois, estilo, formatou));
  },
  desfazer() {
    return json(ponte.desfazer());
  },
  conferirLimites() {
    return json(ponte.conferir_limites());
  },
  recorte({ tabela, linha }) {
    const r = json(ponte.recorte_tabela(tabela, linha, "/work/recorte.png"));
    if (r.ok) r.png = py.FS.readFile("/work/recorte.png");
    return r;
  },
  previa({ modelo, links, linksUsuario }) {
    let caminhoModelo = "";
    if (modelo) {
      caminhoModelo = "/work/modelo_usuario.docx";
      py.FS.writeFile(caminhoModelo, new Uint8Array(modelo));
    }
    const r = json(ponte.previa_word("/work/previa/previa.docx", caminhoModelo, links, linksUsuario || "{}"));
    if (r.ok) r.wordBytes = py.FS.readFile(r.word);
    return r;
  },
  gerar({ nome, modelo, relatorio, links, linksUsuario }) {
    let caminhoModelo = "";
    if (modelo) {
      caminhoModelo = "/work/modelo_usuario.docx";
      py.FS.writeFile(caminhoModelo, new Uint8Array(modelo));
    }
    const r = json(ponte.gerar(`/work/saida/${nome}`, caminhoModelo, relatorio, links, linksUsuario || "{}"));
    if (r.ok) {
      r.wordBytes = lerSeExistir(r.word);
      r.relatorioBytes = r.relatorio ? lerSeExistir(r.relatorio) : null;
    }
    return r;
  },
};

self.onmessage = async ({ data }) => {
  const { id, acao, dados } = data;
  try {
    await pronto;
    const resultado = await acoes[acao](dados || {});
    self.postMessage({ id, ok: true, resultado });
  } catch (erro) {
    self.postMessage({ id, ok: false, erro: String(erro && erro.message ? erro.message : erro) });
  }
};

pronto.then(
  (info) => self.postMessage({ tipo: "pronto", info }),
  (erro) => self.postMessage({ tipo: "falha", erro: String(erro && erro.message ? erro.message : erro) }),
);
