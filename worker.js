/* Roda o Formatador SEFA (Python) dentro do navegador, num Web Worker.
   A página (app.js) manda pedidos {id, acao, dados} e recebe {id, ok, resultado | erro}.

   Ao atualizar o programa: copie os .py novos para a pasta py/ e aumente VERSAO,
   para que os navegadores não usem os arquivos antigos guardados em cache. */

const VERSAO = "3.24-web2";

// Módulos do programa desktop (sem alteração) + a ponte com a página
const ARQUIVOS = [
  "web_ponte.py", "extrator_pdf.py", "importar_pdf.py", "formatador_sefa.py",
  "links_no_word.py", "links_legislacao.py", "modelo_sefa.docx", "tabela_links.tsv",
  // Consolidador SEFA (aba Aplicar alterações): módulos do programa sem alteração + a ponte
  "consolidador_ponte.py", "modelos/estilos_sefa.xml", "modelos/estilos_sefa_atual.xml",
  ...["__init__", "editorial", "engine", "index", "io", "model", "notes", "observations", "parser", "refs",
    "report", "temporal"].map((m) => `consolidador/${m}.py`),
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
    // no-cache: o navegador sempre confere com o site se o arquivo mudou (a tabela de links é
    // trocada com frequência); quando não mudou, usa o que guardou, sem baixar de novo
    const resposta = await fetch(`${base}py/${nome}?v=${VERSAO}`, { cache: "no-cache" });
    if (!resposta.ok) throw new Error(`não encontrei py/${nome} no site (erro ${resposta.status})`);
    if (nome.includes("/")) py.FS.mkdirTree(`/formatador/${nome.slice(0, nome.lastIndexOf("/"))}`);
    py.FS.writeFile(`/formatador/${nome}`, new Uint8Array(await resposta.arrayBuffer()));
  }
  py.runPython("import sys; sys.path.insert(0, '/formatador')");
  ponte = py.pyimport("web_ponte");
  py.FS.mkdirTree("/work/saida");
  py.FS.mkdirTree("/work/previa");
  const info = JSON.parse(ponte.versao());
  info.consolidador = py.runPython("import consolidador; consolidador.__version__");
  return info;
}

// O Consolidador só é importado quando a aba Aplicar alterações é usada
let consPonte = null;
function cons() {
  if (!consPonte) consPonte = py.pyimport("consolidador_ponte");
  return consPonte;
}

function limparPasta(pasta) {
  try {
    for (const f of py.FS.readdir(pasta)) if (f !== "." && f !== "..") py.FS.unlink(`${pasta}/${f}`);
  } catch {
    py.FS.mkdirTree(pasta);
  }
}

// Grava a legislação e o ato alterador com os nomes originais (o relatório e o plano mostram esses nomes)
function gravarEntradasConsolidar({ base, baseNome, alterador, alteradorNome }) {
  const caminhos = {};
  for (const [papel, dados, nome] of [["base", base, baseNome], ["alterador", alterador, alteradorNome]]) {
    if (!dados) continue;
    const pasta = `/work/consolidar/${papel}`;
    limparPasta(pasta);
    caminhos[papel] = `${pasta}/${nomeArquivo(nome, papel === "base" ? "legislacao.docx" : "ato_alterador.docx")}`;
    py.FS.writeFile(caminhos[papel], new Uint8Array(dados));
  }
  return caminhos;
}

function caminhoAtual(papel) {
  const pasta = `/work/consolidar/${papel}`;
  try {
    const f = py.FS.readdir(pasta).find((x) => x !== "." && x !== "..");
    return f ? `${pasta}/${f}` : "";
  } catch {
    return "";
  }
}

function camposConsolidar(d) {
  return [d.ato || "", d.publicacao || "", d.efeitos || "", d.anoNotas || "auto"];
}

const pronto = iniciar();

function json(texto) {
  return JSON.parse(texto);
}

function nomeArquivo(nome, padrao) {
  const base = String(nome || "").split(/[\\/]/).pop().replace(/[\x00-\x1f]/g, "").trim();
  return !base || base === "." || base === ".." ? padrao : base;
}

function nomeSeguro(nome) {
  const base = nomeArquivo(nome, "entrada.pdf");
  return /\.pdf$/i.test(base) ? base : `${base}.pdf`;
}

function lerSeExistir(caminho) {
  try {
    return py.FS.readFile(caminho);
  } catch {
    return null;
  }
}

const acoes = {
  /* ---- Consolidador (aba Aplicar alterações) ---- */
  consAnalisar(d) {
    const c = gravarEntradasConsolidar(d);
    return json(cons().analisar(c.base, c.alterador, ...camposConsolidar(d)));
  },
  consAlternar({ id }) {
    return json(cons().alternar(id));
  },
  consCorrigir({ id, acao, corpo, destino, alcance, efeitos, renumerar, nova }) {
    return json(cons().corrigir(id, acao, corpo, destino, alcance, efeitos, renumerar, nova));
  },
  consAdicionar() {
    return json(cons().adicionar());
  },
  consRemover({ id }) {
    return json(cons().remover(id));
  },
  consAprovar() {
    return json(cons().aprovar_extracao());
  },
  consGerar(d) {
    const destino = `/work/consolidar/saida/${nomeArquivo(d.nome, "legislacao_CONSOLIDADO.docx")}`;
    const r = json(cons().gerar(destino, !!d.realcar, ...camposConsolidar(d)));
    if (r.ok) r.wordBytes = py.FS.readFile(r.word);
    return r;
  },
  consRelatorio(d) {
    const r = json(cons().relatorio("/work/consolidar/saida/plano.html", ...camposConsolidar(d)));
    if (r.ok) r.htmlBytes = py.FS.readFile(r.relatorio);
    return r;
  },
  consSalvarPlano(d) {
    const r = json(cons().salvar_plano("/work/consolidar/saida/Plano_consolidacao.json", ...camposConsolidar(d)));
    if (r.ok) r.jsonBytes = py.FS.readFile(r.plano);
    return r;
  },
  consAbrirPlano(d) {
    gravarEntradasConsolidar(d);
    py.FS.mkdirTree("/work/consolidar/plano");
    py.FS.writeFile("/work/consolidar/plano/plano.json", new Uint8Array(d.plano));
    return json(cons().abrir_plano("/work/consolidar/plano/plano.json", caminhoAtual("base"), caminhoAtual("alterador")));
  },

  async versao() {
    return await pronto;
  },
  analisar({ pdf, nome, consulta, maxPaginas, ocorrencia }) {
    // o PDF fica com o nome original, como no programa do computador: a data no nome entra na
    // linha de publicação ("2026.10.14.DOE.pdf"), e "Extra" e "Previsão" no nome geram os avisos
    const caminho = `/work/pdf/${nomeSeguro(nome)}`;
    try { py.FS.readdir("/work/pdf").filter((f) => f !== "." && f !== "..").forEach((f) => py.FS.unlink(`/work/pdf/${f}`)); }
    catch { py.FS.mkdirTree("/work/pdf"); }
    py.FS.writeFile(caminho, new Uint8Array(pdf));
    return json(ponte.analisar(caminho, consulta, maxPaginas, ocorrencia || 0));
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
