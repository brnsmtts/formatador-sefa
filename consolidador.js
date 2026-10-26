/* Aba "Aplicar alterações": o Consolidador SEFA na web.
   A análise, a validação e a geração do Word são as do programa do computador (pasta py/consolidador),
   chamadas pelo worker.js por meio do py/consolidador_ponte.py. Usa as funções de app.js
   ($, h, executar, escolherDestino, gravar, salvarArquivo, confirmar, caixaErro, baixar). */
"use strict";

const consEst = {
  base: null,          // File: legislação a ser alterada
  alterador: null,     // File: ato alterador
  analisados: null,    // {base, alterador} usados na análise (ou no plano aberto)
  estado: null,        // última resposta do Python
  ativo: "",           // id da proposta em revisão
  sujo: false,         // campos da proposta editados e ainda não validados
  carregando: false,
  aba: "antes",
  versao: "",
};

const EXT_BASE = /\.(docx|doc)$/i;
const EXT_ALTERADOR = /\.(docx|doc|pdf|txt|html?)$/i;
const MSG_DOC = "O navegador não consegue converter arquivos .doc (formato antigo do Word). Abra o arquivo no Word, " +
  "use Salvar como > Documento do Word (.docx) e escolha a cópia. O original fica como está.";

/* ---------------- abas do topo ---------------- */
const MODOS = {
  formatar: { aba: "aba-formatar", painel: "modo-formatar", sub: "Importar ato do PDF do diário oficial",
    titulo: "Formatador SEFA · importar ato do PDF" },
  consolidar: { aba: "aba-consolidar", painel: "modo-consolidar", sub: "Inserir na legislação as alterações de um ato",
    titulo: "Formatador SEFA · aplicar alterações" },
};
let rodapeFormatar = "";

function mudarModo(modo, foco = false) {
  if (!MODOS[modo]) modo = "formatar";
  if (document.body.dataset.modo === modo) return;
  if (document.body.dataset.modo === "consolidar") ampliarRevisao(false);
  document.body.dataset.modo = modo;
  for (const [nome, m] of Object.entries(MODOS)) {
    const ativa = nome === modo;
    $(m.aba).setAttribute("aria-selected", ativa ? "true" : "false");
    $(m.aba).tabIndex = ativa ? 0 : -1;
    $(m.painel).hidden = !ativa;
  }
  $("sub-modo").textContent = MODOS[modo].sub;
  document.title = MODOS[modo].titulo;
  atualizarRodape();
  try { localStorage.setItem("formatador-sefa:modo", modo); } catch { /* sem armazenamento */ }
  const hash = modo === "consolidar" ? "#alteracoes" : "";
  if (location.hash !== hash) history.replaceState(null, "", location.pathname + location.search + hash);
  if (foco) $(MODOS[modo].aba).focus();
}

function atualizarRodape() {
  if (document.body.dataset.modo === "consolidar") {
    if (!rodapeFormatar) rodapeFormatar = $("rodape").textContent;
    $("rodape").textContent = consEst.versao
      ? `Consolidador SEFA ${consEst.versao}, versão web. Os arquivos são processados neste computador; os originais não são alterados.`
      : "Consolidador SEFA, versão web.";
  } else if (rodapeFormatar) {
    $("rodape").textContent = rodapeFormatar;
  }
}

function aoProntoConsolidador(info) {
  rodapeFormatar = $("rodape").textContent;
  consEst.versao = info.consolidador || "";
  atualizarRodape();
}

for (const [nome, m] of Object.entries(MODOS)) {
  $(m.aba).addEventListener("click", () => mudarModo(nome));
  $(m.aba).addEventListener("keydown", (e) => {
    if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
      e.preventDefault();
      mudarModo(nome === "formatar" ? "consolidar" : "formatar", true);
    }
  });
}
window.addEventListener("hashchange", () => mudarModo(location.hash === "#alteracoes" ? "consolidar" : "formatar"));

/* ---------------- arquivos ---------------- */
function tamanho(bytes) {
  return bytes > 1048576 ? `${(bytes / 1048576).toFixed(1).replace(".", ",")} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

function erroAnalise(titulo, texto) {
  $("cons-analisar-erro").replaceChildren(caixaErro(titulo, texto));
}

function escolherArquivoCons(papel, arquivo) {
  if (!arquivo) return;
  const ext = papel === "base" ? EXT_BASE : EXT_ALTERADOR;
  if (!ext.test(arquivo.name)) {
    erroAnalise("Arquivo não aceito", papel === "base"
      ? "A legislação a ser alterada deve estar em Word (.docx)."
      : "Use o ato alterador em Word (.docx), PDF com texto, TXT ou HTML.");
    return;
  }
  if (/\.doc$/i.test(arquivo.name)) {
    erroAnalise(`${arquivo.name}: arquivo .doc`, MSG_DOC);
    return;
  }
  $("cons-analisar-erro").replaceChildren();
  consEst[papel] = arquivo;
  $(`cons-${papel}-nome`).textContent = arquivo.name;
  $(`cons-${papel}-tam`).textContent = tamanho(arquivo.size);
  $(`cons-${papel}-escolhido`).hidden = false;
  $(`cons-soltar-${papel}`).hidden = true;
  if (papel === "alterador" && consEst.analisados && consEst.analisados.alterador !== arquivo) {
    // trocar o alterador não pode reaproveitar o número e as datas da análise anterior (como no programa)
    $("cons-ato").value = "";
    $("cons-publicacao").value = "";
    $("cons-efeitos").value = "";
  }
  atualizarConsolidador();
}

for (const papel of ["base", "alterador"]) {
  const zona = $(`cons-soltar-${papel}`);
  const entrada = $(`cons-arquivo-${papel}`);
  entrada.addEventListener("change", () => { escolherArquivoCons(papel, entrada.files[0]); entrada.value = ""; });
  ["dragenter", "dragover"].forEach((t) => zona.addEventListener(t, (e) => { e.preventDefault(); zona.classList.add("sobre"); }));
  ["dragleave", "drop"].forEach((t) => zona.addEventListener(t, () => zona.classList.remove("sobre")));
  zona.addEventListener("drop", (e) => {
    e.preventDefault();
    e.stopPropagation();
    escolherArquivoCons(papel, e.dataTransfer.files[0]);
  });
  $(`cons-trocar-${papel}`).addEventListener("click", () => entrada.click());
}

function arquivosMudaram() {
  const a = consEst.analisados;
  return !!(a && (a.base !== consEst.base || a.alterador !== consEst.alterador));
}

function campos() {
  return { ato: $("cons-ato").value, publicacao: $("cons-publicacao").value, efeitos: $("cons-efeitos").value,
    anoNotas: $("cons-ano-notas").value };
}

/* ---------------- análise ---------------- */
async function lerArquivos() {
  const base = await consEst.base.arrayBuffer();
  const alterador = await consEst.alterador.arrayBuffer();
  return { base, baseNome: consEst.base.name, alterador, alteradorNome: consEst.alterador.name };
}

async function analisarCons() {
  if (!consEst.base || !consEst.alterador) {
    erroAnalise("Selecione os arquivos", "Escolha a legislação base e o ato alterador.");
    return;
  }
  $("cons-analisar-erro").replaceChildren();
  $("cons-saida").replaceChildren();
  definirStatus("Lendo os arquivos e localizando os dispositivos. Aguarde a análise.");
  const arquivos = await lerArquivos();
  let r;
  try {
    r = await executar("consAnalisar", { ...arquivos, ...campos() }, [arquivos.base, arquivos.alterador]);
  } catch (e) {
    r = { ok: false, erro: e.message };
  }
  if (!r.ok) {
    definirStatus("A operação não foi concluída. Os arquivos de entrada foram preservados.");
    erroAnalise("Não foi possível concluir a análise", r.erro + "\nOs arquivos de entrada foram preservados.");
    return;
  }
  consEst.analisados = { base: consEst.base, alterador: consEst.alterador };
  receberAnalise(r, "Análise concluída. Confira cada proposta. Marque ou desmarque na coluna Aplicar as mudanças que vão para o Word.");
}

function receberAnalise(r, status) {
  // como na janela do programa: os campos passam a mostrar o que foi lido do ato
  $("cons-ato").value = r.ato || "";
  $("cons-publicacao").value = r.publicacao || "";
  $("cons-efeitos").value = r.efeitos || "";
  $("cons-ano-notas").value = r.ano_notas || "auto";
  $("cons-corpos").replaceChildren(...(r.corpos || []).map((c) => h("option", { value: c })));
  $("cons-destinos").replaceChildren(...[...new Set(r.destinos || [])].map((d) => h("option", { value: d })));
  consEst.estado = r;
  consEst.ativo = r.ativo || "";
  consEst.sujo = false;
  consEst.aba = "antes";
  $("cons-vazio").hidden = true;
  $("cons-trabalho").hidden = false;
  $("cons-aviso-texto").textContent = r.resumo_analise || "";
  $("cons-aviso").classList.toggle("pendente", r.pendentes > 0);
  $("cons-aviso").hidden = !r.resumo_analise;
  definirStatus(status);
  desenharCons();
  $("cons-trabalho").scrollIntoView({ behavior: "smooth", block: "start" });
}

function definirStatus(texto) {
  $("cons-status").textContent = texto || "";
}

/* ---------------- lista e revisão ---------------- */
function proposta(id) {
  return consEst.estado ? consEst.estado.operacoes.find((o) => o.id === id) || null : null;
}

function classeSituacao(s) {
  return { "pronto": "pronto", "revisar": "revisar", "já consta": "ja-consta", "outro ato": "outro-ato" }[s] || "outro-ato";
}

function desenharCons() {
  const r = consEst.estado;
  if (!r) return atualizarConsolidador();
  $("cons-resumo").textContent = r.resumo;
  if (!proposta(consEst.ativo)) consEst.ativo = r.operacoes.length ? r.operacoes[0].id : "";
  $("cons-linhas").replaceChildren(...r.operacoes.map((op) => {
    const marca = h("input", { type: "checkbox", "aria-label": `Aplicar ${op.id}`,
      title: op.situacao === "pronto" ? "Aplicar esta mudança no Word" : "Essa proposta precisa de correção antes de ser aplicada" });
    marca.checked = op.selecionada;
    marca.addEventListener("click", (e) => { e.stopPropagation(); alternarCons(op.id, marca); });
    const linha = h("tr", { tabindex: "0", "aria-selected": op.id === consEst.ativo ? "true" : "false", dataset: { id: op.id } },
      h("td", {}, marca),
      h("td", {}, op.id),
      h("td", {}, h("span", { class: `situacao ${classeSituacao(op.situacao)}` }, op.aplicada ? "aplicada" : op.situacao)),
      h("td", {}, op.acao),
      h("td", { class: "destino" }, op.destino || "Destino pendente"),
      h("td", { class: "efeitos" }, op.efeitos),
      h("td", { class: "origem" }, op.origem));
    linha.addEventListener("click", () => selecionarCons(op.id));
    linha.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); selecionarCons(op.id); }
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        const outra = e.key === "ArrowDown" ? linha.nextElementSibling : linha.previousElementSibling;
        if (outra) { selecionarCons(outra.dataset.id); outra.focus(); }
      }
    });
    return linha;
  }));
  $("cons-obs").textContent = r.observacoes;
  preencherRevisao();
  atualizarConsolidador();
}

function preencherRevisao() {
  const op = proposta(consEst.ativo);
  consEst.carregando = true;
  $("cons-rev-erro").hidden = true;
  if (!op) {
    $("cons-rev-titulo").textContent = "Nenhuma proposta selecionada";
    $("cons-rev-motivo").textContent = "Analise os arquivos e clique em uma proposta na lista acima.";
    $("cons-rev-dica").textContent = "";
    $("cons-antes").textContent = "Nenhuma proposta selecionada. Analise os arquivos e clique em uma proposta na lista acima.";
    $("cons-nova").value = "";
    $("cons-trecho").textContent = "O trecho do ato alterador aparecerá aqui quando uma proposta for selecionada.";
    $("cons-notas").textContent = "As datas identificadas e as remissões aparecerão aqui quando uma proposta for selecionada.";
  } else {
    $("cons-rev-titulo").textContent = `${op.id} · ${op.situacao.toUpperCase()} · ${op.destino || "destino pendente"}`;
    $("cons-rev-motivo").textContent = op.motivo;
    $("cons-rev-dica").textContent = op.dica;
    $("cons-acao").value = op.acao_rotulo;
    $("cons-corpo").value = op.corpo;
    $("cons-destino").value = op.destino;
    $("cons-alcance").value = op.alcance;
    $("cons-efeitos-op").value = op.efeitos_campo;
    $("cons-renumerar").value = op.renumerar;
    $("cons-antes").textContent = op.antes;
    $("cons-nova").value = op.nova;
    $("cons-trecho").textContent = op.trecho;
    $("cons-notas").textContent = op.notas;
    // o que precisa de correção já abre os campos
    $("cons-corrigir").open = op.situacao === "revisar" || !op.destino;
  }
  marcarSujo(false);
  mostrarAbaCons(consEst.aba);
  consEst.carregando = false;
}

function mostrarAbaCons(aba) {
  consEst.aba = aba;
  for (const b of document.querySelectorAll(".cons-abas .aba")) {
    const ativa = b.dataset.aba === aba;
    b.setAttribute("aria-selected", ativa ? "true" : "false");
    b.tabIndex = ativa ? 0 : -1;
  }
  for (const p of document.querySelectorAll(".cons-painel [data-painel]")) p.hidden = p.dataset.painel !== aba;
}
for (const b of document.querySelectorAll(".cons-abas .aba")) {
  b.addEventListener("click", () => mostrarAbaCons(b.dataset.aba));
  b.addEventListener("keydown", (e) => {
    if (e.key !== "ArrowRight" && e.key !== "ArrowLeft") return;
    e.preventDefault();
    const abas = [...document.querySelectorAll(".cons-abas .aba")];
    const i = (abas.indexOf(b) + (e.key === "ArrowRight" ? 1 : -1) + abas.length) % abas.length;
    mostrarAbaCons(abas[i].dataset.aba);
    abas[i].focus();
  });
}

function marcarSujo(valor) {
  consEst.sujo = valor;
  $("cons-sujo").hidden = !valor;
}
for (const id of ["cons-acao", "cons-corpo", "cons-destino", "cons-alcance", "cons-efeitos-op", "cons-renumerar", "cons-nova"]) {
  $(id).addEventListener("input", () => { if (!consEst.carregando && consEst.ativo) marcarSujo(true); });
}

function selecionarCons(id) {
  if (id === consEst.ativo) return;
  if (consEst.sujo && consEst.ativo) {
    definirStatus("Os campos editados da proposta anterior não foram validados. A proposta salva permanece no plano.");
  }
  consEst.ativo = id;
  for (const tr of $("cons-linhas").children) tr.setAttribute("aria-selected", tr.dataset.id === id ? "true" : "false");
  preencherRevisao();
  atualizarConsolidador();
}

function receberEstado(r, aviso = true) {
  if (!r.ok) {
    mostrarErroRevisao(r.erro);
    return false;
  }
  // o resumo da análise vale para o momento da análise; depois de uma mudança, o resumo da lista é o atual
  $("cons-aviso").hidden = true;
  consEst.estado = r;
  if (r.ativo) consEst.ativo = r.ativo;
  if (aviso && r.mensagem) definirStatus(r.mensagem);
  desenharCons();
  return true;
}

function mostrarErroRevisao(texto) {
  $("cons-rev-erro").textContent = texto;
  $("cons-rev-erro").hidden = false;
}

async function chamar(acao, dados = {}) {
  try {
    return await executar(acao, dados);
  } catch (e) {
    return { ok: false, erro: e.message };
  }
}

async function alternarCons(id, marca) {
  const op = proposta(id);
  if (!op) return;
  marca.checked = op.selecionada;     // a marca só muda com a resposta do programa
  if (consEst.ativo !== id) selecionarCons(id);
  const r = await chamar("consAlternar", { id });
  if (receberEstado(r) && !r.mensagem) definirStatus("");
}

async function validarCons() {
  const op = proposta(consEst.ativo);
  if (!op) return false;
  $("cons-rev-erro").hidden = true;
  const r = await chamar("consCorrigir", {
    id: op.id, acao: $("cons-acao").value, corpo: $("cons-corpo").value, destino: $("cons-destino").value,
    alcance: $("cons-alcance").value, efeitos: $("cons-efeitos-op").value, renumerar: $("cons-renumerar").value,
    nova: $("cons-nova").value,
  });
  if (!r.ok) {
    mostrarErroRevisao(`Confira os campos: ${r.erro}`);
    return false;
  }
  consEst.sujo = false;
  receberEstado(r);
  return true;
}

async function adicionarCons() {
  const r = await chamar("consAdicionar");
  if (receberEstado(r)) {
    $("cons-corrigir").open = true;
    mostrarAbaCons("nova");
    $("cons-destino").focus();
  }
}

async function removerCons() {
  const op = proposta(consEst.ativo);
  if (!op) return;
  const ok = await confirmar(`Remover a proposta ${op.id} (${op.destino || "destino pendente"}) do plano?`, "Remover", "Manter");
  if (!ok) return;
  consEst.sujo = false;
  consEst.ativo = "";
  receberEstado(await chamar("consRemover", { id: op.id }));
}

async function aprovarCons() {
  const n = consEst.estado ? consEst.estado.pendentes_pdf : 0;
  if (!n) {
    definirStatus("Não há propostas de PDF/HTML aguardando conferência.");
    return;
  }
  const ok = await confirmar(`Você conferiu as ${n} propostas extraídas do PDF/HTML com o texto de origem?`, "Sim, conferi", "Ainda não");
  if (!ok) return;
  receberEstado(await chamar("consAprovar"));
}

/* ---------------- ampliar ---------------- */
function ampliarRevisao(ligar) {
  const painel = $("cons-revisao");
  ligar = ligar === undefined ? !painel.classList.contains("ampliada") : ligar;
  painel.classList.toggle("ampliada", ligar);
  document.body.classList.toggle("cons-ampliada", ligar);
  $("cons-ampliar").textContent = ligar ? "Voltar" : "Ampliar";
  $("cons-ampliar").setAttribute("aria-pressed", ligar ? "true" : "false");
  if (ligar) painel.scrollTop = 0;
}
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && $("cons-revisao").classList.contains("ampliada") && !document.querySelector("dialog[open]")) {
    ampliarRevisao(false);
    $("cons-ampliar").focus();
  }
  if (e.key === "Enter" && e.ctrlKey && document.body.dataset.modo === "consolidar" && consEst.ativo
      && $("cons-revisao").contains(document.activeElement)) {
    e.preventDefault();
    validarCons();
  }
});

/* ---------------- gerar, relatório e plano ---------------- */
function conferirAntes() {
  // Application.sync da janela do programa (a parte que depende da página)
  if (consEst.sujo) return "Há campos editados na proposta atual. Clique em Validar correção antes de continuar.";
  if (!consEst.estado) return "Faça a análise primeiro.";
  if (arquivosMudaram()) return "A seleção de arquivos mudou. Clique em Analisar alterações novamente.";
  return "";
}

function mostrarSaida(...filhos) {
  $("cons-saida").replaceChildren(...filhos);
  $("cons-saida").scrollIntoView({ behavior: "smooth", block: "nearest" });
}

async function gerarCons() {
  const problema = conferirAntes();
  if (problema) return mostrarSaida(caixaErro("Gerar Word", problema));
  const nome = consEst.estado.nome_saida;
  const destino = await escolherDestino(nome, "docx");      // a janela Salvar como abre ainda dentro do clique
  if (!destino) return;
  definirStatus("Gerando o Word com as mudanças selecionadas.");
  let nomeSaida = destino.nome || nome;
  if (!/\.docx$/i.test(nomeSaida)) nomeSaida += ".docx";
  const r = await chamar("consGerar", { nome: nomeSaida, realcar: $("cons-realcar").checked, ...campos() });
  if (!r.ok) {
    definirStatus("A operação não foi concluída. Os arquivos de entrada foram preservados.");
    return mostrarSaida(caixaErro("Não foi possível gerar o Word", r.erro));
  }
  const bytes = r.wordBytes;
  receberEstado(r, false);
  let salvoComo;
  try {
    salvoComo = await gravar(destino, bytes, TIPO_DOCX);
  } catch (e) {
    return mostrarSaida(caixaErro("Não foi possível salvar o Word", String(e && e.message ? e.message : e)));
  }
  definirStatus(`Word salvo: ${salvoComo}. ${r.pendentes_fora} pendências ficaram fora do documento.`);
  mostrarSaida(h("div", { class: "resultado cons-resultado", role: "status" },
    h("h3", {}, destino.handle ? "Word atualizado salvo" : "Word atualizado baixado"),
    h("p", {}, salvoComo),
    h("p", {}, `Foram aplicadas ${r.aplicadas_agora} ${r.aplicadas_agora === 1 ? "mudança" : "mudanças"}.` +
      (r.pendentes_fora ? ` ${r.pendentes_fora} ${r.pendentes_fora === 1 ? "pendência ficou" : "pendências ficaram"} fora do documento.` : "")),
    h("p", {}, "Confira o documento final no Word."),
    h("div", { class: "botoes" },
      h("button", { class: "botao", type: "button", onclick: () => salvarArquivo(bytes, salvoComo, "docx", TIPO_DOCX) }, "Salvar de novo"))));
}

async function relatorioCons() {
  const problema = conferirAntes();
  if (problema) return mostrarSaida(caixaErro("Relatório", problema));
  // a aba é aberta ainda dentro do clique, para o navegador não bloquear
  const janela = window.open("", "_blank");
  if (janela) janela.document.write("<p style='font:16px system-ui;margin:2em'>Montando o relatório…</p>");
  const r = await chamar("consRelatorio", campos());
  if (!r.ok) {
    if (janela) janela.close();
    return mostrarSaida(caixaErro("Relatório", r.erro));
  }
  const blob = new Blob([r.htmlBytes], { type: "text/html;charset=utf-8" });
  if (janela) {
    janela.location.href = URL.createObjectURL(blob);
  } else {
    baixar(r.htmlBytes, "Plano_consolidacao.html", "text/html");
  }
}

async function salvarPlanoCons() {
  const problema = conferirAntes();
  if (problema) return mostrarSaida(caixaErro("Salvar plano", problema));
  const destino = await escolherDestino("Plano_consolidacao.json", "json");
  if (!destino) return;
  const r = await chamar("consSalvarPlano", campos());
  if (!r.ok) return mostrarSaida(caixaErro("Salvar plano", r.erro));
  await gravar(destino, r.jsonBytes, "application/json");
  definirStatus("Plano salvo. Para continuar depois, abra-o com a mesma legislação e o mesmo ato alterador.");
}

async function abrirPlanoCons(arquivo) {
  if (!arquivo) return;
  if (!consEst.base || !consEst.alterador) {
    erroAnalise("Abrir plano salvo", "Escolha antes, nos passos 1 e 2, a mesma legislação e o mesmo ato alterador usados no plano. " +
      "O plano guarda as propostas e as correções, mas não os arquivos.");
    return;
  }
  $("cons-analisar-erro").replaceChildren();
  const arquivos = await lerArquivos();
  const plano = await arquivo.arrayBuffer();
  const r = await chamar("consAbrirPlano", { ...arquivos, plano });
  if (!r.ok) return erroAnalise("Abrir plano salvo", r.erro);
  consEst.analisados = { base: consEst.base, alterador: consEst.alterador };
  receberAnalise(r, `Plano ${arquivo.name} aberto. Confira as propostas antes de gerar o Word.`);
}

async function exemploCons() {
  if (est.ocupado || !est.motorPronto) return;
  try {
    const nomes = ["01_Legislacao_base_FICTICIA.docx", "02_Ato_alterador_FICTICIO.docx"];
    const [b, a] = await Promise.all(nomes.map(async (n) => {
      const resp = await fetch(`exemplos/consolidador/${n}`);
      if (!resp.ok) throw new Error(`não encontrei exemplos/consolidador/${n} no site`);
      return new File([await resp.blob()], n, { type: TIPO_DOCX });
    }));
    escolherArquivoCons("base", b);
    escolherArquivoCons("alterador", a);
    $("cons-ato").value = "Lei 100.000/2099";
    $("cons-publicacao").value = "22/10/2099";
    $("cons-efeitos").value = "01/11/2099";
    await analisarCons();
  } catch (e) {
    erroAnalise("Exemplo de uso", String(e && e.message ? e.message : e));
  }
}

/* ---------------- botões ---------------- */
$("cons-analisar").addEventListener("click", analisarCons);
$("cons-exemplo").addEventListener("click", exemploCons);
$("cons-exemplo-2").addEventListener("click", exemploCons);
$("cons-abrir-plano").addEventListener("click", () => $("cons-arquivo-plano").click());
$("cons-arquivo-plano").addEventListener("change", () => {
  const f = $("cons-arquivo-plano").files[0];
  $("cons-arquivo-plano").value = "";
  abrirPlanoCons(f);
});
$("cons-validar").addEventListener("click", async () => {
  if (await validarCons() && $("cons-revisao").classList.contains("ampliada")) ampliarRevisao(false);
});
$("cons-ampliar").addEventListener("click", () => ampliarRevisao());
$("cons-adicionar").addEventListener("click", adicionarCons);
$("cons-remover").addEventListener("click", removerCons);
$("cons-aprovar").addEventListener("click", aprovarCons);
$("cons-gerar").addEventListener("click", gerarCons);
$("cons-relatorio").addEventListener("click", relatorioCons);
$("cons-salvar-plano").addEventListener("click", salvarPlanoCons);
$("cons-aviso-ok").addEventListener("click", () => { $("cons-aviso").hidden = true; });

function atualizarConsolidador() {
  const livre = est.motorPronto && !est.ocupado;
  const r = consEst.estado;
  $("cons-analisar").disabled = !livre || !consEst.base || !consEst.alterador;
  $("cons-analisar").textContent = est.ocupado && document.body.dataset.modo === "consolidar" ? "Processando…" : "Analisar alterações";
  for (const id of ["cons-exemplo", "cons-exemplo-2", "cons-abrir-plano"]) $(id).disabled = !livre;
  $("cons-passo-gerar").setAttribute("aria-disabled", r ? "false" : "true");
  const op = proposta(consEst.ativo);
  for (const id of ["cons-validar", "cons-ampliar", "cons-remover"]) $(id).disabled = !livre || !op;
  $("cons-adicionar").disabled = !livre || !r;
  $("cons-aprovar").disabled = !livre || !r || !r.pendentes_pdf;
  $("cons-aprovar").textContent = r && r.pendentes_pdf ? `Aprovar extração do PDF/HTML (${r.pendentes_pdf})` : "Aprovar extração do PDF/HTML";
  const motivos = [];
  if (r) {
    if (arquivosMudaram()) motivos.push("Analisar de novo (a legislação ou o ato alterador mudou)");
    if (!r.selecionadas) motivos.push("Selecionar pelo menos uma mudança pronta na coluna Aplicar");
  }
  $("cons-gerar").disabled = !livre || !r || motivos.length > 0;
  $("cons-relatorio").disabled = !livre || !r;
  $("cons-salvar-plano").disabled = !livre || !r;
  const lista = $("cons-motivos");
  if (r && motivos.length) {
    lista.replaceChildren(h("li", { style: "list-style:none;margin-left:-18px" }, "Para gerar o Word, falta:"), ...motivos.map((m) => h("li", {}, m)));
    lista.hidden = false;
  } else {
    lista.hidden = true;
  }
}

/* ---------------- início ---------------- */
(() => {
  let modo = location.hash === "#alteracoes" ? "consolidar" : "";
  if (!modo) {
    try { modo = localStorage.getItem("formatador-sefa:modo") || "formatar"; } catch { modo = "formatar"; }
  }
  document.body.dataset.modo = "";
  mudarModo(modo);
  atualizarConsolidador();
})();
