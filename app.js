/* Formatador SEFA na web: tela de importar ato do PDF.
   Toda a leitura do PDF e a geração do Word acontecem no worker.js (Python/Pyodide). */
"use strict";

const $ = (id) => document.getElementById(id);
const BLOCO_VAZIO = $("vazio");

function h(tag, attrs = {}, ...filhos) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === false || v == null) continue;
    if (k === "class") el.className = v;
    else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
    else if (k === "dataset") Object.assign(el.dataset, v);
    else el.setAttribute(k, v === true ? "" : v);
  }
  for (const f of filhos.flat()) {
    if (f == null || f === false) continue;
    el.append(f instanceof Node ? f : document.createTextNode(String(f)));
  }
  return el;
}

const est = {
  motorPronto: false,
  ocupado: false,
  pdf: null,          // File
  modelo: null,       // File (opcional)
  analise: null,      // último estado vindo do Python
  analisado: null,    // {pdf, consulta, maxPaginas} usados na análise
  aba: "texto",
  selPar: null,
  selCel: null,       // {t, r, c}
  info: null,
  ultimoResultado: null,
  avisoDesatualizado: false,
  versaoEstado: 0,    // muda a cada análise ou correção; a visualização do Word se refaz quando muda
  word: null,         // {chave, no} da última visualização montada
  pedidoWord: 0,
};

/* ---------------- links colados na busca assistida (guardados neste navegador) ---------------- */
const CHAVE_LINKS = "formatador-sefa:links-usuario";
function lerLinksUsuario() {
  try {
    const dados = JSON.parse(localStorage.getItem(CHAVE_LINKS) || "{}");
    return dados && typeof dados === "object" ? dados : {};
  } catch {
    return {};
  }
}
function gravarLinksUsuario(dados) {
  try {
    localStorage.setItem(CHAVE_LINKS, JSON.stringify(dados));
    return true;
  } catch {
    return false;
  }
}
function validarEndereco(texto) {
  const t = (texto || "").trim();
  if (!t) return { vazio: true };
  let u;
  try {
    u = new URL(t);
  } catch {
    return { ok: false, erro: "Isto não é um endereço de internet. Copie o endereço da barra do navegador." };
  }
  if (!/^https?:$/.test(u.protocol)) return { ok: false, erro: "O endereço precisa começar com https://." };
  if (u.hostname === "app.sefa.pa.gov.br" && !u.searchParams.get("idAto")) {
    return { ok: false, erro: "Este é o endereço da lista de resultados. Clique no ato e copie o endereço da página dele (ele tem idAto= no final)." };
  }
  return { ok: true, url: u.href };
}

/* ---------------- comunicação com o worker ---------------- */
if (location.protocol === "file:") {
  definirMotor("falha", "Abra pelo endereço do site ou pelo SERVIR_LOCAL.bat");
  BLOCO_VAZIO.prepend(h("div", { class: "erro-caixa" },
    h("strong", {}, "Esta página não funciona aberta direto do disco"),
    "O navegador bloqueia o carregamento do Python quando o arquivo é aberto com duplo clique. ",
    "Use o endereço do site publicado ou rode o SERVIR_LOCAL.bat e abra http://localhost:8000."));
  throw new Error("aberto via file://");
}

const worker = new Worker("worker.js");
let seq = 0;
const pendentes = new Map();

function pedir(acao, dados = {}, transferir = []) {
  return new Promise((resolver, rejeitar) => {
    const id = ++seq;
    pendentes.set(id, { resolver, rejeitar });
    worker.postMessage({ id, acao, dados }, transferir);
  });
}

worker.onmessage = ({ data }) => {
  if (data.tipo === "progresso") return definirMotor("carregando", data.texto);
  if (data.tipo === "pronto") {
    est.motorPronto = true;
    est.info = data.info;
    definirMotor("pronto", "Pronto");
    $("dica-links").textContent = `Usa a tabela de links do site (${data.info.tabela_links} atos) e os links que você salvou. ` +
      "As demais citações ficam em amarelo, e a página mostra antes de salvar as que ficaram sem link, para você procurá-las no site da SEFA.";
    $("rodape").textContent = `Formatador SEFA 3.18, versão web. PyMuPDF ${data.info.pymupdf}, ` +
      `pdfplumber ${data.info.pdfplumber}. O PDF é processado neste computador.`;
    return atualizar();
  }
  if (data.tipo === "falha") {
    definirMotor("falha", "O formatador não carregou");
    BLOCO_VAZIO.prepend(h("div", { class: "erro-caixa" },
      h("strong", {}, "Não foi possível carregar o formatador"),
      data.erro + "\nRecarregue a página. Se continuar, a rede pode estar bloqueando arquivos .wasm ou .whl do site."));
    return;
  }
  const p = pendentes.get(data.id);
  if (!p) return;
  pendentes.delete(data.id);
  data.ok ? p.resolver(data.resultado) : p.rejeitar(new Error(data.erro));
};
worker.onerror = (e) => {
  definirMotor("falha", "O formatador não carregou");
  console.error(e);
};

function definirMotor(estado, texto) {
  $("motor").dataset.estado = estado;
  $("motor-texto").textContent = texto;
}

async function executar(acao, dados, transferir) {
  est.ocupado = true;
  atualizar();
  const anterior = $("motor-texto").textContent;
  definirMotor("carregando", "Processando");
  try {
    return await pedir(acao, dados, transferir);
  } finally {
    est.ocupado = false;
    definirMotor(est.motorPronto ? "pronto" : "carregando", est.motorPronto ? "Pronto" : anterior);
    atualizar();
  }
}

/* ---------------- passo 1: PDF ---------------- */
function escolherPdf(arquivo) {
  if (!arquivo) return;
  if (!/\.pdf$/i.test(arquivo.name) && arquivo.type !== "application/pdf") {
    $("analisar-erro").replaceChildren(caixaErro("Arquivo não é PDF", `${arquivo.name} não parece ser um PDF do diário oficial.`));
    return;
  }
  est.pdf = arquivo;
  $("pdf-nome").textContent = arquivo.name;
  $("pdf-tam").textContent = `${(arquivo.size / 1048576).toFixed(1).replace(".", ",")} MB`;
  $("soltar").hidden = true;
  $("arquivo-escolhido").hidden = false;
  $("analisar-erro").replaceChildren();
  atualizar();
  $("consulta").focus();
}
$("arquivo-pdf").addEventListener("change", (e) => { escolherPdf(e.target.files[0]); e.target.value = ""; });
$("trocar-pdf").addEventListener("click", () => $("arquivo-pdf").click());
const soltar = $("soltar");
["dragenter", "dragover"].forEach((t) => soltar.addEventListener(t, (e) => { e.preventDefault(); soltar.classList.add("sobre"); }));
["dragleave", "drop"].forEach((t) => soltar.addEventListener(t, () => soltar.classList.remove("sobre")));
soltar.addEventListener("drop", (e) => { e.preventDefault(); escolherPdf(e.dataTransfer.files[0]); });
document.addEventListener("dragover", (e) => e.preventDefault());
document.addEventListener("drop", (e) => { e.preventDefault(); if (e.dataTransfer.files[0]) escolherPdf(e.dataTransfer.files[0]); });

/* ---------------- passo 2: analisar ---------------- */
$("consulta").addEventListener("input", atualizar);
$("max-paginas").addEventListener("input", atualizar);
$("consulta").addEventListener("keydown", (e) => { if (e.key === "Enter" && !$("analisar").disabled) analisar(); });
$("analisar").addEventListener("click", analisar);

function entradaAtual() {
  return { pdf: est.pdf, consulta: $("consulta").value.trim(), maxPaginas: $("max-paginas").value.trim() };
}
function analiseDesatualizada() {
  if (!est.analisado) return false;
  const a = entradaAtual();
  return a.pdf !== est.analisado.pdf || a.consulta !== est.analisado.consulta || a.maxPaginas !== est.analisado.maxPaginas;
}

async function analisar() {
  const entrada = entradaAtual();
  $("analisar-erro").replaceChildren();
  $("gerar-saida").replaceChildren();
  est.ultimoResultado = null;
  est.ocupado = true;   // já bloqueia o botão enquanto o arquivo é lido
  atualizar();
  const buffer = await entrada.pdf.arrayBuffer();
  let r;
  try {
    r = await executar("analisar", { pdf: buffer, consulta: entrada.consulta, maxPaginas: Number(entrada.maxPaginas) }, [buffer]);
  } catch (e) {
    r = { ok: false, erro: e.message };
  }
  if (!r.ok) {
    est.analise = null;
    est.analisado = null;
    $("analisar-erro").replaceChildren(caixaErro("Não foi possível analisar o PDF", r.erro));
    desenharPrevia();
    atualizar();
    return;
  }
  est.analise = r;
  est.analisado = entrada;
  est.versaoEstado += 1;
  est.word = null;
  est.selPar = null;
  est.selCel = null;
  est.aba = "texto";
  $("op-municipios").checked = r.municipios;
  desenharPrevia();
  atualizar();
  $("previa").scrollIntoView({ behavior: "smooth", block: "start" });
}

function caixaErro(titulo, texto) {
  return h("div", { class: "erro-caixa", role: "alert" }, h("strong", {}, titulo), texto);
}

/* ---------------- passo 3: conferência ---------------- */
function errosPorTabela(a) {
  const mapa = new Map();
  for (const e of a.erros_tabela) {
    const m = e.match(/^Tabela (\d+):\s*(.*)$/s);
    if (m) mapa.set(Number(m[1]) - 1, m[2]);
  }
  return mapa;
}

function pendenciasParaGerar(a) {
  const lista = [];
  for (const b of a.bloqueios) lista.push(`Bloqueio: ${b}`);
  const erros = errosPorTabela(a);
  for (const [t] of erros) lista.push(`Conferir a tabela ${t + 1} no PDF`);
  if (a.limites.length && !a.limites_conferidos) lista.push("Conferir o início e o fim do ato");
  return lista;
}

function desenharConferencia() {
  const lista = $("lista-conf");
  const a = est.analise;
  $("passo-conferir").setAttribute("aria-disabled", a ? "false" : "true");
  if (!a) {
    lista.replaceChildren(h("li", {}, "Analise o PDF para ver o que precisa de conferência."));
    return;
  }
  const itens = [];
  for (const b of a.bloqueios) {
    itens.push(h("li", { class: "bloqueio" }, h("strong", {}, "Exportação bloqueada"),
      h("span", { class: "det" }, b, " Será preciso reconstruir esse trecho manualmente no Word.")));
  }
  if (a.limites.length && !a.limites_conferidos) {
    itens.push(h("li", { class: "pendente" }, h("strong", {}, "Confira o início e o fim do ato"),
      h("span", { class: "det" }, a.limites.join("; ")),
      h("button", { class: "botao", type: "button", onclick: conferirLimites, disabled: est.ocupado }, "Conferi início e fim")));
  } else if (a.limites.length) {
    itens.push(h("li", { class: "ok" }, "Início e fim do ato conferidos"));
  } else {
    itens.push(h("li", {}, "Início e fim do ato sem alertas",
      h("span", { class: "det" }, `Páginas ${a.paginas.join(", ")} do PDF.`)));
  }
  const erros = errosPorTabela(a);
  if (!a.tabelas.length) itens.push(h("li", {}, "Nenhuma tabela neste ato"));
  for (const t of a.tabelas) {
    if (erros.has(t.i)) {
      itens.push(h("li", { class: "pendente" }, h("strong", {}, `Tabela ${t.i + 1} precisa de conferência`),
        h("span", { class: "det" }, erros.get(t.i)),
        h("button", { class: "botao", type: "button", onclick: () => irParaTabela(t.i) }, "Abrir a tabela")));
    } else {
      itens.push(h("li", { class: "ok" }, `Tabela ${t.i + 1} ${t.conferida ? "conferida por você" : "sem divergência com o PDF"}`));
    }
  }
  const atencao = a.paragrafos.filter((p) => p.atencao.length);
  if (atencao.length) {
    itens.push(h("li", { class: "pendente" },
      h("strong", {}, `${atencao.length} ${atencao.length === 1 ? "trecho marcado" : "trechos marcados"} em amarelo`),
      h("span", { class: "det" }, "Não impedem o Word, mas confira-os no PDF."),
      h("button", { class: "botao", type: "button", onclick: () => irParaParagrafo(atencao[0].i) }, "Ir para o primeiro")));
  }
  lista.replaceChildren(...itens);
}

function confirmar(texto, sim = "Sim, conferi", nao = "Ainda não") {
  return new Promise((resolver) => {
    const dlg = $("dlg-confirma");
    $("dlg-confirma-texto").textContent = texto;
    $("dlg-confirma-sim").textContent = sim;
    $("dlg-confirma-nao").textContent = nao;
    dlg.returnValue = "";
    dlg.addEventListener("close", () => resolver(dlg.returnValue === "sim"), { once: true });
    dlg.showModal();
    $("dlg-confirma-sim").focus();
  });
}

async function conferirLimites() {
  const ok = await confirmar("Você conferiu no PDF que o título e todas as páginas do ato foram incluídos, " +
    "até sua assinatura, protocolo ou início do ato seguinte?");
  if (!ok) return;
  aplicarEstado(await executar("conferirLimites"));
}

/* ---------------- passo 4: gerar ---------------- */
$("op-municipios").addEventListener("change", async (e) => {
  if (!est.analise) return;
  aplicarEstado(await executar("municipios", { valor: e.target.checked }));
});
$("modelo-trocar").addEventListener("click", () => $("arquivo-modelo").click());
$("arquivo-modelo").addEventListener("change", (e) => {
  const f = e.target.files[0];
  if (!f) return;
  est.modelo = f;
  $("modelo-nome").textContent = f.name;
  $("modelo-padrao").hidden = false;
  e.target.value = "";
  refazerWordSeVisivel();
});
$("modelo-padrao").addEventListener("click", () => {
  est.modelo = null;
  $("modelo-nome").textContent = "padrão da SEFA";
  $("modelo-padrao").hidden = true;
  refazerWordSeVisivel();
});
$("gerar").addEventListener("click", () => gerar("baixar"));
$("editar").addEventListener("click", () => gerar("editor"));
$("gerar-pdf").addEventListener("click", () => gerar("pdf"));
$("op-links").addEventListener("change", refazerWordSeVisivel);

function baixar(bytes, nome, tipo) {
  const url = URL.createObjectURL(new Blob([bytes], { type: tipo }));
  const a = h("a", { href: url, download: nome });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
const TIPO_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/* Salvar escolhendo pasta e nome: usa a janela "Salvar como" do Edge e do Chrome. Em navegadores sem
   esse recurso, cai no download comum. escolherDestino() precisa ser a primeira coisa do clique. */
const PODE_ESCOLHER_PASTA = typeof window.showSaveFilePicker === "function";
const TIPOS_ARQUIVO = {
  docx: { description: "Documento do Word", accept: { [TIPO_DOCX]: [".docx"] } },
  html: { description: "Relatório de conferência", accept: { "text/html": [".html"] } },
  tsv: { description: "Linhas da tabela de links", accept: { "text/tab-separated-values": [".tsv"] } },
};

async function escolherDestino(nome, tipo) {
  if (!PODE_ESCOLHER_PASTA) return { nome };
  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: nome, id: `formatador-${tipo}`, startIn: "documents", types: [TIPOS_ARQUIVO[tipo]],
    });
    return { handle, nome: handle.name };
  } catch (e) {
    if (e && e.name === "AbortError") return null;      // a pessoa cancelou
    return { nome };
  }
}

async function gravar(destino, dados, mime) {
  const blob = dados instanceof Blob ? dados : new Blob([dados], { type: mime });
  if (destino.handle) {
    const escrita = await destino.handle.createWritable();
    await escrita.write(blob);
    await escrita.close();
  } else {
    baixar(blob, destino.nome, mime);
  }
  return destino.nome;
}

async function salvarArquivo(dados, nome, tipo, mime) {
  const destino = await escolherDestino(nome, tipo);
  if (!destino) return null;
  try {
    return await gravar(destino, dados, mime);
  } catch (e) {
    alert(`Não foi possível salvar o arquivo: ${e && e.message ? e.message : e}`);
    return null;
  }
}

async function gerarNoWorker() {
  const a = est.analise;
  est.ocupado = true;
  atualizar();
  const modelo = est.modelo ? await est.modelo.arrayBuffer() : null;
  let r;
  try {
    r = await executar("gerar", {
      nome: a.nome_sugerido, modelo,
      relatorio: $("op-relatorio").checked, links: $("op-links").checked,
      linksUsuario: JSON.stringify(lerLinksUsuario()),
    }, modelo ? [modelo] : []);
  } catch (e) {
    r = { ok: false, erro: e.message };
  }
  if (!r.ok) {
    $("gerar-saida").replaceChildren(caixaErro("Não foi possível gerar o Word", r.erro));
    return null;
  }
  return r;
}

/* Gera o Word e confere os links ANTES de baixar: se alguma citação ficou sem link,
   mostra a busca assistida e só baixa depois, uma única vez. */
async function gerar(destino = "baixar") {
  est.destino = destino;
  est.alvo = null;
  if (destino === "baixar") {          // primeiro escolhe onde salvar, ainda dentro do clique
    const alvo = await escolherDestino(est.analise.nome_sugerido, "docx");
    if (!alvo) return;
    est.alvo = alvo;
  }
  $("gerar-saida").replaceChildren();
  const r = await gerarNoWorker();
  if (!r) return;
  if (!(r.pendentes_busca || []).length) return entregar(r);
  $("gerar-saida").replaceChildren(painelAntesDeBaixar(r));
  $("gerar-saida").scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function botoesDoResultado(r, nomeWord) {
  const nomeRel = nomeWord.replace(/\.docx$/i, "_CONFERIR.html");
  const botoes = [
    h("button", { class: "botao", type: "button", onclick: () => salvarArquivo(r.wordBytes, nomeWord, "docx", TIPO_DOCX) }, "Salvar o Word de novo"),
    h("button", { class: "botao", type: "button", onclick: () => pdfDoResultado(r) }, "Gerar o PDF"),
  ];
  if (r.relatorioBytes) {
    botoes.push(h("button", { class: "botao", type: "button", onclick: () => salvarArquivo(r.relatorioBytes, nomeRel, "html", "text/html") }, "Salvar o relatório"));
  }
  return botoes;
}

async function entregar(r) {
  if (est.destino === "editor") return entregarNoEditor(r);
  if (est.destino === "pdf") return entregarPdf(r);
  const nomeWord = est.analise.nome_sugerido;
  est.ultimoResultado = { ...r, nomeWord };
  const alvo = est.alvo || { nome: nomeWord };
  let salvoComo;
  try {
    salvoComo = await gravar(alvo, r.wordBytes, TIPO_DOCX);
  } catch (e) {
    $("gerar-saida").replaceChildren(caixaErro("Não foi possível salvar o Word", String(e && e.message ? e.message : e)));
    return;
  }
  const restantes = (r.pendentes_busca || []).map((p) => p.exemplo);
  $("gerar-saida").replaceChildren(...[h("div", { class: "resultado", role: "status" },
    h("h3", {}, alvo.handle ? "Word salvo" : "Word baixado"),
    h("p", {}, salvoComo),
    r.resumo_links ? h("p", {}, r.resumo_links) : null,
    restantes.length ? h("p", {}, `Ficaram em amarelo no Word: ${restantes.join("; ")}.`) : null,
    r.aviso_links ? h("p", { class: "msg-link" }, r.aviso_links) : null,
    h("p", {}, "Compare com o PDF antes de usar."),
    h("div", { class: "botoes" }, botoesDoResultado(r, nomeWord))),
    r.fora_do_site && r.fora_do_site.length ? h("p", { class: "dica" },
      `Sem link e sem destaque, porque esse tipo de ato não consta no site da SEFA: ${r.fora_do_site.join("; ")}.`) : null,
  ].filter(Boolean));
}

function painelAntesDeBaixar(r) {
  const pendentes = r.pendentes_busca;
  const itens = pendentes.map((p) => {
    const campo = h("input", { type: "url", inputmode: "url", placeholder: "Cole aqui o endereço da página do ato",
      "aria-label": `Endereço de ${p.exemplo}`, spellcheck: "false" });
    const msg = h("p", { class: "msg-link", hidden: true });
    const li = h("li", { class: "pend" },
      h("strong", {}, p.exemplo),
      h("span", { class: "det" }, `${p.qtd} ${p.qtd === 1 ? "ocorrência" : "ocorrências"} no ato`),
      h("a", { class: "botao", href: p.busca, target: "_blank", rel: "noopener" }, "Procurar no site da SEFA"),
      campo, msg);
    return { p, campo, msg, li };
  });
  const noEditor = est.destino === "editor";
  const fim = { editor: "abrir no editor", pdf: "gerar o PDF" }[est.destino] || "salvar";
  const aplicar = h("button", { class: "botao principal largo", type: "button", disabled: true }, `Aplicar os links e ${fim}`);
  const semLinks = h("button", { class: "botao largo", type: "button" },
    { editor: "Abrir no editor sem esses links", pdf: "Gerar o PDF sem esses links" }[est.destino] || "Salvar sem esses links");
  const conferir = () => {
    let validos = 0;
    let erros = 0;
    for (const it of itens) {
      const v = validarEndereco(it.campo.value);
      it.msg.hidden = !v.erro;
      it.msg.textContent = v.erro || "";
      it.campo.setAttribute("aria-invalid", v.erro ? "true" : "false");
      if (v.ok) validos += 1;
      if (v.erro) erros += 1;
    }
    aplicar.disabled = validos === 0 || erros > 0;
    aplicar.textContent = validos > 1 ? `Aplicar os ${validos} links e ${fim}` : `Aplicar o link e ${fim}`;
    if (validos === 0) aplicar.textContent = `Aplicar os links e ${fim}`;
  };
  itens.forEach((it) => it.campo.addEventListener("input", conferir));
  aplicar.addEventListener("click", async () => {
    aplicar.disabled = semLinks.disabled = true;
    const novos = {};
    for (const it of itens) {
      const v = validarEndereco(it.campo.value);
      if (v.ok) novos[it.p.chave] = v.url;
    }
    if (!gravarLinksUsuario({ ...lerLinksUsuario(), ...novos })) {
      alert("O navegador não deixou guardar os links. Eles serão usados só neste Word.");
    }
    desenharLinksSalvos();
    const r2 = await gerarNoWorker();
    if (r2) entregar(r2);
  });
  semLinks.addEventListener("click", () => entregar(r));
  return h("div", { class: "busca-assistida", role: "region", "aria-label": "Citações sem link" },
    h("h3", {}, `${{ editor: "Antes de abrir no editor", pdf: "Antes de gerar o PDF" }[est.destino] || "Antes de salvar"}: ${pendentes.length} ${pendentes.length === 1 ? "citação está" : "citações estão"} sem link`),
    h("p", { class: "dica" }, "A versão online não consulta o site da SEFA sozinha. Para incluir os links, clique em " +
      "Procurar no site da SEFA, abra o ato e cole aqui o endereço da página dele (o que tem idAto= no final). " +
      "Os links ficam guardados neste navegador e entram sozinhos nos próximos atos."),
    h("ol", {}, itens.map((it) => it.li)),
    aplicar, semLinks);
}

/* ---------------- links salvos neste navegador ---------------- */
function desenharLinksSalvos() {
  const dados = lerLinksUsuario();
  const chaves = Object.keys(dados).sort((a, b) => a.localeCompare(b, "pt-BR"));
  $("salvos").hidden = chaves.length === 0;
  $("salvos-n").textContent = chaves.length;
  $("salvos-lista").replaceChildren(...chaves.map((c) => h("li", {},
    h("a", { href: dados[c], target: "_blank", rel: "noopener" }, c.split("|").filter(Boolean).join(" ")),
    " ",
    h("button", { class: "ligacao", type: "button", onclick: () => {
      const atual = lerLinksUsuario();
      delete atual[c];
      gravarLinksUsuario(atual);
      desenharLinksSalvos();
      refazerWordSeVisivel();
    } }, "Remover"))));
}
$("salvos-exportar").addEventListener("click", () => {
  const dados = lerLinksUsuario();
  const linhas = Object.keys(dados).sort((a, b) => a.localeCompare(b, "pt-BR")).map((c) => `${c}\t${dados[c]}`);
  if (linhas.length) salvarArquivo(linhas.join("\r\n") + "\r\n", "links_novos.tsv", "tsv", "text/tab-separated-values;charset=utf-8");
});

/* ---------------- prévia ---------------- */
function aplicarEstado(r) {
  if (!r.ok) {
    alert(r.erro);
    return;
  }
  est.analise = r;
  est.versaoEstado += 1;
  $("gerar-saida").replaceChildren();
  desenharPrevia();
  atualizar();
}

function irParaParagrafo(i) {
  est.aba = "texto";
  est.selPar = i;
  desenharPrevia();
  const el = document.querySelector(`.par[data-i="${i}"]`);
  if (el) { el.scrollIntoView({ behavior: "smooth", block: "center" }); el.querySelector("textarea")?.focus(); }
}
function irParaTabela(t) {
  est.aba = "tabelas";
  desenharPrevia();
  document.getElementById(`tabela-${t}`)?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function desenharPrevia() {
  const previa = $("previa");
  const a = est.analise;
  est.avisoDesatualizado = analiseDesatualizada();
  if (!a) {
    previa.replaceChildren(BLOCO_VAZIO);
    return;
  }
  const abas = h("div", { class: "abas", role: "tablist" },
    h("button", { class: "aba", role: "tab", type: "button", "aria-selected": String(est.aba === "texto"),
      onclick: () => { est.aba = "texto"; desenharPrevia(); } }, `Texto do ato (${a.paragrafos.length})`),
    h("button", { class: "aba", role: "tab", type: "button", "aria-selected": String(est.aba === "tabelas"),
      onclick: () => { est.aba = "tabelas"; desenharPrevia(); } }, `Tabelas (${a.tabelas.length})`),
    h("button", { class: "aba aba-word", role: "tab", type: "button", "aria-selected": String(est.aba === "word"),
      title: "Mostra o Word como ele será baixado",
      onclick: () => { est.aba = "word"; desenharPrevia(); } }, iconeOlho(), "Ver como no Word"));
  const resumo = h("div", { class: "resumo-previa" },
    `${a.paginas.length > 1 ? "Páginas" : "Página"} ${a.paginas.join(", ")} do PDF` +
    (a.valores ? ` · ${a.valores} valores monetários lidos nas tabelas` : ""));
  const folha = h("div", { class: est.aba === "word" ? "folha vista" : "folha", role: "tabpanel" });
  if (est.avisoDesatualizado) {
    folha.append(h("div", { class: "aviso-folha bloqueio" },
      "Você mudou o PDF, o número do ato ou o limite de páginas. Clique em Analisar PDF de novo para atualizar a prévia."));
  }
  if (est.aba === "texto") desenharTexto(folha, a);
  else if (est.aba === "tabelas") desenharTabelas(folha, a);
  else desenharWord(folha);
  const desfazer = a.pode_desfazer ? h("button", { class: "botao desfazer", type: "button", onclick: desfazerUltima,
    title: "Desfaz a última correção ou conferência (Ctrl+Z)" }, "Desfazer") : null;
  previa.replaceChildren(h("div", { class: "cabeca-previa" }, abas, desfazer, resumo), folha);
}

function desenharTexto(folha, a) {
  let paginaAnterior = null;
  const erros = errosPorTabela(a);
  for (const [tipo, indice, pagina] of a.ordem) {
    if (pagina !== paginaAnterior) {
      folha.append(h("div", { class: "sep-pagina" }, `página ${pagina} do PDF`));
      paginaAnterior = pagina;
    }
    if (tipo === "t") {
      const t = a.tabelas[indice];
      const pendente = erros.has(indice);
      folha.append(h("div", { class: "marcador" },
        h("span", { class: "margem" }, "tabela"),
        h("div", { class: "cartao" },
          h("strong", {}, `Tabela ${indice + 1}`),
          ` ${t.linhas.length} linhas, ${t.linhas[0] ? t.linhas[0].length : 0} colunas `,
          h("span", { class: `selo ${pendente ? "pendente" : "ok"}` },
            pendente ? "Precisa de conferência" : t.conferida ? "Conferida por você" : "Sem divergência"),
          h("button", { class: "ligacao", type: "button", onclick: () => irParaTabela(indice) }, "Abrir a tabela"))));
      continue;
    }
    if (tipo === "img") {
      folha.append(h("div", { class: "marcador" }, h("span", { class: "margem" }, "imagem"),
        h("div", { class: "cartao" }, "Brasão do anexo, copiado do PDF para o Word")));
      continue;
    }
    const p = a.paragrafos[indice];
    const sel = est.selPar === p.i;
    const classes = ["par", `e-${p.estilo}`, p.atencao.length ? "atencao" : "", sel ? "sel" : ""].join(" ");
    const corpo = h("div", { class: "corpo" });
    if (sel) {
      corpo.append(editorParagrafo(p, a.estilos));
    } else {
      corpo.append(h("p", { class: "txt" }, p.texto || "(vazio)"));
      for (const n of p.atencao) corpo.append(h("p", { class: "nota" }, `Conferir: ${n}`));
    }
    const el = h("div", {
      class: classes, dataset: { i: p.i },
      tabindex: sel ? null : "0", role: sel ? null : "button",
      "aria-label": sel ? null : `Corrigir trecho com estilo ${p.estilo}`,
      title: `Parte: ${p.parte}`,
    }, h("div", { class: "margem" }, p.estilo), corpo);
    if (!sel) {
      const abrir = () => { est.selPar = p.i; desenharPrevia(); document.querySelector(`.par[data-i="${p.i}"] textarea`)?.focus(); };
      el.addEventListener("click", abrir);
      el.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); abrir(); } });
    }
    folha.append(el);
  }
}

function editorParagrafo(p, estilos) {
  const area = h("textarea", { "aria-label": "Texto do trecho" });
  area.value = p.texto;
  area.rows = Math.min(14, Math.max(3, Math.ceil(p.texto.length / 80) + 1));
  const seletor = h("select", { "aria-label": "Tipo do parágrafo" },
    estilos.map((s) => h("option", { value: s, selected: s === p.estilo }, s)));
  const aplicar = async () => {
    const r = await executar("editarParagrafo", { indice: p.i, texto: area.value, estilo: seletor.value });
    est.selPar = null;
    aplicarEstado(r);
  };
  const cancelar = () => { est.selPar = null; desenharPrevia(); document.querySelector(`.par[data-i="${p.i}"]`)?.focus(); };
  area.addEventListener("keydown", (e) => { if (e.key === "Escape") cancelar(); });
  return h("div", { class: "editor" }, area,
    p.atencao.length ? h("p", { class: "nota" }, `Conferir: ${p.atencao.join("; ")}`) : null,
    h("div", { class: "linha" },
      h("label", {}, "Tipo do parágrafo "), seletor,
      h("button", { class: "botao principal", type: "button", onclick: aplicar, disabled: est.ocupado }, "Aplicar correção"),
      h("button", { class: "botao", type: "button", onclick: cancelar }, "Cancelar")));
}

function desenharTabelas(folha, a) {
  if (!a.tabelas.length) {
    folha.append(h("p", { class: "sem-tabelas" }, "O programa não encontrou tabelas neste ato."));
    return;
  }
  const erros = errosPorTabela(a);
  for (const t of a.tabelas) {
    const pendente = erros.has(t.i);
    const bloco = h("section", { class: "bloco-tabela", id: `tabela-${t.i}` },
      h("header", {},
        h("h3", {}, `Tabela ${t.i + 1}`),
        h("span", { class: "onde" }, `${t.paginas.length > 1 ? "páginas" : "página"} ${t.paginas.join(", ")} do PDF`),
        h("span", { class: `selo ${pendente ? "pendente" : "ok"}` },
          pendente ? "Precisa de conferência" : t.conferida ? "Conferida por você" : "Sem divergência")));
    if (pendente) bloco.append(h("ul", { class: "problemas" }, erros.get(t.i).split("; ").map((x) => h("li", {}, x))));
    const tabela = h("table", {});
    t.linhas.forEach((linha, r) => {
      tabela.append(h("tr", {}, linha.map((valor, c) => {
        const sel = est.selCel && est.selCel.t === t.i && est.selCel.r === r && est.selCel.c === c;
        return h("td", {
          class: sel ? "sel" : null, tabindex: "0",
          "aria-label": `Linha ${r + 1}, coluna ${c + 1}: ${valor}`,
          onclick: () => { est.selCel = { t: t.i, r, c }; desenharPrevia(); focarEditorCelula(t.i); },
          onkeydown: (e) => { if (e.key === "Enter") { est.selCel = { t: t.i, r, c }; desenharPrevia(); focarEditorCelula(t.i); } },
        }, valor);
      })));
    });
    bloco.append(h("div", { class: "rolagem" }, tabela));
    bloco.append(h("div", { class: "acoes-tabela" },
      h("button", { class: "botao", type: "button", onclick: () => verNoPdf(t.i, 0), disabled: est.ocupado }, "Ver no PDF"),
      pendente ? h("button", { class: "botao principal", type: "button", onclick: () => conferirTabela(t.i), disabled: est.ocupado },
        "Conferi a tabela no PDF") : null));
    if (est.selCel && est.selCel.t === t.i) bloco.append(editorCelula(t));
    folha.append(bloco);
  }
}

function focarEditorCelula(t) {
  document.querySelector(`#tabela-${t} .editor-celula textarea`)?.focus();
}

function editorCelula(t) {
  const { r, c } = est.selCel;
  const area = h("textarea", { id: "celula-texto" });
  area.value = t.linhas[r][c];
  const fechar = () => { est.selCel = null; desenharPrevia(); };
  area.addEventListener("keydown", (e) => { if (e.key === "Escape") fechar(); });
  return h("div", { class: "editor-celula" },
    h("label", { for: "celula-texto" }, `Linha ${r + 1}, coluna ${c + 1}`),
    area,
    h("div", { class: "acoes-tabela" },
      h("button", { class: "botao principal", type: "button", disabled: est.ocupado, onclick: async () => {
        const res = await executar("editarCelula", { tabela: t.i, linha: r, coluna: c, texto: area.value });
        aplicarEstado(res);
      } }, "Aplicar correção"),
      h("button", { class: "botao", type: "button", onclick: () => verNoPdf(t.i, r), disabled: est.ocupado }, "Ver esta linha no PDF"),
      h("button", { class: "botao", type: "button", onclick: fechar }, "Cancelar")),
    h("p", { class: "dica" }, "Depois de corrigir uma célula, confira a tabela inteira no PDF de novo."));
}

async function verNoPdf(t, linha) {
  const r = await executar("recorte", { tabela: t, linha });
  if (!r.ok) return alert(r.erro);
  const img = $("dlg-recorte-img");
  if (img.src) URL.revokeObjectURL(img.src);
  img.src = URL.createObjectURL(new Blob([r.png], { type: "image/png" }));
  $("dlg-recorte-titulo").textContent = r.legenda;
  $("dlg-recorte").showModal();
}

async function conferirTabela(t) {
  const ok = await confirmar(`Você comparou cada rótulo e cada valor da tabela ${t + 1} com o PDF original? ` +
    "Diferenças automáticas serão aceitas após sua conferência; a revisão ficará registrada no relatório, se for gerado.");
  if (!ok) return;
  aplicarEstado(await executar("conferirTabela", { tabela: t }));
}

/* ---------------- ver como no Word ---------------- */
function iconeOlho() {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", "0 0 24 24"); svg.setAttribute("width", "18"); svg.setAttribute("height", "18");
  svg.setAttribute("aria-hidden", "true");
  const olho = document.createElementNS(ns, "path");
  olho.setAttribute("d", "M1.8 12S5.6 5.2 12 5.2 22.2 12 22.2 12 18.4 18.8 12 18.8 1.8 12 1.8 12z");
  olho.setAttribute("fill", "none"); olho.setAttribute("stroke", "currentColor"); olho.setAttribute("stroke-width", "1.8");
  const pupila = document.createElementNS(ns, "circle");
  pupila.setAttribute("cx", "12"); pupila.setAttribute("cy", "12"); pupila.setAttribute("r", "3.2"); pupila.setAttribute("fill", "currentColor");
  svg.append(olho, pupila);
  return svg;
}

function chaveWord() {
  let salvos = "";
  try { salvos = localStorage.getItem(CHAVE_LINKS) || ""; } catch { /* sem armazenamento */ }
  return [est.versaoEstado, $("op-links").checked, est.modelo ? `${est.modelo.name}:${est.modelo.size}` : "padrao", salvos].join("|");
}

function refazerWordSeVisivel() {
  if (est.analise && est.aba === "word") desenharPrevia();
}

function desenharWord(folha) {
  const chave = chaveWord();
  if (est.word && est.word.chave === chave) {
    folha.append(est.word.no);
    return;
  }
  let no;
  if (est.word) {           // mantém a versão anterior na tela enquanto a nova é montada
    no = est.word.no;
    no.classList.add("atualizando");
  } else {
    no = h("div", { class: "vista-word" }, h("p", { class: "carregando-word" }, "Montando o Word…"));
  }
  folha.append(no);
  montarWord(no, chave, ++est.pedidoWord);
}

async function montarWord(no, chave, pedido) {
  const modelo = est.modelo ? await est.modelo.arrayBuffer() : null;
  let r;
  try {
    r = await executar("previa", { modelo, links: $("op-links").checked, linksUsuario: JSON.stringify(lerLinksUsuario()) },
      modelo ? [modelo] : []);
  } catch (e) {
    r = { ok: false, erro: e.message };
  }
  if (pedido !== est.pedidoWord) return;      // já há um pedido mais novo
  no.classList.remove("atualizando");
  if (!r.ok) {
    no.replaceChildren(caixaErro("Não foi possível montar a visualização do Word", r.erro));
    return;
  }
  const area = h("div", { class: "area-word" });
  try {
    await docx.renderAsync(r.wordBytes, area, area, {
      className: "docx", inWrapper: true, breakPages: true, ignoreLastRenderedPageBreak: true,
      renderHeaders: true, renderFooters: true, useBase64URL: true,
    });
  } catch (e) {
    no.replaceChildren(caixaErro("Não foi possível desenhar o Word", String(e && e.message ? e.message : e)));
    return;
  }
  if (pedido !== est.pedidoWord) return;
  ligarWord(area, r.mapa);
  const aviso = h("p", { class: "aviso-word" },
    "Este é o Word que será baixado" + ($("op-links").checked ? ", com os links atuais" : "") + ". " +
    "Clique num parágrafo para corrigir o texto, aplicar negrito, itálico ou sublinhado, dividir, juntar ou apagar; " +
    "clique numa tabela para abri-la. " +
    "Ctrl+clique abre um link. A divisão em páginas não aparece aqui, porque só o Word a calcula." +
    (r.sem_link ? ` ${r.sem_link} ${r.sem_link === 1 ? "citação está" : "citações estão"} em amarelo, sem link.` : ""));
  no.replaceChildren(aviso, area);
  est.word = { chave, no };
  ajustarZoomWord(area);
}

function ajustarZoomWord(area) {
  const pagina = area.querySelector("section.docx");
  const wrapper = area.querySelector(".docx-wrapper");
  if (!pagina || !wrapper) return;
  wrapper.style.zoom = "";
  const disponivel = area.clientWidth - 24;
  const largura = pagina.offsetWidth;
  if (largura > disponivel && disponivel > 0) wrapper.style.zoom = String(Math.max(.45, disponivel / largura));
}
window.addEventListener("resize", () => {
  const area = document.querySelector(".area-word");
  if (area) ajustarZoomWord(area);
});

function ligarWord(area, mapa) {
  const paragrafos = [...area.querySelectorAll("section.docx article > p")];
  let j = 0;
  for (const el of paragrafos) {
    const k = el.textContent.replace(/\s+/g, "");
    if (!k) continue;
    let achou = -1;
    for (let d = 0; d < 8 && j + d < mapa.length; d++) {
      if (mapa[j + d].k === k) { achou = j + d; break; }
    }
    if (achou < 0) continue;
    j = achou + 1;
    const i = mapa[achou].i;
    if (i < 0) continue;
    el.dataset.i = i;
    el.classList.add("editavel");
    el.tabIndex = 0;
    el.setAttribute("role", "button");
    el.setAttribute("aria-label", "Corrigir este trecho");
  }
  area.querySelectorAll("section.docx article > table").forEach((tb, t) => {
    tb.classList.add("editavel-tabela");
    tb.dataset.t = t;
    tb.title = `Tabela ${t + 1}: clique para abrir na aba Tabelas`;
  });
  area.querySelectorAll("a[href]").forEach((a) => { a.title = `${a.href}\nCtrl+clique para abrir`; });
  area.addEventListener("click", (e) => {
    const link = e.target.closest("a[href]");
    if (link) {
      e.preventDefault();
      if (e.ctrlKey || e.metaKey) { window.open(link.href, "_blank", "noopener"); return; }
    }
    const p = e.target.closest("p.editavel");
    if (p) return abrirEdicaoWord(Number(p.dataset.i));
    const tb = e.target.closest("table.editavel-tabela");
    if (tb) irParaTabela(Number(tb.dataset.t));
  });
  area.addEventListener("keydown", (e) => {
    const p = e.target.closest && e.target.closest("p.editavel");
    if (p && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); abrirEdicaoWord(Number(p.dataset.i)); }
  });
}

/* ---------------- editor do trecho (aberto pela visualização do Word) ---------------- */
const editor = { indice: null, formatou: false, inicial: "", estiloInicial: "", temNbsp: false, cursor: null };
const areaEditor = $("dlg-editar-texto");

function runsParaDom(runs) {
  return runs.map((r) => {
    let el = document.createDocumentFragment();
    r.t.split("\n").forEach((parte, k) => {
      if (k) el.append(document.createElement("br"));
      if (parte) el.append(document.createTextNode(parte));
    });
    for (const [flag, tag] of [["u", "u"], ["i", "i"], ["b", "b"]]) {
      if (r[flag]) { const env = document.createElement(tag); env.append(el); el = env; }
    }
    return el;
  });
}

function estiloDoTexto(no) {
  const el = no.parentElement;
  const cs = getComputedStyle(el);
  const b = cs.fontWeight === "bold" || Number(cs.fontWeight) >= 600;
  const i = cs.fontStyle === "italic" || cs.fontStyle === "oblique";
  let u = false;
  for (let n = el; n; n = n.parentElement) {
    if (getComputedStyle(n).textDecorationLine.includes("underline")) { u = true; break; }
    if (n === areaEditor) break;
  }
  return { b, i, u };
}

/* Converte o conteúdo do editor em trechos {t, b, i, u}. A marca de divisão separa em dois blocos. */
function serializarEditor() {
  const blocos = [[]];
  const por = (t, f) => {
    const lista = blocos[blocos.length - 1];
    const ult = lista[lista.length - 1];
    if (ult && ult.b === f.b && ult.i === f.i && ult.u === f.u) ult.t += t;
    else lista.push({ t, ...f });
  };
  const neutro = { b: false, i: false, u: false };
  const andar = (no) => {
    for (const filho of no.childNodes) {
      if (filho.nodeType === Node.TEXT_NODE) {
        const t = editor.temNbsp ? filho.data : filho.data.replace(/\u00a0/g, " ");
        if (t) por(t, estiloDoTexto(filho));
      } else if (filho.nodeName === "BR") {
        por("\n", neutro);
      } else if (filho.id === "marca-divisao") {
        blocos.push([]);
      } else if (filho.nodeName === "DIV" || filho.nodeName === "P") {
        if (blocos[blocos.length - 1].length) por("\n", neutro);
        andar(filho);
      } else {
        andar(filho);
      }
    }
  };
  andar(areaEditor);
  return blocos;
}

function erroEditor(texto) {
  const el = $("dlg-editar-erro");
  el.textContent = texto || "";
  el.hidden = !texto;
}

function vizinhosNoDocumento(a, i) {
  const pos = (k) => a.ordem.findIndex(([tipo, indice]) => tipo === "p" && indice === k);
  const aqui = pos(i);
  return {
    anterior: i > 0 && pos(i - 1) === aqui - 1,
    proximo: i < a.paragrafos.length - 1 && pos(i + 1) === aqui + 1,
  };
}

function atualizarBotoesFmt() {
  document.querySelectorAll("#dlg-editar .fmt").forEach((bt) => {
    let ativo = false;
    try { ativo = document.queryCommandState(bt.dataset.cmd); } catch { /* sem suporte */ }
    bt.setAttribute("aria-pressed", String(ativo));
  });
}

async function abrirEdicaoWord(i) {
  const a = est.analise;
  const p = a && a.paragrafos[i];
  if (!p || est.ocupado) return;
  const r = await executar("trecho", { indice: i });
  if (!r.ok) return alert(r.erro);
  editor.indice = i;
  editor.formatou = false;
  editor.cursor = null;
  editor.temNbsp = p.texto.includes("\u00a0");
  areaEditor.replaceChildren(...runsParaDom(r.runs));
  editor.inicial = JSON.stringify(serializarEditor()[0]);
  editor.estiloInicial = p.estilo;
  $("dlg-editar-estilo").replaceChildren(...a.estilos.map((s) => h("option", { value: s, selected: s === p.estilo }, s)));
  const nota = $("dlg-editar-nota");
  nota.hidden = !p.atencao.length;
  nota.textContent = p.atencao.length ? `Conferir: ${p.atencao.join("; ")}` : "";
  erroEditor("");
  const viz = vizinhosNoDocumento(a, i);
  const semVizinho = "Não há parágrafo vizinho nessa direção, ou há uma tabela ou imagem no meio.";
  $("ed-juntar-ant").disabled = !viz.anterior;
  $("ed-juntar-ant").title = viz.anterior ? "Junta este parágrafo ao de cima" : semVizinho;
  $("ed-juntar-prox").disabled = !viz.proximo;
  $("ed-juntar-prox").title = viz.proximo ? "Junta o parágrafo de baixo a este" : semVizinho;
  try { document.execCommand("styleWithCSS", false, false); } catch { /* sem suporte */ }
  $("dlg-editar").showModal();
  areaEditor.focus();
  const sel = window.getSelection();
  sel.selectAllChildren(areaEditor);
  sel.collapseToEnd();
  atualizarBotoesFmt();
}

function editorMudou() {
  return JSON.stringify(serializarEditor()[0]) !== editor.inicial || editor.formatou ||
    $("dlg-editar-estilo").value !== editor.estiloInicial;
}

async function concluirEdicao(promessa) {
  const r = await promessa;
  if (!r.ok) { erroEditor(r.erro); return false; }
  $("dlg-editar").close();
  aplicarEstado(r);
  return true;
}

document.querySelectorAll("#dlg-editar .fmt").forEach((bt) => {
  bt.addEventListener("mousedown", (e) => e.preventDefault());   // mantém a seleção no texto
  bt.addEventListener("click", () => {
    areaEditor.focus();
    document.execCommand(bt.dataset.cmd);
    editor.formatou = true;
    atualizarBotoesFmt();
  });
});
areaEditor.addEventListener("input", (e) => {
  if (e.inputType && e.inputType.startsWith("format")) editor.formatou = true;
  erroEditor("");
});
areaEditor.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    erroEditor("Para criar um parágrafo novo, use Dividir no cursor.");
  }
});
areaEditor.addEventListener("paste", (e) => {
  e.preventDefault();
  const texto = (e.clipboardData.getData("text/plain") || "").replace(/\r?\n/g, " ");
  document.execCommand("insertText", false, texto);
});
areaEditor.addEventListener("drop", (e) => e.preventDefault());
document.addEventListener("selectionchange", () => {
  if (!$("dlg-editar").open) return;
  const sel = window.getSelection();
  if (sel.rangeCount && areaEditor.contains(sel.anchorNode)) editor.cursor = sel.getRangeAt(0).cloneRange();
  atualizarBotoesFmt();
});
$("ed-dividir").addEventListener("mousedown", (e) => e.preventDefault());   // não tira o cursor do texto

$("dlg-editar-cancelar").addEventListener("click", () => $("dlg-editar").close());
$("dlg-editar-aplicar").addEventListener("click", async () => {
  if (!editorMudou()) { $("dlg-editar").close(); return; }
  await concluirEdicao(executar("aplicarTrecho", {
    indice: editor.indice, runs: JSON.stringify(serializarEditor()[0]),
    estilo: $("dlg-editar-estilo").value, formatou: editor.formatou,
  }));
});
$("ed-apagar").addEventListener("click", () => concluirEdicao(executar("apagar", { indice: editor.indice })));
$("ed-dividir").addEventListener("click", async () => {
  const sel = window.getSelection();
  let faixa = sel.rangeCount && areaEditor.contains(sel.anchorNode) ? sel.getRangeAt(0) : editor.cursor;
  if (!faixa || !areaEditor.contains(faixa.startContainer)) {
    erroEditor("Clique no ponto do texto onde o parágrafo deve ser dividido.");
    return;
  }
  faixa = faixa.cloneRange();
  faixa.collapse(true);
  const marca = document.createElement("span");
  marca.id = "marca-divisao";
  faixa.insertNode(marca);
  const blocos = serializarEditor();
  marca.remove();
  areaEditor.normalize();
  if (blocos.length !== 2) { erroEditor("Clique no ponto do texto onde o parágrafo deve ser dividido."); return; }
  await concluirEdicao(executar("dividir", {
    indice: editor.indice, antes: JSON.stringify(blocos[0]), depois: JSON.stringify(blocos[1]),
    estilo: $("dlg-editar-estilo").value, formatou: editor.formatou,
  }));
});
async function juntarNoEditor(direcao) {
  if (editorMudou()) {      // guarda primeiro o que foi corrigido aqui
    const r1 = await executar("aplicarTrecho", {
      indice: editor.indice, runs: JSON.stringify(serializarEditor()[0]),
      estilo: $("dlg-editar-estilo").value, formatou: editor.formatou,
    });
    if (!r1.ok) { erroEditor(r1.erro); return; }
    est.analise = r1;
  }
  const alvo = direcao === "anterior" ? editor.indice - 1 : editor.indice;
  const ok = await concluirEdicao(executar("juntar", { indice: alvo }));
  if (!ok) { $("dlg-editar").close(); aplicarEstado(est.analise); }
}
$("ed-juntar-ant").addEventListener("click", () => juntarNoEditor("anterior"));
$("ed-juntar-prox").addEventListener("click", () => juntarNoEditor("proximo"));

/* ---------------- desfazer ---------------- */
async function desfazerUltima() {
  if (!est.analise || !est.analise.pode_desfazer || est.ocupado) return;
  aplicarEstado(await executar("desfazer"));
}
document.addEventListener("keydown", (e) => {
  if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.key.toLowerCase() !== "z") return;
  const alvo = e.target;
  if (document.querySelector("dialog[open]") || alvo.closest("input, textarea, select, [contenteditable='true']")) return;
  e.preventDefault();
  desfazerUltima();
});

/* ---------------- editor do Word (SuperDoc), última etapa antes de baixar ---------------- */
const TEXTOS_BARRA = {
  bold: "Negrito", italic: "Itálico", underline: "Sublinhado", strikethrough: "Tachado",
  "font-family": "Fonte", "font-size": "Tamanho da fonte", "text-color": "Cor da fonte", "highlight-color": "Cor do realce",
  search: "Localizar", link: "Link", image: "Imagem", table: "Tabela", "table-actions": "Ações da tabela",
  "insert-row-before": "Inserir linha acima", "insert-row-after": "Inserir linha abaixo",
  "insert-column-before": "Inserir coluna à esquerda", "insert-column-after": "Inserir coluna à direita",
  "delete-row": "Excluir linha", "delete-column": "Excluir coluna", "delete-table": "Excluir tabela",
  "remove-borders": "Remover bordas", "merge-cells": "Mesclar células", "split-cell": "Dividir célula", "fix-tables": "Corrigir tabelas",
  "text-align": "Alinhamento", "bullet-list": "Marcadores", "numbered-list": "Numeração",
  "indent-decrease": "Diminuir recuo", "indent-increase": "Aumentar recuo", zoom: "Zoom", "measurement-unit": "Unidade de medida",
  undo: "Desfazer", redo: "Refazer", "clear-formatting": "Limpar formatação", "copy-format": "Pincel de formatação",
  "line-height": "Espaçamento entre linhas", "linked-style-label": "Estilos", "linked-style": "Estilos", ruler: "Régua",
  "formatting-marks": "Marcas de formatação",
  "document-mode-editing": "Edição", "document-mode-suggesting": "Sugestão", "document-mode-viewing": "Visualização",
};
const ITENS_FORA_DA_BARRA = ["acceptTrackedChangeBySelection", "rejectTrackedChangeOnSelection", "documentMode", "ai"];

async function carregarEditor() {
  if (!document.querySelector("link[data-superdoc]")) {
    const css = h("link", { rel: "stylesheet", href: "editor/superdoc.css" });
    css.dataset.superdoc = "1";
    document.head.append(css);
  }
  if (!window.SuperDoc) await import("./editor/superdoc.js");
  return window.SuperDoc;
}

function entregarNoEditor(r) {
  const nomeWord = est.analise.nome_sugerido;
  est.ultimoResultado = { ...r, nomeWord };
  const restantes = (r.pendentes_busca || []).map((p) => p.exemplo);
  $("gerar-saida").replaceChildren(h("div", { class: "resultado", role: "status" },
    h("h3", {}, "Word aberto no editor"),
    h("p", {}, nomeWord),
    r.resumo_links ? h("p", {}, r.resumo_links) : null,
    restantes.length ? h("p", {}, `Ficaram em amarelo no Word: ${restantes.join("; ")}.`) : null,
    h("div", { class: "botoes" },
      h("button", { class: "botao", type: "button", onclick: () => abrirEditor(r) }, "Voltar ao editor"),
      h("button", { class: "botao", type: "button", onclick: () => salvarArquivo(r.wordBytes, nomeWord, "docx", TIPO_DOCX) }, "Salvar sem editar"),
      h("button", { class: "botao", type: "button", onclick: () => pdfDoResultado(r) }, "Gerar o PDF"))));
  abrirEditor(r);
}

function estadoEditor(texto) {
  $("editor-estado").textContent = texto;
}

function mostrarEditor() {
  $("editor-word").hidden = false;
  document.body.classList.add("com-editor");
}

async function abrirEditor(r) {
  if (est.editor && est.editor.r === r) {          // volta à mesma sessão, com as edições feitas
    mostrarEditor();
    return est.editor;
  }
  if (est.editor && est.editor.mudou) {
    const ok = await confirmar("O editor tem alterações num Word anterior que ainda não foram salvas. " +
      "Abrir o Word novo descarta essas alterações.", "Abrir o Word novo", "Voltar ao Word anterior");
    if (!ok) { mostrarEditor(); return est.editor; }
  }
  descartarEditor();
  const nome = est.analise.nome_sugerido;
  est.editor = { r, nome, mudou: false, pronto: false, sd: null };
  est.editor.prontoPromessa = new Promise((resolver) => { est.editor.avisarPronto = resolver; });
  $("editor-arquivo").textContent = nome;
  $("editor-relatorio").hidden = !r.relatorioBytes;
  $("editor-baixar").disabled = true;
  $("editor-pdf").disabled = true;
  $("editor-barra").replaceChildren();
  $("editor-documento").replaceChildren();
  estadoEditor("Abrindo o editor… (na primeira vez, ele é baixado e pode levar alguns segundos)");
  mostrarEditor();
  try {
    const SuperDoc = await carregarEditor();
    const sessao = est.editor;
    sessao.sd = new SuperDoc({
      selector: "#editor-documento",
      document: new File([r.wordBytes], nome, { type: TIPO_DOCX }),
      documentMode: "editing",
      ui: {
        toolbar: { container: "#editor-barra", excludeItems: ITENS_FORA_DA_BARRA, strings: TEXTOS_BARRA },
        comments: false,
      },
      // todas as páginas ficam montadas (os atos são curtos), o que permite salvar o PDF completo
      layoutEngineOptions: { virtualization: { enabled: false } },
      measurementUnit: "cm",
      // em telas estreitas (celular), a página encolhe para caber na largura
      ...(window.innerWidth < 860 ? { zoom: { mode: "fit-width" } } : {}),
      onReady: () => {
        sessao.pronto = true;
        const marcarAjuste = () => {
          let modo = "manual";
          try { modo = sessao.sd.getZoomState().mode; } catch { /* versão sem getZoomState */ }
          $("editor-documento").classList.toggle("ajustado", modo === "fit-width");
        };
        marcarAjuste();
        try { sessao.sd.on("zoomChange", marcarAjuste); } catch { /* sem o evento */ }
        $("editor-baixar").disabled = false;
        $("editor-pdf").disabled = false;
        estadoEditor("Pronto para editar");
        sessao.avisarPronto(sessao);
      },
      onEditorUpdate: () => {
        if (!sessao.pronto) return;
        sessao.mudou = true;
        estadoEditor("Há alterações ainda não salvas");
      },
      onException: (e) => console.warn("Editor:", e),
    });
  } catch (e) {
    estadoEditor("");
    $("editor-documento").replaceChildren(caixaErro("Não foi possível abrir o editor",
      `${e && e.message ? e.message : e}\nO Word gerado continua disponível em Salvar sem editar.`));
    est.editor.avisarPronto(null);
  }
  return est.editor;
}

async function baixarDoEditor() {
  const sessao = est.editor;
  if (!sessao || !sessao.sd) return;
  const destino = await escolherDestino(sessao.nome, "docx");
  if (!destino) return;
  $("editor-baixar").disabled = true;
  estadoEditor("Preparando o Word…");
  try {
    const blob = await sessao.sd.export({ exportType: ["docx"], triggerDownload: false });
    const salvoComo = await gravar(destino, blob, TIPO_DOCX);
    sessao.mudou = false;
    estadoEditor(`Word editado salvo: ${salvoComo}`);
    document.querySelectorAll(".aviso-editor").forEach((el) => el.remove());
  } catch (e) {
    estadoEditor("");
    alert(`Não foi possível gerar o Word editado: ${e && e.message ? e.message : e}`);
  } finally {
    $("editor-baixar").disabled = false;
  }
}

/* PDF: usa a impressão do navegador sobre as páginas montadas pelo editor. O resultado é um PDF
   com texto pesquisável e a mesma paginação da tela. O nome sugerido vem do título da página. */
let estiloPagina = null;
const esperar = (ms) => new Promise((resolver) => setTimeout(resolver, ms));

/* O editor só mantém desenhadas as páginas próximas da tela. Para imprimir, passa por todas,
   copia cada página já desenhada para uma área de impressão e imprime só essa área. */
async function copiarPaginasParaImpressao() {
  const area = $("editor-documento");
  const layout = area.querySelector(".superdoc-layout");
  const paginas = [...area.querySelectorAll(".superdoc-page")];
  if (!layout || !paginas.length) return null;
  const raiz = h("div", { id: "impressao" });
  let destino = raiz;
  const cadeia = [];
  for (let el = layout; el && el !== area; el = el.parentElement) cadeia.unshift(el);
  for (const el of cadeia) {          // mesmos ancestrais (classes), para o estilo do editor valer nas cópias
    const copia = el.cloneNode(false);
    copia.removeAttribute("id");
    destino.append(copia);
    destino = copia;
  }
  const posicao = area.scrollTop;
  for (const pagina of paginas) {
    pagina.scrollIntoView({ block: "start" });
    for (let k = 0; k < 40 && !pagina.firstElementChild; k++) await esperar(50);
    await esperar(80);
    const copia = pagina.cloneNode(true);
    copia.querySelectorAll('[class*="sd-v2-local-selection"], [class*="remote-cursor"]').forEach((el) => el.remove());
    destino.append(copia);
  }
  area.scrollTop = posicao;
  const ultima = destino.lastElementChild;
  if (ultima) ultima.classList.add("ultima-pagina");
  document.body.append(raiz);
  return raiz;
}

async function salvarPdf(opcoes = {}) {
  const pagina = document.querySelector("#editor-documento .superdoc-page");
  if (!pagina || !est.editor) return;
  $("editor-pdf").disabled = true;
  estadoEditor("Preparando as páginas…");
  document.activeElement && document.activeElement.blur && document.activeElement.blur();
  const sel = window.getSelection();
  if (sel) sel.removeAllRanges();
  document.getElementById("impressao")?.remove();
  // a impressão precisa das páginas no tamanho real (100%), mesmo que a tela esteja com outro zoom
  const sd = est.editor.sd;
  let zoomAntes = null;
  try {
    if (sd && sd.getZoom() !== 100) {
      zoomAntes = sd.getZoomState ? sd.getZoomState() : { mode: "manual", value: sd.getZoom() };
      sd.setZoom(100);
      await esperar(600);
    }
  } catch { zoomAntes = null; }
  const paginaReal = document.querySelector("#editor-documento .superdoc-page") || pagina;
  const { width, height } = paginaReal.getBoundingClientRect();
  const impressao = await copiarPaginasParaImpressao();
  try {
    if (zoomAntes) {
      if (zoomAntes.mode === "fit-width") sd.setZoomMode("fit-width");
      else sd.setZoom(zoomAntes.value || 100);
    }
  } catch { /* mantém 100% */ }
  $("editor-pdf").disabled = false;
  if (!impressao) return;
  estiloPagina = estiloPagina || document.head.appendChild(document.createElement("style"));
  estiloPagina.textContent = `@page { size: ${width}px ${height}px; margin: 0; }`;
  const tituloAntes = document.title;
  document.title = est.editor.nome.replace(/\.docx$/i, "");
  document.body.classList.add("imprimindo");
  estadoEditor("Na janela de impressão, escolha Salvar como PDF");
  const restaurar = () => {
    document.body.classList.remove("imprimindo");
    impressao.remove();
    document.title = tituloAntes;
    estadoEditor(est.editor && est.editor.mudou ? "Há alterações ainda não salvas" : "Pronto para editar");
    if (opcoes.fecharDepois) fecharEditor();
  };
  window.addEventListener("afterprint", restaurar, { once: true });
  window.print();
}

function descartarEditor() {
  if (est.editor && est.editor.sd) {
    try { est.editor.sd.destroy(); } catch { /* já destruído */ }
  }
  est.editor = null;
  $("editor-documento").replaceChildren();
  $("editor-barra").replaceChildren();
}

/* Fechar só esconde: as edições continuam lá até gerar outro Word, e "Voltar ao editor" as traz de volta. */
function fecharEditor() {
  $("editor-word").hidden = true;
  document.body.classList.remove("com-editor");
  const sessao = est.editor;
  if (sessao && sessao.mudou) {
    const aviso = $("gerar-saida").querySelector(".resultado");
    if (aviso && !aviso.querySelector(".aviso-editor")) {
      aviso.append(h("p", { class: "msg-link aviso-editor" },
        "O editor tem alterações ainda não salvas. Use Voltar ao editor e depois Salvar o Word editado."));
    }
  }
}

$("editor-baixar").addEventListener("click", baixarDoEditor);
$("editor-pdf").addEventListener("click", () => salvarPdf());

/* Gerar PDF direto (sem editar): abre o Word no editor só para montar as páginas, abre a impressão
   e fecha o editor em seguida. */
async function pdfDoResultado(r) {
  const sessao = await abrirEditor(r);
  if (!sessao) return;
  estadoEditor("Preparando o PDF…");
  const pronta = await sessao.prontoPromessa;
  if (!pronta || est.editor !== sessao) return;
  await esperar(400);
  await salvarPdf({ fecharDepois: true });
}

function entregarPdf(r) {
  const nomeWord = est.analise.nome_sugerido;
  est.ultimoResultado = { ...r, nomeWord };
  const restantes = (r.pendentes_busca || []).map((p) => p.exemplo);
  $("gerar-saida").replaceChildren(h("div", { class: "resultado", role: "status" },
    h("h3", {}, "PDF enviado para a impressão"),
    h("p", {}, "Na janela que abriu, escolha Salvar como PDF; ali você escolhe a pasta e o nome do arquivo."),
    r.resumo_links ? h("p", {}, r.resumo_links) : null,
    restantes.length ? h("p", {}, `Ficaram em amarelo: ${restantes.join("; ")}.`) : null,
    h("div", { class: "botoes" },
      h("button", { class: "botao", type: "button", onclick: () => pdfDoResultado(r) }, "Abrir a impressão de novo"),
      h("button", { class: "botao", type: "button", onclick: () => salvarArquivo(r.wordBytes, nomeWord, "docx", TIPO_DOCX) }, "Salvar o Word"))));
  pdfDoResultado(r);
}
$("editor-fechar").addEventListener("click", fecharEditor);
$("editor-relatorio").addEventListener("click", () => {
  const r = est.editor && est.editor.r;
  if (r && r.relatorioBytes) salvarArquivo(r.relatorioBytes, est.editor.nome.replace(/\.docx$/i, "_CONFERIR.html"), "html", "text/html");
});
window.addEventListener("beforeunload", (e) => {
  if (est.editor && est.editor.mudou) { e.preventDefault(); e.returnValue = ""; }
});

/* ---------------- estado dos botões ---------------- */
function atualizar() {
  const entrada = entradaAtual();
  $("analisar").disabled = !est.motorPronto || est.ocupado || !entrada.pdf || !entrada.consulta;
  $("analisar").textContent = est.ocupado ? "Processando…" : "Analisar PDF";
  desenharConferencia();
  const a = est.analise;
  if (a && analiseDesatualizada() !== est.avisoDesatualizado) desenharPrevia();
  $("passo-gerar").setAttribute("aria-disabled", a ? "false" : "true");
  const motivos = $("motivos");
  let lista = [];
  if (a) {
    lista = pendenciasParaGerar(a);
    if (analiseDesatualizada()) lista.unshift("Analisar o PDF de novo (o PDF ou o número do ato mudou)");
  }
  $("gerar").disabled = !a || est.ocupado || lista.length > 0;
  $("editar").disabled = $("gerar").disabled;
  $("gerar-pdf").disabled = $("gerar").disabled;
  if (a && lista.length) {
    motivos.replaceChildren(h("li", { style: "list-style:none;margin-left:-18px" }, "Para gerar o Word, falta:"),
      ...lista.map((m) => h("li", {}, m)));
    motivos.hidden = false;
  } else {
    motivos.hidden = true;
  }
}

desenharLinksSalvos();
atualizar();
