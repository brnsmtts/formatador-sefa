# Formatador SEFA 3.18, versão web

É a tela **Importar ato do PDF** rodando inteira no navegador. O Python roda dentro da página com o Pyodide, no computador de quem a abre. Ninguém precisa instalar Python, rodar o INSTALAR.bat ou ter permissão de administrador. O PDF não é enviado para nenhum servidor.

## Testar no seu computador

Dê dois cliques em `SERVIR_LOCAL.bat`. O navegador abre em http://localhost:8000. Para parar, feche a janela preta.

Aberta direto com dois cliques no `index.html`, a página **não funciona**: o navegador bloqueia o carregamento do Python em arquivos abertos do disco.

## Publicar

### GitHub Pages

1. No GitHub, crie um repositório (por exemplo, `formatador-sefa`).
2. Em **Add file > Upload files**, arraste todo o conteúdo desta pasta, mantendo as subpastas `py`, `pyodide` e `wheels`. O maior arquivo tem 18 MB, abaixo do limite de 25 MB do envio pelo navegador.
3. Em **Settings > Pages**, escolha **Deploy from a branch**, a branch `main` e a pasta `/ (root)`.
4. Em alguns minutos o site fica em `https://SEU-USUARIO.github.io/formatador-sefa/`.

No plano gratuito, o repositório precisa ser público.

### Netlify

Entre em https://app.netlify.com/drop e arraste esta pasta inteira. O site é publicado na hora, sem repositório.

### Em qualquer caso

Quem abre o site consegue baixar os `.py`, o `modelo_sefa.docx` e a `tabela_links.tsv`, porque é o navegador que os executa.

O primeiro acesso baixa cerca de 41 MB (Python e bibliotecas). Nos seguintes, o navegador usa o que guardou e a página fica pronta em segundos.

## Atualizar para uma nova versão do formatador

1. Copie para a pasta `py/` os novos `extrator_pdf.py`, `importar_pdf.py`, `formatador_sefa.py`, `links_no_word.py`, `links_legislacao.py`, `modelo_sefa.docx` e a `tabela_links.tsv` atualizada. Eles entram sem nenhuma alteração.
2. Em `worker.js`, aumente o número em `const VERSAO = "3.18-web11";` (por exemplo, para `3.19-web1`). Assim os navegadores não usam os arquivos antigos guardados.
3. Publique de novo.

Se uma versão nova passar a importar outro módulo, acrescente o nome do arquivo à lista `ARQUIVOS` em `worker.js`.

O `web_ponte.py` é o único arquivo próprio da web. Ele:
- desenha os recortes de imagem com o PyMuPDF, porque o pypdfium2 do pdfplumber não tem edição para esta versão do Pyodide;
- faz os links consultarem só a tabela;
- troca as janelas e diálogos pelo envio e download de arquivos.

## PyMuPDF baixado do PyPI

Para o pacote caber num arquivo só, a biblioteca PyMuPDF (18 MB) não vem na pasta `wheels`. No primeiro acesso, o site a baixa do PyPI (files.pythonhosted.org), o repositório oficial de bibliotecas Python; depois o navegador guarda o arquivo. Se a rede bloquear esse endereço, o indicador do topo fica vermelho e a mensagem explica o que fazer: colocar o arquivo `pymupdf-1.28.2-cp313-abi3-pyemscripten_2025_0_wasm32.whl` na pasta `wheels` e publicar de novo. Com o arquivo na pasta, o site passa a usá-lo e não consulta mais o PyPI.

## Diferenças em relação ao desktop

**Hyperlinks.** Só a `tabela_links.tsv` do site é consultada. A página não consegue falar com o site de legislação da SEFA nem com o Planalto, porque o navegador bloqueia pedidos a outros domínios. Citações fora da tabela ficam destacadas em amarelo, e o resumo final lista quais são. Para ampliar a cobertura, rode a versão desktop, que acrescenta os atos novos à tabela, e publique a tabela atualizada.

**Busca assistida.** Ao clicar em **Gerar o Word**, a página confere os links antes de baixar. Se todas as citações tiverem link, o Word é baixado na hora. Se alguma ficar sem link, a página a lista primeiro. Cada citação tem um botão que abre o site da SEFA numa nova aba, já pesquisando o número e o ano. Abra o ato certo, copie o endereço da página dele (o que tem `idAto=` no final), cole na página e clique em **Aplicar os links e baixar**, ou em **Baixar sem esses links**. Nos dois casos o Word é baixado uma única vez. Os links colados ficam guardados naquele navegador e entram sozinhos nos próximos atos. Em **Links salvos neste navegador**, dá para remover um link errado ou exportar as linhas no formato da `tabela_links.tsv`, para incorporá-las à tabela oficial do site.

**Ver como no Word.** O botão ao lado das abas da prévia mostra na tela o Word que será baixado, com as fontes e recuos do modelo, as tabelas, os links e os destaques em amarelo. Clicar num parágrafo abre a correção daquele trecho: dá para mudar o texto e o tipo do parágrafo, aplicar negrito, itálico ou sublinhado em parte do texto, dividir o parágrafo no cursor, juntá-lo ao anterior ou ao próximo e apagá-lo. A visualização se refaz depois de cada correção. Clicar numa tabela abre a aba Tabelas, onde as células continuam sendo corrigidas, para manter a exigência de conferir a tabela de novo. O botão **Desfazer** (ou Ctrl+Z fora do texto) volta a última correção ou conferência. Parágrafos apagados aparecem no relatório de conferência, e cada trecho corrigido, unido ou dividido fica anotado nele. A divisão em páginas não aparece, porque só o Word a calcula. A visualização usa as bibliotecas docx-preview (Apache 2.0) e JSZip (MIT), na pasta `lib/`.

**Editor do Word.** O botão **Editar antes de salvar**, no passo 4, gera o Word com as mesmas conferências e links do botão Gerar o Word, mas, em vez de baixá-lo, abre-o num editor parecido com o Word, dentro do próprio site. Ele tem páginas, fonte, tamanho, negrito, itálico, sublinhado, cores, realce, links, tabelas, alinhamento, listas, recuos, espaçamento, estilos do modelo, régua e localizar. **Salvar o Word editado** salva o arquivo com as alterações. **Salvar em PDF** abre a janela de impressão do navegador já com as páginas do editor no tamanho da folha: escolha o destino **Salvar como PDF** e clique em Salvar. O PDF sai com texto pesquisável, a mesma paginação do editor e o nome do ato sugerido. **Fechar o editor** só esconde a tela: as alterações continuam lá, e **Voltar ao editor** as traz de volta até você gerar outro Word. As conferências do formatador valem até a geração; o que se altera no editor não entra no relatório, como acontece quando se edita o Word baixado. O editor é a SuperDoc (licença AGPL-3.0, código aberto), empacotada na pasta `editor/`. Ele só é baixado quando alguém o abre pela primeira vez (cerca de 25 MB). A régua usa centímetros. Em telas estreitas, como a do celular, a página é reduzida para caber na largura.

**Hyperlinks até o órgão.** Quando o órgão vem colado ao número, como em "Portaria nº 293/2026-SEFA/GS", o link passa a cobrir o órgão também. Se a data vier logo depois, ela entra no link, como já acontecia com as outras citações. A chave na tabela de links não muda (`Portaria|293|2026`). Essa mudança foi feita no `links_legislacao.py`, que é o mesmo arquivo do desktop: copie-o também para a pasta do programa desktop.

**Salvar com escolha de pasta e nome.** No Edge e no Chrome, os botões que salvam arquivos (Gerar o Word, Salvar o Word de novo, Salvar o relatório, Salvar o Word editado e Exportar para a tabela) abrem a janela "Salvar como", com o nome do ato já sugerido; o navegador lembra a última pasta usada. Em Gerar o Word, a janela aparece logo no clique e o Word é gravado no lugar escolhido assim que fica pronto. Em navegadores sem esse recurso, os arquivos vão para a pasta de downloads.

**Gerar PDF.** O botão Gerar PDF, no passo 4 e no quadro do Word gerado, faz as mesmas conferências, monta as páginas e abre a janela de impressão: escolha **Salvar como PDF**, a pasta e o nome. Não é preciso abrir o editor antes.

**Word.** O Word é baixado e não abre sozinho. O relatório de conferência, quando marcado, vem num botão à parte.

**PyMuPDF.** O site usa o PyMuPDF 1.28.2, única versão com edição para navegador. Nos testes, os resultados foram idênticos aos do 1.26.6 do desktop.

## Testes feitos

Foram 18 atos da pasta `exemplos` e de `tests/fixtures`: DOE, DO-e, DOU real, portaria conjunta, viradas de página, hifenização, tabelas entre colunas e anexos. Cada um foi processado no Python do computador e no Python do navegador. Os Word e os relatórios saíram idênticos byte a byte, e os bloqueios se repetiram nos dois: tabela sem grade, célula mesclada e DOU de três colunas.

No Chromium, a tela foi usada como um usuário: análise, correção de parágrafo e de célula, "Ver no PDF", conferência de tabela e de limites, e download do Word. Os arquivos baixados conferiram com a referência do desktop.

## Pastas

| Pasta ou arquivo | Conteúdo |
|---|---|
| `index.html`, `app.js` | a tela |
| `worker.js` | carrega o Python e roda o formatador sem travar a tela |
| `py/` | módulos do formatador, modelo, tabela de links e `web_ponte.py` |
| `pyodide/` | Python 3.13 para navegador (Pyodide 0.29.5) com lxml, Pillow e cryptography |
| `wheels/` | python-docx, pdfplumber, pdfminer.six e PyMuPDF para navegador |
| `lib/` | docx-preview e JSZip, usadas no botão Ver como no Word |
| `editor/` | editor do Word (SuperDoc), usado no botão Editar antes de salvar |
| `SERVIR_LOCAL.bat`, `servir_local.py` | teste no computador |
| `.nojekyll` | faz o GitHub Pages servir os arquivos sem processá-los |
