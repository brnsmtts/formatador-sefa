# Formatador SEFA 3.24, versão web

O site tem duas abas no topo: **Formatar ato**, a tela Importar ato do PDF do Formatador SEFA, e **Aplicar alterações**, o Consolidador SEFA (veja a seção própria). As duas rodam inteiras no navegador. O Python roda dentro da página com o Pyodide, no computador de quem a abre. Ninguém precisa instalar Python, rodar o INSTALAR.bat ou ter permissão de administrador. O PDF não é enviado para nenhum servidor.

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
2. Em `worker.js`, aumente o número em `const VERSAO = "3.24-web1";` (por exemplo, para `3.24-web2` ou `3.25-web1`). Assim os navegadores não usam os arquivos antigos guardados.
3. Publique de novo.

Se uma versão nova passar a importar outro módulo, acrescente o nome do arquivo à lista `ARQUIVOS` em `worker.js`.

Antes de copiar, confira se a versão nova do computador traz as correções feitas a partir da versão web (o fim do ALTERACOES dela deve falar delas). A 3.21 a 3.23 foram feitas sem elas, e foi preciso juntá-las de novo. A 3.24 do computador já traz todas.

O `web_ponte.py` é o único arquivo próprio da web. Ele:
- desenha os recortes de imagem com o PyMuPDF, porque o pypdfium2 do pdfplumber não tem edição para esta versão do Pyodide;
- faz os links consultarem só a tabela;
- troca as janelas e diálogos pelo envio e download de arquivos.

## Aba Aplicar alterações (Consolidador SEFA 1.4.2)

A aba **Aplicar alterações** é o Consolidador SEFA rodando no navegador. Os módulos da pasta `py/consolidador/` e o `py/modelos/estilos_sefa.xml` vêm do programa do computador, sem alteração; a `gui.py` e a `cli.py` ficam de fora, porque a janela é a própria página. O `py/consolidador_ponte.py` refaz para a página o que a janela do programa faz: a seleção das propostas, a validação da correção, a mudança manual, a aprovação da extração de PDF/HTML e a abertura do plano salvo.

Como usar:
1. Escolha a legislação a ser alterada, em Word (.docx), e o ato alterador (Word, PDF com texto, TXT ou HTML). Não precisa ser o Word gerado na aba Formatar ato.
2. Se quiser, informe o ato alterador, a publicação e os efeitos gerais (em branco, o programa lê do ato) e o ano nas notas.
3. Clique em **Analisar alterações**. Cada mudança vira uma linha da lista, com a situação, a ação, o destino, os efeitos e a origem. Clique numa linha para conferir as abas Texto atual na base, Nova redação, Trecho do ato alterador, Efeitos e remissões e Observações.
4. Para corrigir, edite a nova redação ou abra **Corrigir ação, destino e datas** e clique em **Validar correção** (ou Ctrl+Enter). **Ampliar** mostra a revisão na tela inteira; Esc volta. A coluna **Aplicar** marca o que vai para o Word. Há também **Adicionar mudança manual**, **Remover proposta** e **Aprovar extração do PDF/HTML**, como no programa.
5. **Gerar Word atualizado** salva uma cópia, com o nome da base terminado em `_CONSOLIDADO.docx`. **Realçar alterações em amarelo** marca os textos novos e as notas desta geração. **Ver relatório** abre o plano numa aba nova; **Salvar plano** grava o plano (.json) para continuar depois.

Diferenças em relação ao programa do computador:
- **Legislação em .doc**: o navegador não tem Word nem LibreOffice para converter. Abra o arquivo no Word, use Salvar como > Documento do Word (.docx) e escolha a cópia.
- **Abrir plano salvo**: escolha antes, nos passos 1 e 2, a mesma legislação e o mesmo ato alterador do plano; eles são reconhecidos pelo conteúdo, e não pela pasta. Assim abre também o plano salvo no programa do computador.
- Não há os botões Abrir base no Word, Abrir ato alterador e Abrir resultado: os arquivos estão no seu computador.
- **Exemplo de uso** carrega os dois Word fictícios da pasta `exemplos/consolidador/`.

Para atualizar o Consolidador numa versão nova, copie os `.py` da pasta `consolidador` do programa (menos `gui.py` e `cli.py`) para `py/consolidador/` e o `modelos/estilos_sefa.xml` para `py/modelos/`, e aumente a `VERSAO` em `worker.js`. Se a versão nova tiver um módulo novo, acrescente o nome dele à lista `ARQUIVOS` em `worker.js`.

Testes: no Python do computador e no do navegador, o exemplo fictício saiu igual ao `03_Resultado_ESPERADO.docx` do programa, e os casos da LC 123 com a LC 238/26 (com a base em .docx), da Constituição Federal com a EC 136/25 e da Lei 5.530 com a Lei 11.233/25 saíram iguais ao Word do motor do programa chamado direto (21, 9 e 4 propostas, todas prontas). Com o ato alterador em PDF, as 4 propostas ficaram pendentes de conferência e, depois de Aprovar extração, o Word saiu igual ao do ato em Word. No navegador, a Constituição levou cerca de 9 segundos para analisar e 19 para gerar.

## Versão 3.24 e Consolidador 1.4.2 (exemplos didáticos)

Correções encontradas com o kit de exemplos didáticos (atos fictícios: a Lei 99.500/2026 num diário de duas colunas, que altera a Lei 99.100/2024; o Decreto 9.850/2026, que altera o Decreto 9.700/2025; e a Lei 99.620/2026 em HTML). Os módulos são os mesmos da 3.24 do programa do computador e do Consolidador 1.4.2.

Aba Formatar ato (`formatador_sefa.py` e `extrator_pdf.py`):
- a citação fechada com **"(AC)"** termina a redação nova, como "(NR)"; antes, o resto do ato saía como redação nova;
- o "(NR)" que cai sozinho na linha de baixo volta para a linha da citação;
- o preâmbulo da lei ("A ASSEMBLEIA LEGISLATIVA DO ESTADO DO PARÁ estatui e eu sanciono a seguinte Lei:") deixa de sair como ementa;
- a citação que continua na outra coluna ou página mantém o estilo de redação nova;
- o título de seção do DOE logo depois da assinatura ("ASSEMBLEIA LEGISLATIVA - LEIS SANCIONADAS") não entra mais no fim do ato.

Aba Aplicar alterações (`py/consolidador/`):
- a citação fechada com "(AC)" não esconde mais a cláusula de efeitos do ato (antes, todas as datas ficavam "não informadas");
- em "quanto ao inciso V do art. 2º e às alíneas “b” e “d” do inciso I do art. 5º da Lei nº 99.100", a lei do fim da enumeração vale para todos os dispositivos;
- as notas de decreto dizem "pelo Decreto" (antes, "pela Decreto");
- o acréscimo que depende de uma renumeração feita mais adiante no mesmo ato fica pronto.

Testes: os 29 atos de referência saíram idênticos aos da 3.23, e os mesmos 33 atos (os 29 e 4 do diário fictício) saíram idênticos no Python do computador e no do navegador. Os três cenários do kit, feitos na tela do site sem nenhum contorno, saíram iguais aos resultados esperados. Os casos reais do Consolidador (CF + EC 136/25, LC 123 + LC 238/26, Lei 5.530 + Lei 11.233/25) continuam iguais aos do programa do computador. Passaram os 188 testes do Formatador e os 106 do Consolidador.

## Versão 3.23

Os módulos `extrator_pdf.py`, `formatador_sefa.py` e `importar_pdf.py` vêm da 3.23 do programa do computador:
- **3.21:** reconhece a Emenda Constitucional ("140", "EC 140" ou "Emenda Constitucional nº 140"), com a data do fim do ato na epígrafe, as Mesas da Câmara e do Senado em separado e o texto entre aspas em parágrafos próprios. Linhas do DOU partidas pela justificação voltam a ser uma linha, e o "e" sozinho na linha fica no fim do inciso.
- **3.22:** o quadro de divulgação da Imprensa Nacional depois das assinaturas do último ato do DOU fica de fora, com aviso; "econômico-" / "financeiro" sai "econômico-financeiro".
- **3.23:** epígrafes com "N.º", como "PORTARIA N.º 737/2026 - SEFA GS", são encontradas, e atos seguidos de outro com essa grafia não o engolem mais.

A 3.21 a 3.23 foram feitas a partir da 3.20 original. As correções que vieram depois dela foram juntadas de novo: o `links_legislacao.py` (link até o órgão, ato que cita a si mesmo, "-" e forma curta na tabela) é o mesmo da versão web 3, e o `extrator_pdf.py` da 3.23 recebeu de novo a correção do hífen invisível em maiúsculas. O pacote da 3.23 atualizada do computador traz os mesmos arquivos.

**Nome do PDF.** O site passou a guardar o PDF com o nome original, como o programa do computador. A data no nome ("2026.10.14.DOE.pdf") entra na linha de publicação, e "Extra" ou "Previsão" no nome geram os mesmos avisos. Antes, o PDF virava "entrada.pdf" e só valia a data impressa no cabeçalho.

Testes: 29 atos (os 18 de sempre, os PDFs de teste da Emenda Constitucional, do quadro do DOU, do "N.º" e dos atos de mesmo número, a Resolução nº 046, as Portarias nº 733 e 734 e a Solução de Consulta nº 010/2026) saíram com Word e relatório idênticos no Python do computador e no do navegador. Os 181 testes do programa do computador passaram.

## Versão web 3 (tabela de links completa)

A tabela de links aceita a forma curta `=idAto` no lugar do endereço. A montagem do `MONTAR_TABELA_LINKS.bat` do programa do computador grava o endereço completo, que serve também para versões anteriores. Para o site achar quase todos os links, monte a tabela no computador e troque por ela a `py/tabela_links.tsv` do site (veja Hyperlinks, em Diferenças em relação ao desktop). Para trocar só a tabela, não é preciso mudar a `VERSAO`: o site confere a cada abertura se os arquivos da pasta `py` mudaram, e basta recarregar a página depois que o GitHub publicar (1 a 2 minutos).

## Versão web 2 (Resolução nº 046/2026)

- **Citação do próprio ato.** Quando o ato cita a si mesmo ("conforme RESOLUÇÃO Nº 046, DE 10 DE AGOSTO DE 2026" dentro da própria Resolução nº 046), a citação fica sem link e sem destaque e não aparece na busca assistida.
- **Ato sem link.** Na busca assistida, cada citação tem a caixa **Não tem link: tirar o destaque amarelo**. Marcada, a citação sai sem link e sem destaque, e a escolha fica guardada no navegador como as demais. Em **Links salvos neste navegador**, ela aparece como "sem link" e pode ser removida. Na `tabela_links.tsv`, o mesmo efeito se obtém com `-` no lugar do endereço (`Lei|6913|2006`, tabulação, `-`).
- **Planalto.** Para leis, leis complementares e decretos que a citação diz federais (como "Lei Complementar Federal nº 160"), aparece também o botão **Abrir no Planalto**. Ele abre o endereço provável do ato no Planalto e o coloca no campo; confira na aba aberta se é o ato certo.
- **Hífen invisível em maiúsculas.** Palavras partidas no fim da linha com o hífen invisível de alguns diários ficavam com um espaço no meio quando escritas em maiúsculas ("SO CIOECONÔMICO", "MADEI RAS"). Agora se unem.

As duas primeiras mudanças estão no `links_legislacao.py` e a última no `extrator_pdf.py`, os mesmos arquivos do programa do computador. O pacote da 3.20 atualizada do computador traz os dois.

## Versão 3.20 (atos com o mesmo número)

Os módulos `extrator_pdf.py`, `formatador_sefa.py` e `importar_pdf.py` vêm da versão 3.20 do programa do computador, sem alteração (veja o ALTERACOES_V3_20.txt dele). Quando o número digitado corresponde a mais de um ato do PDF, Analisar PDF abre a lista dos atos, com a página, a coluna, o título e o começo do texto de cada um, e você escolhe qual importar. Fica marcado o da Secretaria da Fazenda, quando há um só entre eles. Enter ou duplo clique importa; Esc cancela.

O `links_legislacao.py` do site continua sendo o da 3.18 com o ajuste do link até o órgão; a 3.20 do computador voltou a ter o arquivo sem esse ajuste (o pacote da 3.20 atualizada o traz de volta).

## Versão 3.19 (Solução de Consulta)

Na 3.19, os módulos `extrator_pdf.py`, `formatador_sefa.py`, `importar_pdf.py` e `links_no_word.py` passaram a vir da versão 3.19 do programa do computador, sem alteração. Ela reconhece a Solução de Consulta e copia do diário o negrito e o itálico das transcrições (veja o ALTERACOES_V3_19.txt do programa do computador). Na versão web, juntar ou dividir parágrafos na visualização do Word mantém esse negrito e itálico.

O `links_legislacao.py` do site é o da 3.18 com o ajuste do link até o órgão ("Portaria nº 293/2026-SEFA/GS"); a 3.19 do computador ainda não tem esse ajuste.

## PyMuPDF baixado do PyPI

Para o pacote caber num arquivo só, a biblioteca PyMuPDF (18 MB) não vem na pasta `wheels`. No primeiro acesso, o site a baixa do PyPI (files.pythonhosted.org), o repositório oficial de bibliotecas Python; depois o navegador guarda o arquivo. Se a rede bloquear esse endereço, o indicador do topo fica vermelho e a mensagem explica o que fazer: colocar o arquivo `pymupdf-1.28.2-cp313-abi3-pyemscripten_2025_0_wasm32.whl` na pasta `wheels` e publicar de novo. Com o arquivo na pasta, o site passa a usá-lo e não consulta mais o PyPI.

## Diferenças em relação ao desktop

**Hyperlinks.** Só a `tabela_links.tsv` do site é consultada. A página não consegue falar com o site de legislação da SEFA nem com o Planalto, porque o navegador bloqueia pedidos a outros domínios. Citações fora da tabela ficam destacadas em amarelo, e o resumo final lista quais são. Para ampliar a cobertura, rode no programa do computador o `MONTAR_TABELA_LINKS.bat`, que grava na tabela o link de todos os atos do site da SEFA, e troque a `py/tabela_links.tsv` do site pela tabela montada. A tabela montada tem os endereços completos (cerca de 3 MB para 5.738 atos); num teste com 60 mil atos, o Word continuou sendo gerado em cerca de um segundo. A tabela também aceita a forma curta `=idAto` (`Lei|6913|2006`, tabulação, `=101`).

**Busca assistida.** Ao clicar em **Gerar o Word**, a página confere os links antes de baixar. Se todas as citações tiverem link, o Word é baixado na hora. Se alguma ficar sem link, a página a lista primeiro. Cada citação tem um botão que abre o site da SEFA numa nova aba, já pesquisando o número e o ano. Abra o ato certo, copie o endereço da página dele (o que tem `idAto=` no final), cole na página e clique em **Aplicar os links e salvar**, ou em **Salvar sem esses links**. Se o ato não tiver link, marque **Não tem link: tirar o destaque amarelo**. Nos dois casos o Word é baixado uma única vez. Os links colados ficam guardados naquele navegador e entram sozinhos nos próximos atos. Em **Links salvos neste navegador**, dá para remover um link errado ou exportar as linhas no formato da `tabela_links.tsv`, para incorporá-las à tabela oficial do site.

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
| `consolidador.js` | a aba Aplicar alterações |
| `py/` | módulos do formatador, modelo, tabela de links, `web_ponte.py` e `consolidador_ponte.py` |
| `py/consolidador/`, `py/modelos/` | módulos e estilos do Consolidador SEFA |
| `exemplos/consolidador/` | os dois Word fictícios do Exemplo de uso |
| `pyodide/` | Python 3.13 para navegador (Pyodide 0.29.5) com lxml, Pillow e cryptography |
| `wheels/` | python-docx, pdfplumber, pdfminer.six e PyMuPDF para navegador |
| `lib/` | docx-preview e JSZip, usadas no botão Ver como no Word |
| `editor/` | editor do Word (SuperDoc), usado no botão Editar antes de salvar |
| `SERVIR_LOCAL.bat`, `servir_local.py` | teste no computador |
| `.nojekyll` | faz o GitHub Pages servir os arquivos sem processá-los |
