# Relatório de Infrações

Lê a planilha geral de infrações (.xlsx), separa por **EMPRESA** e gera PDF só com as colunas:
Placa, Enquadramento, Pontos, Local, Município, Gravidade, Infrator, Data, Hora, Valor, Vencimento.

## Como usar
Abra o `index.html` no navegador (duplo clique). Funciona 100% offline — as bibliotecas estão em `vendor/`.

1. Arraste o .xlsx na área indicada.
2. Marque as empresas desejadas.
3. **PDF** (na linha) → só aquela empresa · **PDF único** → todas selecionadas, uma por página · **PDFs separados** → um arquivo por empresa.

Empresas com grafias diferentes de maiúsculas/acentos são agrupadas juntas. A logo do PDF fica em `logo.js` (gerada de `logo.png`).

## Filtro de período
Antes de gerar o PDF, escolha a data (infração, vencimento ou cadastro) e o intervalo "De / Até". Para um dia só, use a mesma data nos dois campos. O período sai impresso no PDF.

## Notificações (Comunicação Administrativa)
Aba **Notificações**: arraste o(s) PDF(s) gerados na aba Relatório (ou use "Usar a planilha carregada"). O sistema gera Word e PDF por município (o zip vem com as pastas `Word/` e `PDF/`) a partir de `modelos/modelo-notificacao.docx`, trocando os campos:

| Campo | Conteúdo |
|---|---|
| `{{CIDADE}}` / `{{UF}}` | Município (editável na tela) e UF |
| `{{DATA_RELATORIO}}` | Data de emissão do relatório |
| `{{DATA_EXTENSO}}` | Data da comunicação por extenso |
| `{{INFRACOES}}` / `{{PONTOS}}` | "23 infrações" / "82 pontos" (singular quando for 1) |
| `{{VALOR_TOTAL}}` | Valor total |
| `{{PLACAS}}` | "as placas A, B e C" / "a placa A" |
| `{{PERIODO}}` | "entre 16/03/2023 e 17/04/2026" / "em 14/04/2026" |
| `{{NIC_VALORES}}` | Valores das multas por não identificação do condutor. **Sem nenhuma, o parágrafo inteiro é removido.** |

Os PDFs levam os dados embutidos (campo Palavras-chave), por isso a aba só lê PDFs gerados por este sistema.

**Trocar o texto do modelo:** edite `modelos/modelo-notificacao.docx` no Word mantendo os campos `{{...}}` (cada campo digitado de uma vez, sem formatação no meio) e rode:
```
echo "const MODELO_NOTIFICACAO='$(base64 -w0 modelos/modelo-notificacao.docx)';" > modelos/modelo-notificacao.js
```
Qualquer parágrafo cujo campo fique sem valor some do documento.

## Enviar ao chefe (WhatsApp)
O botão **Enviar ao chefe** baixa o PDF das empresas selecionadas e abre o WhatsApp Web direto na conversa do número configurado em `app.js` (`WPP_CHEFE`), com um resumo já escrito. Por segurança, nenhum navegador deixa um site anexar arquivo no WhatsApp sozinho: é só arrastar o PDF baixado para a conversa e enviar.

## Publicar (acessar de qualquer máquina)
No GitHub: Settings → Pages → Source: *Deploy from a branch* → `main` / `(root)`. O site fica em `https://<usuario>.github.io/relatorio-infracoes/`. A planilha é processada só no navegador — nada é enviado para servidor.
