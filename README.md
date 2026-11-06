# Relatório de Infrações

Lê a planilha geral de infrações (.xlsx), separa por **EMPRESA** e gera PDF só com as colunas:
Placa, Enquadramento, Pontos, Local, Município, Gravidade, Infrator, Data, Hora, Valor, Vencimento.

## Como usar
Abra o `index.html` no navegador (duplo clique). Funciona 100% offline — as bibliotecas estão em `vendor/`.

1. Arraste o .xlsx na área indicada.
2. Marque as empresas desejadas.
3. **PDF** (na linha) → só aquela empresa · **PDF único** → todas selecionadas, uma por página · **PDFs separados** → um arquivo por empresa.

Empresas com grafias diferentes de maiúsculas/acentos são agrupadas juntas. A logo do PDF fica em `logo.js` (gerada de `logo.png`).

## Enviar ao chefe (WhatsApp)
O botão **Enviar ao chefe** baixa o PDF das empresas selecionadas e abre o WhatsApp Web direto na conversa do número configurado em `app.js` (`WPP_CHEFE`), com um resumo já escrito. Por segurança, nenhum navegador deixa um site anexar arquivo no WhatsApp sozinho: é só arrastar o PDF baixado para a conversa e enviar.

## Publicar (acessar de qualquer máquina)
No GitHub: Settings → Pages → Source: *Deploy from a branch* → `main` / `(root)`. O site fica em `https://<usuario>.github.io/relatorio-infracoes/`. A planilha é processada só no navegador — nada é enviado para servidor.
