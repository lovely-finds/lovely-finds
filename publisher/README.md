# Cantinho dos Achados Publisher — MVP

Painel privado de aprovação e publicação para o Instagram profissional **@cantinho_dos_achados_br**.

## Arquitetura

- **Google Apps Script** serve a interface HTML e executa chamadas à Meta no servidor.
- **Script Properties** guarda o token e o ID da conta; não coloque segredos no GitHub nem no HTML.
- **Google Sheets** armazena rascunhos e histórico (a planilha é criada automaticamente na primeira abertura).
- **GitHub** guarda apenas o código-fonte e assets públicos.

> A primeira versão faz publicação imediata de imagem individual após confirmação explícita no painel. Ainda não implementa agendamento, Reels, carrosséis ou Facebook.

## Arquivos

- `Code.gs`: backend (Meta Graph API + armazenamento).
- `Index.html`: painel com visão geral, criação de rascunhos, prévia e histórico.

## Instalação no Google Apps Script

1. Abra https://script.google.com/ e crie um projeto em branco.
2. Renomeie o projeto para `Cantinho dos Achados Publisher`.
3. Copie o conteúdo de `Code.gs` para o arquivo padrão `Código.gs` (ou `Code.gs`).
4. Crie um arquivo HTML chamado exatamente `Index` e copie o conteúdo de `Index.html`.
5. Abra **Configurações do projeto → Propriedades do script** e adicione:

   | Propriedade | Valor |
   | --- | --- |
   | `META_IG_ACCESS_TOKEN` | Token válido gerado para o aplicativo Meta |
   | `IG_USER_ID` | `17841422988989661` |
   | `META_API_VERSION` | `v27.0` |

6. Salve o projeto.
7. Clique em **Implantar → Nova implantação → Aplicativo da web**.
8. Configure:
   - **Executar como:** sua conta Google.
   - **Quem pode acessar:** somente você (restrito à sua conta).
9. Autorize os serviços solicitados e conclua a implantação.
10. Abra a URL do aplicativo gerada pela implantação.

## Segurança

- Nunca cole o token em `Code.gs`, `Index.html`, GitHub, planilhas ou mensagens.
- Armazene o token somente em **Propriedades do script**.
- Não publique a implantação para qualquer pessoa. Use acesso restrito à proprietária.
- Se o token for exposto, revogue-o na Meta e atualize a propriedade do script.
- A interface não envia o token ao navegador.
- Revise cuidadosamente cada rascunho. O botão de publicação pede confirmação antes da chamada `media_publish`.

## Limitações da primeira versão

- Somente imagem individual no Instagram.
- A imagem precisa estar disponível publicamente por HTTPS; os servidores da Meta precisam conseguir buscá-la.
- O Apps Script guarda o histórico numa planilha privada criada automaticamente.
- Os tokens da Meta podem expirar ou ser revogados; atualize a propriedade quando necessário.
- As cotas gratuitas do Apps Script e os limites da API Meta continuam aplicáveis.

## Próximas entregas

1. Testar a instalação privada e publicar uma imagem de teste pelo painel.
2. Integrar publicações orgânicas na Página do Facebook.
3. Adicionar Reels e carrosséis.
4. Criar fila e agendamento.
5. Adicionar seleção de canais e relatórios.
