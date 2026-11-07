# Gabriel & Maeva

Site do casamento · 17 de setembro de 2027 · Convento de Mafra & Casa de Reguengos, Lisboa.

O site inteiro está em `index.html` e é publicado com GitHub Pages.

## Receber as confirmações numa folha de cálculo

1. Crie uma folha no Google Sheets (por exemplo, "Check-in GM 1709").
2. Abra **Extensões › Apps Script**, apague o que lá estiver e cole o conteúdo de `google-apps-script.gs`. Guarde.
3. Clique em **Implementar › Nova implementação**, escolha o tipo **Aplicação Web**:
   - Executar como: **Eu**
   - Quem tem acesso: **Qualquer pessoa**
4. Clique em **Implementar**, autorize o acesso e copie o **URL da aplicação Web** (termina em `/exec`).
5. Em `index.html`, cole esse URL em `var RSVP_ENDPOINT='';`.

Cada check-in aparece numa linha da aba **Check-in**. Se a mesma pessoa voltar a fazer o check-in, a linha dela é atualizada.
