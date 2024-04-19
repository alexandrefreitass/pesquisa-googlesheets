## Google Sheets Data Search Script

<div align="center">
    <img src="https://github.com/alexandrefreitass/pesquisa-sheets/assets/109884524/ec0a6059-f79f-4916-9f69-883c35e37a77" />
</div>

Este repositório contém um script do Google Apps Script projetado para facilitar a busca de dados dentro de uma planilha do Google Sheets chamada "EFETIVO". Ao digitar o Registro (RE) da pessoa na célula "H6", o script automaticamente recupera e exibe os dados completos dessa pessoa na planilha.

### Funcionalidade

O script é ativado automaticamente quando a célula "H6" é editada na planilha "EFETIVO". Ele busca os dados associados ao RE fornecido e, se encontrados, preenche um conjunto especificado de células com esses dados, permitindo tanto a visualização quanto a manipulação subsequente (como exclusão ou pesquisa de dados existentes).

### Como Funciona

1. **Detecção de Edição**: O script detecta uma edição na célula "H6" através de um gatilho `onEdit`.
2. **Execução da Pesquisa**: Se o RE não for zero, o script limpa o conteúdo de um intervalo específico e carrega os dados da planilha.
3. **Filtragem dos Dados**: Os dados são filtrados baseados no RE, e os resultados correspondentes são reinseridos na planilha para visualização.
4. **Alerta**: Se nenhum dado correspondente for encontrado, uma mensagem de alerta é exibida.

### Formatação condicional

Para que ao pesquisa a linha selecionada se destacasse foi utilizado uma formula personalizada na propriedade de formatação condicional do google sheets, definindo como cor de preenchimento verde claro quando:

```
=AND($C9=$C$6; NOT(ISBLANK($C9)))
```

### Código

**Acesso ao Dashboard**: O dashboard pode ser acessado através de [este link](https://docs.google.com/spreadsheets/d/1l5tpHZQJOGiJdXPJ6gKjpFKPRM7iG1vm8e-EMLRuDCk/edit#gid=712832960).

### Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](https://github.com/alexandrefreitass/pesquisa/blob/master/LICENSE) para mais detalhes.
