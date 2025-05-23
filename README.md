# comparador-de-planilha

Este projeto contém uma macro VBA que compara duas planilhas do Excel e gera automaticamente um relatório de diferenças entre elas. O relatório destaca células com valores diferentes, diferenças na cor de fundo, diferenças na cor da fonte, bem como células que foram adicionadas ou removidas. A macro também utiliza uma função auxiliar para converter códigos RGB (utilizados pelo Excel para cores) em nomes comuns, facilitando a leitura do relatório.

A função NomeDaCor recebe um código RGB e retorna o nome da cor correspondente, como "Branco", "Preto", "Vermelho", "Verde", entre outras. Isso permite que o relatório exiba não apenas o código da cor, mas também um nome mais legível.

A macro principal CompararPlanilhasComRelatorio compara célula por célula das duas primeiras planilhas dos arquivos abertos no Excel. Você deve abrir dois arquivos do Excel antes de executar a macro. O primeiro arquivo aberto será considerado como a versão original, e o segundo, como a versão atualizada. Para cada diferença encontrada, a macro colore a célula alterada com uma cor indicativa e adiciona uma linha em uma nova aba chamada "Relatório de Alterações", contendo a célula, o tipo de alteração, o valor anterior e o valor atual.

# Utilização do codigo
Para usar esta macro, você precisa do Microsoft Excel com suporte a macros (VBA). Primeiro, abra o Excel e pressione ALT + F11 para acessar o Editor VBA. Em seguida, crie um novo módulo clicando com o botão direito sobre o projeto do seu arquivo, escolhendo “Inserir” > “Módulo”, e cole o código fornecido. Certifique-se de que dois arquivos do Excel estejam abertos, contendo os dados que deseja comparar. Após isso, volte ao Excel, pressione ALT + F8, selecione CompararPlanilhasComRelatorio e clique em “Executar”. A análise será realizada e o relatório será gerado automaticamente na aba "Relatório de Alterações" do segundo arquivo aberto.

As cores usadas para sinalizar as diferenças são:

Marrom escuro (RGB: 153, 102, 0) para valores diferentes

Roxo escuro (RGB: 102, 0, 153) para cor de fundo diferente

Cinza escuro (RGB: 64, 64, 64) para cor da fonte diferente

Marrom (RGB: 102, 51, 0) para células adicionadas ou removidas

Certifique-se de que os dados que deseja comparar estejam na primeira aba de cada arquivo. A macro utiliza a propriedade UsedRange, portanto apenas células com conteúdo ou formatação serão consideradas na análise.

Essa solução é útil para auditar alterações em planilhas de controle, orçamentos, relatórios financeiros ou qualquer outro tipo de dado tabular no Excel, oferecendo uma forma rápida e visual de identificar mudanças entre versões.
