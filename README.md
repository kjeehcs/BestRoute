
# Comparador de Percursos - BestRoute

O **BestRoute** é um aplicativo desenvolvido em Python que permite comparar prazos de entrega entre diferentes transportadoras utilizando planilhas do Excel. Com uma interface gráfica simples, o usuário pode selecionar até 5 planilhas e visualizar qual transportadora oferece o prazo mais rápido para um dado percurso.

## Funcionalidades

- **Carregamento de Planilhas**: Carregue até 5 planilhas Excel contendo dados de transportadoras.
- **Padronização de Dados**: Remove acentos e caracteres especiais, garantindo uma comparação precisa.
- **Comparação de Prazos**: Compara os prazos de entrega entre as transportadoras e indica a mais rápida.
- **Resultados em Excel**: Salva os resultados da comparação em um novo arquivo Excel.
- **Interface Gráfica**: Interface amigável desenvolvida com Tkinter para facilitar o uso.

## Formato da Planilha

As planilhas devem estar no formato padrão com as seguintes colunas:

- **Origem**
- **Destino**
- **UF**
- **Quantidade de Dias Máximo para entrega** (esta deve ser a coluna D)

Certifique-se de que esses nomes de coluna estejam exatamente como especificado para que o programa funcione corretamente.

## Requisitos

Antes de executar o programa, você precisa ter o Python instalado em sua máquina. Você também precisará instalar as seguintes bibliotecas:

```bash
pip install pandas unidecode openpyxl
```

## Como Usar

1. **Selecione as Planilhas**: Clique nos botões para selecionar até 5 planilhas que contêm os dados das transportadoras.
2. **Execute a Comparação**: Após selecionar as planilhas, clique no botão "Executar Comparação" para iniciar o processo.
3. **Visualize os Resultados**: Após a comparação, uma janela aparecerá com o resultado. Você pode optar por salvar os resultados em um arquivo Excel.

## Estrutura do Código

O código é dividido em várias funções, cada uma responsável por uma parte específica do processo:

- `padronizar_texto(texto)`: Remove acentos e aspas dos textos.
- `carregar_planilha(caminho_arquivo)`: Carrega e padroniza os dados de uma planilha Excel.
- `comparar_planilhas(planilhas_dfs, planilhas_nomes)`: Compara os prazos de entrega entre as planilhas.
- `selecionar_arquivo(n)`: Permite ao usuário escolher os arquivos das planilhas.
- `executar_comparacao()`: Executa a comparação e exibe os resultados.

## Contribuições

Contribuições são bem-vindas! Se você tiver sugestões, correções ou melhorias, sinta-se à vontade para abrir uma issue ou enviar um pull request.


## Contato

Para mais informações ou dúvidas, entre em contato pelo e-mail: [kjeehcs@gmail.com].

![Interface do Comparador de Percursos](https://github.com/kjeehcs/BestRoute/blob/main/Assets/Img/BestRoute-img.png?raw=true)
