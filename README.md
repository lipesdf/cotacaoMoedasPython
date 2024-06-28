# Sistema de Cotação de Moedas | API | Python - Tkinter

## Descrição do Sistema

O sistema de cotação de moedas é uma aplicação desenvolvida em Python utilizando a biblioteca Tkinter para a interface gráfica. Ele consome um microserviço gratuito da Awesome API para fornecer cotações das principais moedas em relação ao Real (BRL).

## Funcionalidades Principais

- **Consulta de Cotações em Tempo Real:** Permite a consulta das cotações das principais moedas estrangeiras em relação ao Real, utilizando a Awesome API como fonte de dados.
- **Importação de Arquivo Excel:** O sistema suporta a importação de arquivos Excel que contêm as moedas e o período de tempo desejados para a cotação. Os resultados são então anexados e exportados em um novo arquivo Excel.

## Detalhes Técnicos

- **Interface Gráfica:** Desenvolvida com Tkinter, proporcionando uma experiência de usuário intuitiva e amigável.
- **API:** Integração com a Awesome API para obtenção das cotações das moedas.
- **Manipulação de Arquivos Excel:** Utiliza bibliotecas como `pandas` para a leitura e escrita de arquivos Excel.

## Instruções para Uso

1. **Consulta Manual de Cotações:**
   - O usuário pode selecionar as moedas desejadas através da interface gráfica e visualizar as cotações em tempo real.

2. **Importação de Arquivo Excel:**
   - O usuário pode carregar um arquivo Excel contendo as moedas desejadas na primeira coluna da tabela e o período de tempo para a cotação.
   - O sistema processa o arquivo, consulta as cotações na API e gera um novo arquivo Excel com os resultados.

## Observações

- **Formato do Arquivo Excel:** É importante manter o padrão do arquivo Excel a ser importado, com as moedas listadas na primeira coluna da tabela, segue exemplo o arquivo `moedas.xlsx`.
- **Dependências:** Certifique-se de que as bibliotecas `pandas`, `openpyxl`, `tkinter`, `numpy`, `datetime` e `requests` estejam instaladas no ambiente Python utilizado.

## Benefícios

- **Automatização de Tarefas:** Facilita a consulta e a compilação de cotações de moedas, economizando tempo e esforço manual.
- **Flexibilidade:** Permite a análise de cotações em períodos específicos, conforme definido pelo usuário no arquivo Excel.
- **Atualização em Tempo Real:** Garante a obtenção de dados atualizados através da integração com a Awesome API.

Este sistema é uma ferramenta eficiente para quem necessita de um método automatizado e confiável para obter cotações de moedas, tanto para análises financeiras quanto para relatórios periódicos.
