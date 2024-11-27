# **Monitoramento de Preços em Sites e Salvamento em Excel**

Este projeto é um **script em Python** que monitora o preço de um produto em um site (exemplo: Amazon) e salva as informações coletadas em uma planilha Excel, atualizando os dados automaticamente em intervalos programados. Ideal para quem quer acompanhar promoções ou quedas de preços!

## **Funcionalidades**
- Acessa um produto em uma loja online.
- Extrai o nome, preço e URL do produto.
- Salva os dados coletados em uma planilha Excel, criando automaticamente uma aba e adicionando os dados na próxima linha disponível.
- Executa o monitoramento em intervalos programados.

## **Requisitos**
Certifique-se de ter os seguintes pacotes instalados antes de executar o script:

- **Selenium**: Para automação do navegador.
- **OpenPyXL**: Para manipulação de arquivos Excel.
- **Schedule**: Para agendar a execução do script.

Para instalar os pacotes necessários:
```bash
pip install selenium openpyxl schedule

