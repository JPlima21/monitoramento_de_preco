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
```
# **Como usar:**
1. Clone este repositório:
```bash
git clone https://github.com/seu-usuario/monitoramento-de-precos.git
cd monitoramento-de-precos
```
2. Configure o ambiente
- Certifique-se de que o **Google Chrome** está instalado.
- Baixe o **ChromeDriver** compatível com sua versão do navegador e adicione-o ao PATH.

3. Edite o script se necessário: 
No script, substitua o link do produto na variavel **URL** para o produto que deseja monitorar:
```bash
url = 'URL_DO_PRODUTO_AQUI'
```

4. Execute o script:
```bash
python monitoramento.py
```

5. Acompanhe os dados salvos: O script salva as informações na planilha
```bash
Planilha_de_preços.xlsx
```

## Exemplo de saída na planilha Excel

Veja como os dados são organizados na planilha após a execução do script:

![Exemplo de saída](https://github.com/user-attachments/assets/9202a144-1067-475d-885a-5e5a8f69b94a)

