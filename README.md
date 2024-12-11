
#  Scraper de Produtos - Americanas  

Um script Python para raspar dados de produtos no site da Americanas a partir de links fornecidos em uma planilha Excel. O programa coleta informações como descrição, preço, imagem e armazena os dados em um banco de dados Excel.

---

##  Funcionalidades  

- Raspagem automatizada de dados com **Selenium e GeckoDriver**.  
- Interface gráfica para carregar links através de planilhas Excel com o **Tkinter**.  
- Gerenciamento inteligente de dados: atualiza a planilha automaticamente caso encontre links repetidos.  
- Armazenamento seguro de dados coletados no formato **Excel**.

---

##  Tecnologias  

- **Python**  
- **Selenium**  
- **GeckoDriver**  
- **Tkinter**  
- **Pandas**  
- **Excel (.xlsx)**  

---

##  Pré-requisitos  

Certifique-se de ter as seguintes ferramentas instaladas no seu sistema:  

1. **Python** instalado (versão 3.x).  
2. **GeckoDriver** para o Firefox: [Baixe aqui](https://github.com/mozilla/geckodriver/releases).  
3. **Selenium**: Instale com o pip:  
   ```bash
   pip install selenium pandas openpyxl
   ```  
4. **Firefox** instalado em seu sistema.  

---

##  Configuração  

1. Ajuste os caminhos no código para refletirem onde você instalou o `geckodriver.exe` e o `firefox.exe`.  
   ```python
   driver_path = "C:/caminho_correto/geckodriver.exe"
   firefox_binary_path = "C:/Program Files/Mozilla Firefox/firefox.exe"
   ```  
2. Execute o script. Será exibida uma interface para carregar a planilha Excel com os links a serem processados.  

---

##  Como Usar  

1. Execute o script.  
2. Após abrir a interface, selecione a planilha que contém os links para raspagem (Excel no formato `.xlsx` ou `.xls`).  
3. Aguarde até que os dados sejam processados e armazenados no arquivo `produtos_americanas.xlsx`.  

---

##  Estrutura do Banco de Dados  

Os dados serão salvos no formato **Excel**, contendo as seguintes colunas:  

| **Campo**        | **Descrição**                                  |
|------------------|------------------------------------------------|
| id_produto       | Identificador do produto raspado.              |
| link             | Link do produto no site.                       |
| descrição        | Nome/descrição do produto.                     |
| preço            | Preço do produto.                              |
| imagem           | URL da imagem do produto.                      |
| loja             | Nome da loja (neste caso sempre "Americanas"). |

---

##  Diagnóstico  

O script verifica automaticamente:  

- Se o banco de dados já existe.  
- Se algum link já foi raspado anteriormente para evitar duplicação de dados.

---

##  Erros Comuns  

Caso você encontre erros:  
1. Certifique-se de que o **GeckoDriver**, o **Firefox** e o **Python Selenium** estão corretamente instalados.  
2. Verifique se o caminho configurado para o GeckoDriver e Firefox está correto no código.  

---

##  Contribuindo  

Caso deseje contribuir, faça um **Fork**, crie uma branch com suas melhorias e envie um **Pull Request** com suas mudanças.  

---

##  Licença  

Este projeto está licenciado sob a **MIT License**. Consulte o arquivo [LICENSE](LICENSE) para mais informações.



