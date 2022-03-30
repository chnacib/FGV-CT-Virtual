# FGV-CT_Virtual

![Screenshot_6](https://user-images.githubusercontent.com/84737230/160890478-ac108b8e-929c-482a-9d66-8526361f3cd4.png)

Consulta automatizada de dados na web e envio de mensagem automatizada no whatsapp 

## 💻 Pré-Requisitos

Antes de começar, verifique se você atendeu aos seguintes requisitos:
* Você instalou a versão mais recente de `Python`
* Você tem uma máquina `Windows`
* O diretório deve estar em: `C:\CertPreponaWpp\`
* O `chromedriver` deve ser compatível com a versão do google chrome

## 🤖 Usando o programa

Clique no executável ou digite o comando no seu terminal:
```
python <caminho>\<programa>
```
### Configurando inicialização

A primeira coisa a ser feita no programa é selecionar a data que você quer extrair os dados e apertar em "Get date". A data foi salva com sucesso ao exibir a mensagem "Saved".

![date_get](https://user-images.githubusercontent.com/84737230/160892823-8c1090c3-f4a0-49e0-8902-18f35651b69a.gif)

O próximo passo é importar a planilha `Modelo.xlsx` clicando no botão "Import sheet". Uma janela será aberta no Windows Explorer para a seleção da planilha.

⚠️⚠️⚠️ Somente a planilha `Modelo.xlsx` será aceita para a execução do bot. ⚠️⚠️⚠️

![import_sheet](https://user-images.githubusercontent.com/84737230/160893653-8a570be6-512b-4bfa-999a-04f4a6731fbc.gif)


Quando tu estiver configurado, aperte o botão "Start" e aguarde o processo.

![Screenshot_5](https://user-images.githubusercontent.com/84737230/160894003-8d65b1b6-5d42-4d3c-8a2b-00296771fff5.png)

### Monitoramento

Você poderá monitorar a execução do bot de acordo com a mensagem exibida abaixo do botão START.

![monitoramento](https://user-images.githubusercontent.com/84737230/160895420-d26d92e9-fdb4-48a8-ae93-4dd1e5eafbcd.gif)

A Barrinha de carregamento determina a quantidade de candidatos coletados na Prepona
```
start_value = 100/len(listacpf)
progress_value = start_value
```
        
![loading](https://user-images.githubusercontent.com/84737230/160894298-35d04aef-d7f7-4828-8e77-6d7dd8f41cb7.gif)

Ao terminar todo o processo dentro da Prepona e CertPessoas, a próxima parte é inserir o QRCODE do Whatsapp Business dentro do navegador.
Será apresentada a seguinte mensagem quando for a hora de inserir:

![Screenshot_7](https://user-images.githubusercontent.com/84737230/160896605-5b25812a-ca4e-4aab-9c04-3d69346ea085.png)

#### Lembrando que o programa fica "pausado" enquanto aguarda o QRCODE, então sem correria kkkkk

### Planilha de dados

Ao final da execução do programa, irá aparecer uma mensagem de "Process Finished" abaixo do botão START, e o navegador fechará automaticamente.

![Screenshot_8](https://user-images.githubusercontent.com/84737230/160902442-93e055da-7ff5-4f28-ae46-57d711ecf9d0.png)

Após isso, verifique a pasta `C:\CertPreponaWpp\` ou o local onde está o script caso execute o programa pelo terminal e procure por um documento com nome de `ResultadoPrepona.xlsx`

![resultado_arquivo](https://user-images.githubusercontent.com/84737230/160903342-8a9ee520-1a36-49dd-afe0-18eb8017a319.gif)

Dentro desse arquivo estará as informações coletadas pelo bot(Nome,cpf,data,hora,hardware,instalação,RAM e etc)

![Screenshot_9](https://user-images.githubusercontent.com/84737230/160904513-aea98605-fb10-4ec7-ac7d-cb913a980110.png)

