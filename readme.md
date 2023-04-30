# Automail
### _Envios massivos de emails_  

## Sobre
Automail é uma automação de envios de emails em massa, escrito em python, utilizando templates feitos com HTML e CSS, e obtendo dados de uma planilha do Excel. Abaixo segue um manual rapido de como funciona:
- Os dados são obtidos com o preenchimento da planilha do Excel, sendo necessário salvar a planilha para os dados serem lidos corretamente.
- Pode-se separar em diferentes guias da planilha informações para cada "tema de envio" (exemplo marketing, vendas, financeiro) e no momento que o programa estiver em execução basta informar o nome da guia desejada
- A biblioteca usada pelo python para envio de emails, não permite que o CSS seja externo, sendo assim, utilize CSS interno para criar os templates, e em alguns casos será necessário usar o CSS inline
- O servidor de email usado nesse codigo é o do Google (Gmail), para cada conta de serviços de email, terá uma configuração diferente que permitirá usar serviços de SMTP 
&nbsp;
##  Como usar

1. Insira as informações na planilha e SALVE a planilha, pois os dados não serão capturados caso não salve. 
>Caso queira enviar um email com as mesmas informações para mais de um destinatário, basta colocar na coluna de email, os endereços separados por virgula

2. Se necessário edite o template e salve.
3. Execute o programa e insira as informações de email e senha (a senha aparentemente não estará digitando, mas isso é apenas uma medida de segurança). O programa exibirá se teve sucesso ao logar.
4. Informe o caminho do arquivo de template
5. Informe o caminho da planilha e o nome da guia
6. Pressione Enter e aguarde que o programa estará encaminhando os emails

ℹ️ É recomendável que envie para algum de seus emails para testar se o CSS está funcionando corretamente ℹ️


&nbsp;
## Criando um arquivo executável para Windows

Para criar um arquivo **.exe** instale a biblioteca pyinstaller:
```sh
pip install -U pyinstaller
```

Para salvar o arquivo como apenas um arquivo execute o comando abaixo:
```sh
pyinstaller certificatepy.py –windowed
```
Se quiser consulte a documentação completa do [pyinstaller](https://pyinstaller.org/en/stable/)
