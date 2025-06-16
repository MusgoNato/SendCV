# SENDCV

- SENDCV é uma aplicação Python que automatiza o envio de currículos via Microsoft Outlook para uma lista de contatos de e-mail. Ideal para profissionais e estudantes que desejam aplicar rapidamente para vagas de estágio ou posições júnior, economizando tempo e padronizando a comunicação com recrutadores.

## Estrutura do projeto

```
sendcv/
│
├── main.py              # Script principal
├── .env                 # Suas informações pessoais
├── emails.txt           # Lista de e-mails
├── curriculo.pdf        # Anexo a ser enviado
├── requirements.txt     # Dependências
└── README.md            # Instruções do projeto
```

## Funcionalidades
- Envio de e-mails personalizados com assunto e corpo definidos
- Inclusão automática de anexo (currículo)
- Leitura de contatos de um arquivo .txt
- Intervalo entre os envios para evitar spam
- Variáveis de ambiente para dados pessoais (nome, linkedin, etc)

## Pré-Requisitos
- Python 3.8+ instalado
- Microsoft Outlook instalado e configurado (com conta logada)
- Sistema operacional: Windows (uso do win32com.client)

## Instalação e Configuração
#### Clone o repositório (ou baixe os arquivos)
```
git clone https://github.com/MusgoNato/sendcv.git
cd sendcv
```

#### Crie um ambiente virtual (opcional, mas recomendado)
```
python -m venv venv
venv\Scripts\activate
```
-Instale as dependências
```
pip install -r requirements.txt
```

#### Configure suas variáveis pessoais
```
NAME=Seu Nome completo
LINKEDIN=link do seu linkedin
GITHUB=link do seu github
PHONE=Seu numero de telefone para contato
```
#### Adicione seu currículo na pasta raiz
Nomeie o arquivo como: `curriculum.pdf`

#### Crie um arquivo com os e-mails
Crie um arquivo chamado `emails.txt` e coloque os e-mails separados por ´;´, assim:
```
email1@exemplo.com;email2@exemplo.com;email3@exemplo.com
```

## Como usar?
Execute o script principal:
```
python main.py
```
## Atenção
- O outlook pode bloquear ou pedir confirmação para envio - fique atento durante o processo.
- Evite enviar muito e-mails muito rápido para não cair em spam ou ser bloqueado.
- Sempre revise os e-mails antes de rodar o script em massa.