# BOT-SERVICO
Automação em Python para o lançamento de notas fiscais de serviço no sistema TOTVS.

## 💡 Objetivo

Este projeto automatiza o processo de análise e registro de notas fiscais de serviço emitidas. Ele atua diretamente na interface do TOTVS, realizando o preenchimento de campos com base nos dados extraídos dos arquivos XMLs das notas.

## 🚀 Funcionalidades

- Extração de informações de XMLs.
- Análise e categorização das notas.
- Busca automática de arquivos em pastas padronizadas.
- Interação com sistema TOTVS para preenchimento automatizado.
- Banco de dados local para controle e rastreabilidade.
- Interface gráfica simples para controle da automação.
- Exportável como `.exe` para execução sem dependências externas.

## 📁 Estrutura de Pastas
BOT_SERVICES/
│<br>
├── app/ # Interface gráfica (controle da automação)<br>
├── build/ # Pasta gerada pelo PyInstaller<br>
├── dist/ # Executável e bancos de dados locais<br>
├── env/ # Ambiente virtual (excluído pelo .gitignore)<br>
├── icons/ # Ícones usados na aplicação<br>
├── path/ # JSONs e configurações de caminhos<br>
├── processos/ # Scripts principais de automação (análise, extração, interação web)<br>
├── utils/ # Funções auxiliares (ex: serviços.py)<br>
├── web/ # Módulos relacionados à automação web<br>
├── .gitignore # Arquivos/pastas ignoradas pelo Git<br>
├── KADRIX S.spec # Configuração do PyInstaller<br>
├── README.md # Este arquivo<br>
├── requirements.txt # Dependências do projeto<br>


## ⚙️ Tecnologias Utilizadas

- Python 3.x
- OpenPyXL
- PyAutoGUI / Pyperclip
- Selenium
- Tkinter (CustomTkinter)
- SQLite

## 🧪 Como Executar

1. Crie e ative um ambiente virtual:
   ```bash
   python -m venv env
   source env/bin/activate  # ou .\env\Scripts\activate no Windows

## Instale as dependências:
> pip install -r requirements.txt

## Execute o script principal:
> python -m app.app

## Para criar o .exe:
> pyinstaller --onefile -w --icon=icons/"ICONE CRIADO".ico --name="NOME DA SUA ESCOLHA" app/app.py

## ⚠️ Avisos
- Este projeto é uma versão genérica, sem qualquer vínculo com dados sensíveis ou proprietários. Adaptado exclusivamente para fins educacionais e de portfólio.
