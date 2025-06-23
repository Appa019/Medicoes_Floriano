🌤️ Medições Usina Geradora Floriano
🔗 Acesse o App: https://medicoesflorianocsn.streamlit.app

Sistema completo para processamento e análise de dados meteorológicos (.dat) da Usina Geradora Floriano, com atualização automatizada de planilhas Excel contendo estatísticas mensais e diárias.

📋 Índice
Visão Geral

Funcionalidades

Instalação

Como Usar

Formato dos Arquivos .dat

Estrutura do Código

Arquitetura

Deploy no Streamlit

Troubleshooting

Licença

🎯 Visão Geral
Este aplicativo transforma arquivos .dat em análises estruturadas:

Análises Mensais: estatísticas diárias de mín., máx., média e outliers.

Análises Diárias: detalhamento horário por dia.

Variáveis Monitoradas
Variável	Descrição	Unidade
Temperatura	Temperatura ambiente	°C
Piranômetro 1	Radiação solar (sensor 1)	kW/m²
Piranômetro 2	Radiação solar (sensor 2)	kW/m²
Piranômetro Albedo	Radiação solar refletida	kW/m²
Umidade Relativa	Umidade do ar	%
Velocidade do Vento	Velocidade do vento	m/s
Bateria	Tensão da bateria	V
Temperatura do Logger	Temperatura interna do logger	°C
Bateria Lítio	Tensão da bateria de lítio	V

✨ Funcionalidades
Interface
Design responsivo (desktop/mobile)

Cores institucionais CSN (Azul #00529C e Preto #231F20)

Feedback visual com barras de progresso

Processamento
Suporte a múltiplos arquivos .dat

Processamento paralelo (mensal e diário)

Validação automática e tratamento de erros

Estatísticas
Cálculo de mín., máx., média

Detecção de outliers (IQR)

Mapeamento automático em planilhas Excel

Exportação
Excel pronto para download

Nome inteligente com timestamp

Preservação da formatação original

🚀 Instalação
Pré-requisitos
Python 3.8+

pip instalado

Passos
bash
Copiar
Editar
# Clone o repositório
git clone <url-do-repositorio>
cd medicoes-floriano

# Instale as dependências
pip install -r requirements.txt

# Rode a aplicação
streamlit run app.py
Acesse: http://localhost:8501

📖 Como Usar
Passo 1: Upload do Excel Anual
O Excel deve conter abas como: 01-Analise Mensal, 01-Analise Diaria etc.

Passo 2: Upload dos Arquivos .dat
Selecione múltiplos arquivos no formato padrão.

Passo 3: Processar
Clique em "🚀 Processar Dados" e acompanhe os gráficos de progresso.

Passo 4: Download
Baixe o Excel atualizado com as abas preenchidas automaticamente.

Formato dos Arquivos .dat
Os arquivos devem conter cabeçalhos como:

arduino
Copiar
Editar
"TIMESTAMP","RECORD","Ane_Min",...,"LitBatt_Std"
"TS","RN","m/s",...,"Volts"
"","","Min",...,"Std"
Formato compatível com CR1000X / TOA5 (Campbell Scientific).

🏗️ Estrutura do Código
Arquivos
app.py: Interface + lógica principal

requirements.txt: Dependências

Classe: CompleteWeatherProcessor
Responsável por todo o processamento:

python
Copiar
Editar
def __init__(self):
    self.dados_processados = {}
    self.excel_path = None
Métodos Principais
process_dat_files: Lê e valida múltiplos .dat

_process_monthly_and_daily_data: Gera estatísticas e organiza por data

update_excel_file: Mapeia e escreve os dados nas abas corretas

🏛️ Arquitetura
Fluxo
Copiar
Editar
.dat → Pandas → Validação → Estatísticas → Excel → Download
Design
UI: Streamlit

Processamento: Classe isolada

Manipulação de dados: Pandas

Excel: OpenPyXL

🌐 Deploy no Streamlit
Envie para um repositório GitHub:

bash
Copiar
Editar
/medicoes-floriano
├── app.py
├── requirements.txt
└── README.md
Acesse https://share.streamlit.io

Conecte ao seu repositório e defina:

Branch: main

Main file: app.py

Variáveis de Ambiente (opcional)
toml
Copiar
Editar
# .streamlit/config.toml
[theme]
primaryColor = "#00529C"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F0F2F6"
textColor = "#262730"
🔧 Troubleshooting
Problema	Solução
Arquivo não carrega	Verifique extensão, tamanho e formatação do .dat
Abas não encontradas	Nomenclatura deve ser XX-Analise Mensal e XX-Analise Diaria
Dados não aparecem	Confirme formato de timestamp e integridade dos dados
Erro de memória	Divida os arquivos e monitore o uso de RAM

Ative logs com:

python
Copiar
Editar
import logging
logging.basicConfig(level=logging.DEBUG)
📄 Licença
Projeto desenvolvido para uso interno da Usina Geradora Floriano.
Todos os direitos reservados.
