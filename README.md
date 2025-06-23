🌤️ Medições Usina Geradora Floriano

https://medicoesflorianocsn.streamlit.app

Processador Completo de Dados Meteorológicos

Este aplicativo Streamlit foi desenvolvido para processar dados meteorológicos da Usina Geradora Floriano, realizando análises mensais e diárias de forma automatizada. O sistema processa arquivos de dados (.dat) e atualiza planilhas Excel com estatísticas detalhadas.

📋 Índice

•
Visão Geral

•
Funcionalidades

•
Instalação

•
Como Usar

•
Estrutura do Código

•
Arquitetura

•
Deploy no Streamlit

•
Troubleshooting

🎯 Visão Geral

O aplicativo foi projetado para automatizar o processamento de dados meteorológicos coletados por estações de monitoramento. Ele converte dados brutos em formato .dat em análises estatísticas organizadas, atualizando automaticamente planilhas Excel com:

•
Análises Mensais: Estatísticas diárias (mínimo, máximo, média, outliers)

•
Análises Diárias: Dados horários detalhados para cada dia do mês

Variáveis Monitoradas

O sistema processa as seguintes variáveis meteorológicas:

VariávelDescriçãoUnidadeTemperaturaTemperatura ambiente°CPiranômetro 1Radiação solar (sensor 1)kW/m²Piranômetro 2Radiação solar (sensor 2)kW/m²Piranômetro AlbedoRadiação solar refletidakW/m²Umidade RelativaUmidade do ar%Velocidade do VentoVelocidade do ventom/sBateriaTensão da bateriaVTemperatura do LoggerTemperatura interna do equipamento°CBateria LítioTensão da bateria de lítioV

✨ Funcionalidades

Interface do Usuário

•
Design Responsivo: Interface adaptável para desktop e mobile

•
Cores Corporativas: Utiliza as cores oficiais da CSN (USAFA Blue #00529C e Raisin Black #231F20)

•
Feedback Visual: Barras de progresso e mensagens de status em tempo real

•
Layout Intuitivo: Organização clara dos elementos da interface

Processamento de Dados

•
Upload Múltiplo: Suporte para múltiplos arquivos .dat simultaneamente

•
Validação Automática: Verificação da integridade dos dados

•
Processamento Paralelo: Análises mensais e diárias simultâneas

•
Tratamento de Erros: Gestão robusta de exceções e dados inválidos

Análises Estatísticas

•
Estatísticas Descritivas: Cálculo de mínimo, máximo, média para cada variável

•
Detecção de Outliers: Identificação automática usando método IQR (Interquartile Range)

•
Agregação Temporal: Dados horários agregados em estatísticas diárias

•
Mapeamento Automático: Localização automática das abas corretas no Excel

Exportação

•
Download Automático: Arquivo Excel atualizado disponível para download

•
Nomenclatura Inteligente: Arquivos nomeados com timestamp para controle de versão

•
Preservação de Formatação: Mantém a formatação original das planilhas

🚀 Instalação

Pré-requisitos

•
Python 3.8 ou superior

•
pip (gerenciador de pacotes Python)

Instalação Local

1.
Clone ou baixe os arquivos do projeto

Bash


# Se usando Git
git clone <url-do-repositorio>
cd medicoes-floriano

# Ou baixe os arquivos app.py e requirements.txt


1.
Instale as dependências

Bash


pip install -r requirements.txt


1.
Execute o aplicativo

Bash


streamlit run app.py


1.
Acesse no navegador

Plain Text


http://localhost:8501


Deploy no Streamlit Cloud

1.
Faça upload dos arquivos para um repositório GitHub

2.
Acesse share.streamlit.io

3.
Conecte seu repositório GitHub

4.
Configure o deploy apontando para app.py

5.
O Streamlit instalará automaticamente as dependências do requirements.txt

📖 Como Usar

Passo 1: Upload do Excel Anual

1.
Clique em "Selecione o arquivo Excel anual"

2.
Escolha o arquivo Excel que contém as abas de análise mensal e diária

3.
O arquivo deve seguir a estrutura padrão com abas nomeadas como "XX-Analise Mensal" e "XX-Analise Diaria"

Passo 2: Upload dos Arquivos .dat

1.
Clique em "Selecione os arquivos .dat"

2.
Selecione múltiplos arquivos .dat (use Ctrl+Click para seleção múltipla)

3.
Os arquivos devem estar no formato padrão da estação meteorológica

Passo 3: Processamento

1.
Clique no botão "🚀 Processar Dados"

2.
Acompanhe o progresso através das barras de status

3.
Visualize o resumo dos dados processados

Passo 4: Download

1.
Após o processamento bem-sucedido, clique em "📥 Baixar Excel Atualizado"

2.
O arquivo será baixado com timestamp no nome

3.
Verifique se todas as abas foram atualizadas corretamente

Formato dos Arquivos .dat

Os arquivos .dat devem seguir o formato padrão:

Plain Text


"TOA5","CR1000X","CR1000X","36521","CR1000X.Std.03.02","CPU:MEDICOES_FLORIANO_V1.CR1X","59853","Table1"
"TIMESTAMP","RECORD","Ane_Min","Ane_Max","Ane_Avg","Ane_Std","Temp_Min","Temp_Max","Temp_Avg","Temp_Std","RH_Min","RH_Max","RH_Avg","RH_Std","Pir1_Min","Pir1_Max","Pir1_Avg","Pir1_Std","Pir2_Min","Pir2_Max","Pir2_Avg","Pir2_Std","PirALB_Min","PirALB_Max","PirALB_Avg","PirALB_Std","Batt_Min","Batt_Max","Batt_Avg","Batt_Std","LoggTemp_Min","LoggTemp_Max","LoggTemp_Avg","LoggTemp_Std","LitBatt_Min","LitBatt_Max","LitBatt_Avg","LitBatt_Std"
"TS","RN","m/s","m/s","m/s","m/s","Deg C","Deg C","Deg C","Deg C","%","%","%","%","W/m^2","W/m^2","W/m^2","W/m^2","W/m^2","W/m^2","W/m^2","W/m^2","W/m^2","W/m^2","W/m^2","W/m^2","Volts","Volts","Volts","Volts","Deg C","Deg C","Deg C","Deg C","Volts","Volts","Volts","Volts"
"","","Min","Max","Avg","Std","Min","Max","Avg","Std","Min","Max","Avg","Std","Min","Max","Avg","Std","Min","Max","Avg","Std","Min","Max","Avg","Std","Min","Max","Avg","Std","Min","Max","Avg","Std","Min","Max","Avg","Std"


🏗️ Estrutura do Código

Arquivos Principais

app.py

Arquivo principal do aplicativo Streamlit contendo:

•
Interface do usuário

•
Lógica de processamento

•
Classe CompleteWeatherProcessor

•
Funções de visualização

requirements.txt

Lista de dependências Python:

Plain Text


streamlit>=1.28.0
pandas>=2.0.0
numpy>=1.24.0
openpyxl>=3.1.0


Classe Principal: CompleteWeatherProcessor

A classe CompleteWeatherProcessor é o núcleo do sistema, responsável por:

Inicialização (__init__)

Python


def __init__(self):
    self.dados_processados = {}  # Armazena dados processados
    self.excel_path = None       # Caminho do arquivo Excel
    self.abas_mensais_atualizadas = []  # Lista de abas mensais atualizadas
    self.abas_diarias_atualizadas = []  # Lista de abas diárias atualizadas


Mapeamento de Colunas

Python


self.column_mapping = {
    'Temperatura': {'start_num': 2},
    'Piranometro_1': {'start_num': 33},
    'Piranometro_2': {'start_num': 64},
    'Piranometro_Alab': {'start_num': 95},
    'Umidade_Relativa': {'start_num': 126},
    'Velocidade_Vento': {'start_num': 157}
}


Métodos Principais

process_dat_files(dat_files)

Processa múltiplos arquivos .dat:

1.
Leitura: Lê cada arquivo .dat

2.
Parsing: Converte dados CSV em DataFrame

3.
Validação: Verifica integridade dos dados

4.
Processamento: Chama métodos de análise mensal e diária

_process_monthly_and_daily_data(data)

Processa dados para ambas as análises:

1.
Identificação: Determina mês e ano dos dados

2.
Agregação Diária: Calcula estatísticas por dia

3.
Agregação Horária: Organiza dados por hora

4.
Armazenamento: Salva em estrutura de dados organizada

_calculate_daily_statistics(data)

Calcula estatísticas diárias:

Python


stats[var] = {
    'min': data[f'{var}_Min'].min(),
    'max': data[f'{var}_Max'].max(),
    'avg': data[f'{var}_Avg'].mean(),
    'outliers': self._count_outliers(data, var)
}


_count_outliers(data, variable)

Detecta outliers usando método IQR:

1.
Quartis: Calcula Q1 e Q3

2.
IQR: Calcula Interquartile Range

3.
Limites: Define limites inferior e superior

4.
Contagem: Conta valores fora dos limites

update_excel_file(excel_file)

Atualiza arquivo Excel:

1.
Carregamento: Abre arquivo Excel temporariamente

2.
Localização: Encontra abas corretas

3.
Atualização: Escreve dados nas células apropriadas

4.
Salvamento: Salva alterações

Mapeamento de Células Excel

Análise Mensal

O sistema mapeia dados para células específicas:

Primeira Seção (linhas 3-33):

•
Temperatura: Colunas B-E

•
Piranômetro 1: Colunas H-K

•
Piranômetro 2: Colunas N-Q

•
Piranômetro Albedo: Colunas T-W

•
Umidade Relativa: Colunas Z-AC

Segunda Seção (linhas 37-67):

•
Velocidade do Vento: Colunas B-E

•
Bateria: Colunas H-K

•
Bateria Lítio: Colunas N-Q

•
Temperatura Logger: Colunas T-W

Análise Diária

Dados horários são mapeados em colunas sequenciais:

•
Cada dia ocupa uma coluna

•
Cada hora ocupa uma linha (3-26)

•
Variáveis são agrupadas em blocos de colunas

🏛️ Arquitetura

Fluxo de Dados

Plain Text


Arquivos .dat → Parsing → Validação → Processamento → Excel → Download
     ↓              ↓         ↓            ↓          ↓        ↓
  Upload UI    DataFrame   Verificação   Análises   Update   Usuário


Estrutura de Dados

Python


dados_processados = {
    "2024-01": {
        "monthly_data": {
            1: {  # Dia 1
                "Temp": {"min": 20.1, "max": 35.2, "avg": 27.5, "outliers": 2},
                "Pir1": {"min": 0, "max": 1200, "avg": 450, "outliers": 0},
                # ... outras variáveis
            },
            # ... outros dias
        },
        "daily_data": {
            1: {  # Dia 1
                "00:00": {"Temperatura": 22.1, "Piranometro_1": 0, ...},
                "01:00": {"Temperatura": 21.8, "Piranometro_1": 0, ...},
                # ... outras horas
            },
            # ... outros dias
        }
    },
    # ... outros meses
}


Padrões de Design

Separação de Responsabilidades

•
Interface: Streamlit gerencia UI e interação

•
Processamento: Classe dedicada para lógica de negócio

•
Dados: Pandas para manipulação de dados

•
Excel: OpenPyXL para operações de planilha

Tratamento de Erros

Python


try:
    # Operação principal
    result = process_data()
except SpecificException as e:
    # Tratamento específico
    handle_specific_error(e)
except Exception as e:
    # Tratamento genérico
    handle_general_error(e)


Feedback do Usuário

•
Progress bars para operações longas

•
Mensagens de status em tempo real

•
Códigos de cores para diferentes tipos de mensagem

🌐 Deploy no Streamlit

Configuração do Repositório GitHub

1.
Estrutura de Arquivos:

Plain Text


medicoes-floriano/
├── app.py
├── requirements.txt
└── README.md


1.
Configuração no Streamlit Cloud:

•
Repository: Seu repositório GitHub

•
Branch: main (ou master)

•
Main file path: app.py

•
Python version: 3.8+



Variáveis de Ambiente

O aplicativo não requer variáveis de ambiente especiais, mas você pode configurar:

Plain Text


# .streamlit/config.toml (opcional)
[theme]
primaryColor = "#00529C"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F0F2F6"
textColor = "#262730"


Otimizações para Produção

Cache de Dados

Python


@st.cache_data
def load_data(file):
    return pd.read_csv(file)


Gestão de Memória

•
Limpeza automática de arquivos temporários

•
Uso eficiente de DataFrames

•
Garbage collection explícito quando necessário

🔧 Troubleshooting

Problemas Comuns

1. Erro de Upload de Arquivo

Sintoma: Arquivo não carrega ou erro de formato
Solução:

•
Verifique se o arquivo .dat está no formato correto

•
Confirme que o Excel tem as abas com nomenclatura padrão

•
Verifique o tamanho do arquivo (limite do Streamlit)

2. Abas não Encontradas

Sintoma: Mensagem "Aba não encontrada"
Solução:

•
Verifique nomenclatura das abas: "XX-Analise Mensal" e "XX-Analise Diaria"

•
Confirme que XX corresponde ao mês (01, 02, ..., 12)

•
Verifique se não há espaços extras nos nomes

3. Dados não Processados

Sintoma: Nenhum dado aparece após processamento
Solução:

•
Verifique formato do timestamp nos arquivos .dat

•
Confirme que os dados não estão corrompidos

•
Verifique se há dados suficientes (mínimo 1 dia)

4. Erro de Memória

Sintoma: Aplicativo trava com arquivos grandes
Solução:

•
Processe arquivos menores por vez

•
Verifique recursos disponíveis no servidor

•
Considere otimizações de código

Logs e Debugging

Ativando Logs Detalhados

Python


import logging
logging.basicConfig(level=logging.DEBUG)


Verificação de Dados

Python


# Adicione prints para debug
print(f"Dados processados: {len(self.dados_processados)}")
print(f"Colunas encontradas: {data.columns.tolist()}")


Limitações Conhecidas

1.
Tamanho de Arquivo: Limitado pela configuração do Streamlit Cloud

2.
Formato de Data: Requer formato específico nos arquivos .dat

3.
Nomenclatura de Abas: Deve seguir padrão exato

4.
Encoding: Arquivos devem estar em UTF-8

📞 Suporte

Para problemas técnicos ou dúvidas sobre o uso:

1.
Verifique este README para soluções comuns

2.
Consulte os logs do aplicativo para erros específicos

3.
Teste com arquivos menores para isolar problemas

4.
Verifique a documentação do Streamlit para questões de deploy

📄 Licença

Este projeto foi desenvolvido para uso interno da Usina Geradora Floriano. Todos os direitos reservados.







# Medicoes_Floriano
Python p/ Atualizacao de Medicoes
