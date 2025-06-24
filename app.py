import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os
from datetime import datetime
import warnings
import io
import tempfile
import zipfile

warnings.filterwarnings('ignore')

# Configuração da página
st.set_page_config(
    page_title="Medições Usina Geradora Floriano",
    page_icon="🌤️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado com as cores da CSN
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #00529C 0%, #231F20 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
    }
    
    .main-header h1 {
        color: white;
        font-size: 2.5rem;
        margin: 0;
        font-weight: bold;
    }
    
    .main-header p {
        color: #E8E8E8;
        font-size: 1.2rem;
        margin: 0.5rem 0 0 0;
    }
    
    .stButton > button {
        background-color: #00529C;
        color: white;
        border: none;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        font-weight: bold;
        transition: all 0.3s;
    }
    
    .stButton > button:hover {
        background-color: #231F20;
        transform: translateY(-2px);
    }
    
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .sidebar .sidebar-content {
        background-color: #f8f9fa;
    }
    
    .metric-card {
        background-color: white;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #00529C;
        margin: 0.5rem 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

class CompleteWeatherProcessor:
    """
    Processador completo de dados meteorológicos
    Realiza automaticamente análises mensais E diárias com dados reais por data
    """

    def __init__(self):
        self.dados_processados = {}
        self.excel_path = None
        self.abas_mensais_atualizadas = []
        self.abas_diarias_atualizadas = []

        # Mapeamento de meses
        self.meses = {
            1: "01", 2: "02", 3: "03", 4: "04",
            5: "05", 6: "06", 7: "07", 8: "08",
            9: "09", 10: "10", 11: "11", 12: "12"
        }

        # Mapeamento de colunas para análise diária
        self.column_mapping = {
            'Temperatura': {'start_num': 2},
            'Piranometro_1': {'start_num': 33},
            'Piranometro_2': {'start_num': 64},
            'Piranometro_Alab': {'start_num': 95},
            'Umidade_Relativa': {'start_num': 126},
            'Velocidade_Vento': {'start_num': 157}
        }

    def process_dat_files(self, dat_files):
        """
        Processa múltiplos arquivos .dat para ambas as análises
        """
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_files = len(dat_files)
        self.file_processing_info = []  # Lista para armazenar info de cada arquivo
        
        for i, uploaded_file in enumerate(dat_files):
            status_text.text(f"Processando arquivo {i+1}/{total_files}: {uploaded_file.name}")
            
            try:
                # Ler arquivo .dat usando pandas diretamente (como no Colab)
                uploaded_file.seek(0)  # Reset file pointer
                data = pd.read_csv(uploaded_file, skiprows=4, parse_dates=[0])

                # Contar registros
                total_records = len(data)
                
                # Renomear colunas (igual ao código original do Colab)
                data.columns = [
                    'TIMESTAMP', 'RECORD',
                    'Ane_Min', 'Ane_Max', 'Ane_Avg', 'Ane_Std',
                    'Temp_Min', 'Temp_Max', 'Temp_Avg', 'Temp_Std',
                    'RH_Min', 'RH_Max', 'RH_Avg', 'RH_Std',
                    'Pir1_Min', 'Pir1_Max', 'Pir1_Avg', 'Pir1_Std',
                    'Pir2_Min', 'Pir2_Max', 'Pir2_Avg', 'Pir2_Std',
                    'PirALB_Min', 'PirALB_Max', 'PirALB_Avg', 'PirALB_Std',
                    'Batt_Min', 'Batt_Max', 'Batt_Avg', 'Batt_Std',
                    'LoggTemp_Min', 'LoggTemp_Max', 'LoggTemp_Avg', 'LoggTemp_Std',
                    'LitBatt_Min', 'LitBatt_Max', 'LitBatt_Avg', 'LitBatt_Std'
                ]

                data.set_index('TIMESTAMP', inplace=True)
                
                # Obter informações do período
                start_date = data.index.min()
                end_date = data.index.max()
                days_span = (end_date - start_date).days + 1
                
                # Processar para análises mensais E diárias
                processed_days = self._process_monthly_and_daily_data(data)
                
                # Armazenar informações do arquivo
                file_info = {
                    'arquivo': uploaded_file.name,
                    'registros': total_records,
                    'periodo_inicio': start_date.strftime('%Y-%m-%d %H:%M'),
                    'periodo_fim': end_date.strftime('%Y-%m-%d %H:%M'),
                    'dias_span': days_span,
                    'dias_processados': processed_days,
                    'status': '✅ Processado'
                }
                self.file_processing_info.append(file_info)
                
                # Mostrar progresso detalhado
                st.success(f"✅ {uploaded_file.name}: {total_records} registros, {processed_days} dias processados")

            except Exception as e:
                # Armazenar informações de erro
                error_info = {
                    'arquivo': uploaded_file.name,
                    'registros': 0,
                    'periodo_inicio': 'N/A',
                    'periodo_fim': 'N/A',
                    'dias_span': 0,
                    'dias_processados': 0,
                    'status': f'❌ Erro: {str(e)}'
                }
                self.file_processing_info.append(error_info)
                st.error(f"❌ Erro ao processar {uploaded_file.name}: {str(e)}")
                continue
            
            progress_bar.progress((i + 1) / total_files)

        status_text.text("Processamento concluído!")
        
        # Mostrar resumo detalhado dos arquivos processados
        self._show_file_processing_summary()
        
        if self.dados_processados:
            return True
        else:
            return False

    def _show_file_processing_summary(self):
        """Mostra resumo detalhado do processamento de cada arquivo"""
        if hasattr(self, 'file_processing_info') and self.file_processing_info:
            st.markdown("---")
            st.markdown("### 📄 Resumo do Processamento por Arquivo")
            
            # Criar DataFrame com as informações
            df_files = pd.DataFrame(self.file_processing_info)
            
            # Calcular totais
            total_records = df_files['registros'].sum()
            total_files_success = len([f for f in self.file_processing_info if '✅' in f['status']])
            total_files_error = len([f for f in self.file_processing_info if '❌' in f['status']])
            
            # Mostrar métricas gerais
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>📁 Arquivos Processados</h4>
                    <h2>{total_files_success}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>❌ Arquivos com Erro</h4>
                    <h2>{total_files_error}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>📊 Total de Registros</h4>
                    <h2>{total_records:,}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            # Tabela detalhada
            st.markdown("#### 📋 Detalhes por Arquivo")
            
            # Renomear colunas para exibição
            df_display = df_files.copy()
            df_display.columns = [
                'Arquivo', 'Registros', 'Início', 'Fim', 
                'Dias (Span)', 'Dias Processados', 'Status'
            ]
            
            # Formatar números
            df_display['Registros'] = df_display['Registros'].apply(lambda x: f"{x:,}" if x > 0 else "0")
            
            st.dataframe(df_display, use_container_width=True)

    def _process_monthly_and_daily_data(self, data):
        """
        Processa dados para análises mensais E diárias simultaneamente
        NOVA LÓGICA: Divide dados por data real (um arquivo .dat pode ter dados de múltiplos dias)
        """
        # CORREÇÃO: Processar por data real de cada registro
        data['date'] = data.index.date
        days_processed = 0
        
        # Agrupar dados por data real
        for date in data['date'].unique():
            day_data = data[data['date'] == date]
            
            # Extrair ano, mês e dia da data real
            ano = date.year
            mes_numero = date.month
            dia_numero = date.day
            dataset_key = f"{ano}-{mes_numero:02d}"

            # Inicializar estrutura se não existir
            if dataset_key not in self.dados_processados:
                self.dados_processados[dataset_key] = {
                    'monthly_data': {},  # Para análise mensal
                    'daily_data': {}     # Para análise diária
                }

            # Estatísticas diárias para análise mensal
            stats = self._calculate_daily_statistics(day_data)
            self.dados_processados[dataset_key]['monthly_data'][dia_numero] = stats

            # Dados horários reais (apenas horas que existem nos dados)
            hourly_data_real = self._process_real_hourly_data(day_data)
            if dia_numero not in self.dados_processados[dataset_key]['daily_data']:
                self.dados_processados[dataset_key]['daily_data'][dia_numero] = {}
            self.dados_processados[dataset_key]['daily_data'][dia_numero] = hourly_data_real
            
            days_processed += 1
        
        return days_processed

    def _process_real_hourly_data(self, day_data):
        """
        NOVA FUNÇÃO: Processa apenas dados horários reais (sem preencher 24h artificialmente)
        Regras:
        - Registros de 10:00-10:50 = hora 10:00
        - Média dos registros disponíveis na hora
        - Apenas horas com dados reais nos arquivos .dat
        """
        day_data['hour'] = day_data.index.hour
        day_data['minute'] = day_data.index.minute
        
        # Dicionário para armazenar apenas dados reais
        hourly_real = {}

        # Encontrar todas as horas que realmente existem nos dados
        available_hours = day_data['hour'].unique()
        
        for hour in sorted(available_hours):
            # Filtrar registros da hora atual (00:00 a 00:50, 01:00 a 01:50, etc.)
            hour_records = day_data[
                (day_data['hour'] == hour) & 
                (day_data['minute'].isin([0, 10, 20, 30, 40, 50]))
            ]

            if len(hour_records) > 0:
                # Calcular médias dos registros disponíveis na hora
                hourly_real[f"{hour:02d}:00"] = {
                    'Temperatura': round(hour_records['Temp_Avg'].mean(), 2),
                    'Piranometro_1': round(hour_records['Pir1_Avg'].mean() / 1000, 3),
                    'Piranometro_2': round(hour_records['Pir2_Avg'].mean() / 1000, 3),
                    'Piranometro_Alab': round(hour_records['PirALB_Avg'].mean() / 1000, 3),
                    'Umidade_Relativa': round(hour_records['RH_Avg'].mean(), 2),
                    'Velocidade_Vento': round(hour_records['Ane_Avg'].mean(), 2)
                }

        return hourly_real

    def _calculate_daily_statistics(self, data):
        """Calcula estatísticas diárias para análise mensal"""
        stats = {}
        variables = ['Temp', 'Pir1', 'Pir2', 'PirALB', 'RH', 'Ane', 'Batt', 'LoggTemp', 'LitBatt']

        for var in variables:
            stats[var] = {
                'min': data[f'{var}_Min'].min(),
                'max': data[f'{var}_Max'].max(),
                'avg': data[f'{var}_Avg'].mean(),
                'outliers': self._count_outliers(data, var)
            }

        return stats

    def _count_outliers(self, data, variable):
        """Conta outliers usando método IQR"""
        series = data[f'{variable}_Avg'].dropna()
        if len(series) == 0:
            return 0

        q1 = series.quantile(0.25)
        q3 = series.quantile(0.75)
        iqr = q3 - q1
        lower_bound = q1 - 1.5 * iqr
        upper_bound = q3 + 1.5 * iqr

        outliers = series[(series < lower_bound) | (series > upper_bound)]
        return len(outliers)

    def update_excel_file(self, excel_file):
        """
        Atualiza automaticamente análises mensais E diárias no Excel
        """
        if not self.dados_processados:
            return False, "Nenhum dado processado!"

        try:
            # Salvar arquivo Excel temporariamente
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(excel_file.read())
                self.excel_path = tmp_file.name

            wb = load_workbook(self.excel_path)
            
            sucesso_mensal = 0
            sucesso_diario = 0
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            total_months = len(self.dados_processados)
            
            # Processar cada mês
            for i, (dataset_key, month_data) in enumerate(self.dados_processados.items()):
                ano, mes = dataset_key.split('-')
                mes_numero = int(mes)
                
                status_text.text(f"Atualizando mês {mes}/{ano}...")

                # ANÁLISE MENSAL
                aba_mensal = self._find_sheet(wb.sheetnames, mes_numero, "Mensal")
                if aba_mensal:
                    try:
                        ws_mensal = wb[aba_mensal]
                        dias_mensal = self._update_monthly_data(ws_mensal, month_data['monthly_data'])

                        if aba_mensal not in self.abas_mensais_atualizadas:
                            self.abas_mensais_atualizadas.append(aba_mensal)
                        sucesso_mensal += dias_mensal
                    except Exception as e:
                        return False, f"Erro na análise mensal: {e}"

                # ANÁLISE DIÁRIA (NOVA LÓGICA COM DADOS REAIS)
                aba_diaria = self._find_sheet(wb.sheetnames, mes_numero, "Diaria")
                if aba_diaria:
                    try:
                        ws_diaria = wb[aba_diaria]
                        dias_diario = self._update_daily_data_real(ws_diaria, month_data['daily_data'])

                        if aba_diaria not in self.abas_diarias_atualizadas:
                            self.abas_diarias_atualizadas.append(aba_diaria)
                        sucesso_diario += dias_diario
                    except Exception as e:
                        return False, f"Erro na análise diária: {e}"
                
                progress_bar.progress((i + 1) / total_months)

            # Salvar alterações
            wb.save(self.excel_path)
            status_text.text("Atualização concluída!")

            if sucesso_mensal > 0 and sucesso_diario > 0:
                return True, f"Sucesso! Análise Mensal: {sucesso_mensal} dias, Análise Diária: {sucesso_diario} dias (apenas dados reais)"
            else:
                return False, "Nenhum dado foi atualizado"

        except Exception as e:
            return False, f"Erro geral: {e}"

    def _find_sheet(self, sheet_names, mes_numero, tipo):
        """Encontra aba mensal ou diária"""
        mes_str = f"{mes_numero:02d}"

        possible_names = [
            f"{mes_str}-Analise {tipo}",
            f"{mes_str}-Analyse {tipo}",
            f"{mes_str} Analise {tipo}",
            f"Analise {tipo} {mes_str}"
        ]

        for name in possible_names:
            if name in sheet_names:
                return name

        # Buscar por padrão
        for sheet_name in sheet_names:
            if mes_str in sheet_name and tipo in sheet_name:
                return sheet_name

        return None

    def _update_monthly_data(self, ws, monthly_data):
        """Atualiza dados da análise mensal"""
        dias_atualizados = 0

        for dia_numero, stats in monthly_data.items():
            # Primeira seção (linhas 3-33)
            target_row = dia_numero + 2

            # Temperatura
            ws[f'B{target_row}'] = round(stats['Temp']['min'], 2)
            ws[f'C{target_row}'] = round(stats['Temp']['max'], 2)
            ws[f'D{target_row}'] = round(stats['Temp']['avg'], 2)
            ws[f'E{target_row}'] = int(stats['Temp']['outliers'])

            # Piranômetro 1 (KW)
            ws[f'H{target_row}'] = round(stats['Pir1']['min'] / 1000, 2)
            ws[f'I{target_row}'] = round(stats['Pir1']['max'] / 1000, 2)
            ws[f'J{target_row}'] = round(stats['Pir1']['avg'] / 1000, 2)
            ws[f'K{target_row}'] = int(stats['Pir1']['outliers'])

            # Piranômetro 2 (KW)
            ws[f'N{target_row}'] = round(stats['Pir2']['min'] / 1000, 2)
            ws[f'O{target_row}'] = round(stats['Pir2']['max'] / 1000, 2)
            ws[f'P{target_row}'] = round(stats['Pir2']['avg'] / 1000, 2)
            ws[f'Q{target_row}'] = int(stats['Pir2']['outliers'])

            # Piranômetro ALB (KW)
            ws[f'T{target_row}'] = round(stats['PirALB']['min'] / 1000, 2)
            ws[f'U{target_row}'] = round(stats['PirALB']['max'] / 1000, 2)
            ws[f'V{target_row}'] = round(stats['PirALB']['avg'] / 1000, 2)
            ws[f'W{target_row}'] = int(stats['PirALB']['outliers'])

            # Umidade Relativa
            ws[f'Z{target_row}'] = round(stats['RH']['min'], 2)
            ws[f'AA{target_row}'] = round(stats['RH']['max'], 2)
            ws[f'AB{target_row}'] = round(stats['RH']['avg'], 2)
            ws[f'AC{target_row}'] = int(stats['RH']['outliers'])

            # Segunda seção (linhas 37-67)
            target_row_2 = dia_numero + 36

            # Velocidade do Vento
            ws[f'B{target_row_2}'] = round(stats['Ane']['min'], 2)
            ws[f'C{target_row_2}'] = round(stats['Ane']['max'], 2)
            ws[f'D{target_row_2}'] = round(stats['Ane']['avg'], 2)
            ws[f'E{target_row_2}'] = int(stats['Ane']['outliers'])

            # Bateria
            ws[f'H{target_row_2}'] = round(stats['Batt']['min'], 2)
            ws[f'I{target_row_2}'] = round(stats['Batt']['max'], 2)
            ws[f'J{target_row_2}'] = round(stats['Batt']['avg'], 2)
            ws[f'K{target_row_2}'] = int(stats['Batt']['outliers'])

            # LitBat
            ws[f'N{target_row_2}'] = round(stats['LitBatt']['min'], 2)
            ws[f'O{target_row_2}'] = round(stats['LitBatt']['max'], 2)
            ws[f'P{target_row_2}'] = round(stats['LitBatt']['avg'], 2)
            ws[f'Q{target_row_2}'] = int(stats['LitBatt']['outliers'])

            # LogTemp
            ws[f'T{target_row_2}'] = round(stats['LoggTemp']['min'], 2)
            ws[f'U{target_row_2}'] = round(stats['LoggTemp']['max'], 2)
            ws[f'V{target_row_2}'] = round(stats['LoggTemp']['avg'], 2)
            ws[f'W{target_row_2}'] = int(stats['LoggTemp']['outliers'])

            dias_atualizados += 1

        return dias_atualizados

    def _update_daily_data_real(self, ws, daily_data):
        """
        ATUALIZADA: Atualiza dados da análise diária usando apenas dados reais
        (sem preencher 24h artificialmente)
        """
        dias_atualizados = 0

        for dia_numero, day_hourly_data in daily_data.items():
            # Processar apenas as horas que realmente existem nos dados
            for hour_str, hour_data in day_hourly_data.items():
                # Extrair número da hora (ex: "10:00" -> 10)
                hour_num = int(hour_str.split(':')[0])
                row_num = hour_num + 3  # 00:00 = linha 3, 01:00 = linha 4, etc.

                # Atualizar cada variável na planilha
                for variable, value in hour_data.items():
                    col_letter = self._get_column_for_variable_and_day(variable, dia_numero)
                    if col_letter:
                        try:
                            ws[f'{col_letter}{row_num}'] = value
                        except:
                            pass

            dias_atualizados += 1

        return dias_atualizados

    def _get_column_for_variable_and_day(self, variable, dia_numero):
        """Calcula letra da coluna para análise diária"""
        if variable not in self.column_mapping:
            return None

        start_col_num = self.column_mapping[variable]['start_num']
        target_col_num = start_col_num + (dia_numero - 1)
        return get_column_letter(target_col_num)

    def get_updated_excel_file(self):
        """Retorna o arquivo Excel atualizado"""
        if self.excel_path and os.path.exists(self.excel_path):
            with open(self.excel_path, 'rb') as f:
                return f.read()
        return None

    def show_summary(self):
        """Mostra resumo dos dados processados"""
        if not self.dados_processados:
            return None

        summary_data = []
        total_days = 0
        
        for dataset_key, month_data in self.dados_processados.items():
            ano, mes = dataset_key.split('-')
            dias_no_mes = len(month_data['monthly_data'])
            total_days += dias_no_mes
            summary_data.append({
                'Mês/Ano': f"{mes}/{ano}",
                'Dias Processados': dias_no_mes,
                'Horas com Dados': f"{dias_no_mes * 24}h*"  # Aproximado
            })

        return summary_data, total_days

    def show_data_preview(self):
        """Mostra preview detalhada dos dados processados (apenas dados reais)"""
        if not self.dados_processados:
            return
        
        st.markdown("---")
        st.markdown("### 🔍 Preview dos Dados Processados (Apenas Dados Reais)")
        
        # Tabs para diferentes visualizações
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Estatísticas Gerais", "📈 Gráficos", "📋 Dados Mensais", "⏰ Dados Horários Reais"])
        
        with tab1:
            self._show_general_statistics()
        
        with tab2:
            self._show_charts()
        
        with tab3:
            self._show_monthly_data_preview()
        
        with tab4:
            self._show_hourly_data_preview_real()

    def _show_general_statistics(self):
        """Mostra estatísticas gerais dos dados"""
        st.markdown("#### 📊 Estatísticas por Variável")
        
        # Coletar todas as estatísticas
        all_stats = {}
        variables = ['Temp', 'Pir1', 'Pir2', 'PirALB', 'RH', 'Ane', 'Batt', 'LoggTemp', 'LitBatt']
        var_names = {
            'Temp': 'Temperatura (°C)',
            'Pir1': 'Piranômetro 1 (kW/m²)',
            'Pir2': 'Piranômetro 2 (kW/m²)',
            'PirALB': 'Piranômetro Albedo (kW/m²)',
            'RH': 'Umidade Relativa (%)',
            'Ane': 'Velocidade Vento (m/s)',
            'Batt': 'Bateria (V)',
            'LoggTemp': 'Temp. Logger (°C)',
            'LitBatt': 'Bateria Lítio (V)'
        }
        
        for var in variables:
            all_values = []
            outliers_count = 0
            
            for dataset_key, month_data in self.dados_processados.items():
                for dia_numero, stats in month_data['monthly_data'].items():
                    if var in stats:
                        all_values.extend([stats[var]['min'], stats[var]['max'], stats[var]['avg']])
                        outliers_count += stats[var]['outliers']
            
            if all_values:
                all_stats[var_names[var]] = {
                    'Mínimo Global': round(min(all_values), 2),
                    'Máximo Global': round(max(all_values), 2),
                    'Média Global': round(sum(all_values) / len(all_values), 2),
                    'Total Outliers': outliers_count
                }
