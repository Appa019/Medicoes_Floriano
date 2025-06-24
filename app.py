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

class RealWeatherProcessor:
    """
    Processador de dados meteorológicos com foco em dados reais
    Processa apenas dados que realmente existem nos arquivos .dat
    Suporta múltiplos dias por arquivo e horários não consecutivos
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
        Processa múltiplos arquivos .dat focando apenas em dados reais
        """
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_files = len(dat_files)
        self.file_processing_info = []
        
        for i, uploaded_file in enumerate(dat_files):
            status_text.text(f"Processando arquivo {i+1}/{total_files}: {uploaded_file.name}")
            
            try:
                # Ler arquivo .dat
                uploaded_file.seek(0)
                data = pd.read_csv(uploaded_file, skiprows=4, parse_dates=[0])

                # Contar registros
                total_records = len(data)
                
                # Renomear colunas
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
                
                # 🔧 NOVA LÓGICA: Processar por data real
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
                
                st.success(f"✅ {uploaded_file.name}: {total_records} registros, {processed_days} dias processados")

            except Exception as e:
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
        
        # Mostrar resumo detalhado
        self._show_file_processing_summary()
        
        return bool(self.dados_processados)

    def _show_file_processing_summary(self):
        """Mostra resumo detalhado do processamento"""
        if hasattr(self, 'file_processing_info') and self.file_processing_info:
            st.markdown("---")
            st.markdown("### 📄 Resumo do Processamento por Arquivo")
            
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
            
            df_display = df_files.copy()
            df_display.columns = [
                'Arquivo', 'Registros', 'Início', 'Fim', 
                'Dias (Span)', 'Dias Processados', 'Status'
            ]
            
            df_display['Registros'] = df_display['Registros'].apply(lambda x: f"{x:,}" if x > 0 else "0")
            
            st.dataframe(df_display, use_container_width=True)

    def _process_monthly_and_daily_data(self, data):
        """
        🔧 NOVA FUNÇÃO: Processa dados por data real de cada registro
        Suporta múltiplos dias por arquivo .dat
        """
        # Adicionar coluna de data real para cada registro
        data['date'] = data.index.date
        days_processed = 0
        
        # 🔧 NOVA LÓGICA: Agrupar dados por data real (um .dat pode ter vários dias)
        for date in data['date'].unique():
            day_data = data[data['date'] == date]
            
            # Extrair ano, mês e dia da data real do registro
            ano = date.year
            mes_numero = date.month
            dia_numero = date.day
            dataset_key = f"{ano}-{mes_numero:02d}"

            # Inicializar estrutura de dados
            if dataset_key not in self.dados_processados:
                self.dados_processados[dataset_key] = {
                    'monthly_data': {},  # Para análise mensal
                    'daily_data': {}     # Para análise diária
                }

            # Estatísticas diárias para análise mensal
            stats = self._calculate_daily_statistics(day_data)
            self.dados_processados[dataset_key]['monthly_data'][dia_numero] = stats

            # 🔧 NOVA FUNCIONALIDADE: Processar apenas dados horários reais
            hourly_data_real = self._process_real_hourly_data(day_data)
            if dia_numero not in self.dados_processados[dataset_key]['daily_data']:
                self.dados_processados[dataset_key]['daily_data'][dia_numero] = {}
            self.dados_processados[dataset_key]['daily_data'][dia_numero] = hourly_data_real
            
            days_processed += 1
        
        return days_processed

    def _process_real_hourly_data(self, day_data):
        """
        🔧 NOVA FUNÇÃO: Processa apenas dados horários que realmente existem
        Elimina o preenchimento artificial de 24h
        
        Regras implementadas:
        - Registros de 10:00-10:50 = hora 10:00
        - Média dos registros disponíveis na hora
        - Apenas horas com dados reais nos arquivos .dat
        """
        day_data['hour'] = day_data.index.hour
        day_data['minute'] = day_data.index.minute
        
        # Dicionário para armazenar apenas dados reais
        hourly_real = {}

        # 🔧 MUDANÇA PRINCIPAL: Encontrar apenas horas que realmente existem
        available_hours = sorted(day_data['hour'].unique())
        
        for hour in available_hours:
            # Filtrar registros da hora atual que seguem padrão de 10 min
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
        🔧 ATUALIZADA: Atualiza Excel com dados reais apenas
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

                # 🔧 ANÁLISE DIÁRIA (NOVA LÓGICA - APENAS DADOS REAIS)
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
        🔧 NOVA FUNÇÃO: Atualiza dados da análise diária usando apenas dados reais
        Não força 24 horas - usa apenas as horas que existem nos dados
        """
        dias_atualizados = 0

        for dia_numero, day_hourly_data in daily_data.items():
            # 🔧 NOVA LÓGICA: Processar apenas as horas que realmente existem
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
        """Mostra resumo dos dados processados (apenas dados reais)"""
        if not self.dados_processados:
            return None

        summary_data = []
        total_days = 0
        total_hours = 0
        
        for dataset_key, month_data in self.dados_processados.items():
            ano, mes = dataset_key.split('-')
            dias_no_mes = len(month_data['monthly_data'])
            total_days += dias_no_mes
            
            # 🔧 CONTAGEM REAL: Contar apenas horas com dados reais
            horas_reais = 0
            for dia_numero, day_data in month_data['daily_data'].items():
                horas_reais += len(day_data)
            
            total_hours += horas_reais
            
            summary_data.append({
                'Mês/Ano': f"{mes}/{ano}",
                'Dias Processados': dias_no_mes,
                'Horas com Dados Reais': horas_reais
            })

        return summary_data, total_days, total_hours

    def show_data_preview(self):
        """🔧 ATUALIZADA: Preview focada em dados reais"""
        if not self.dados_processados:
            return
        
        st.markdown("---")
        st.markdown("### 🔍 Preview dos Dados Processados (Apenas Dados Reais)")
        st.info("🎯 **Dados Reais**: Processando apenas horários que realmente existem nos arquivos .dat")
        
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
        
        # Mostrar em tabela
        if all_stats:
            df_stats = pd.DataFrame(all_stats).T
            st.dataframe(df_stats, use_container_width=True)

    def _show_charts(self):
        """Mostra gráficos dos dados"""
        try:
            st.markdown("#### 📈 Visualizações")
            
            # Preparar dados para gráficos
            chart_data = []
            
            for dataset_key, month_data in self.dados_processados.items():
                ano, mes = dataset_key.split('-')
                for dia_numero, stats in month_data['monthly_data'].items():
                    chart_data.append({
                        'Data': f"{ano}-{mes}-{dia_numero:02d}",
                        'Temperatura Média': round(stats['Temp']['avg'], 2),
                        'Radiação Solar 1': round(stats['Pir1']['avg'] / 1000, 3),
                        'Radiação Solar 2': round(stats['Pir2']['avg'] / 1000, 3),
                        'Umidade Relativa': round(stats['RH']['avg'], 2),
                        'Velocidade Vento': round(stats['Ane']['avg'], 2)
                    })
            
            if chart_data:
                df_chart = pd.DataFrame(chart_data)
                df_chart['Data'] = pd.to_datetime(df_chart['Data'])
                df_chart = df_chart.sort_values('Data')
                
                # Gráfico de temperatura
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**🌡️ Temperatura Média Diária**")
                    st.line_chart(df_chart.set_index('Data')['Temperatura Média'])
                
                with col2:
                    st.markdown("**☀️ Radiação Solar Média**")
                    radiation_data = df_chart.set_index('Data')[['Radiação Solar 1', 'Radiação Solar 2']]
                    st.line_chart(radiation_data)
                
                # Gráfico de umidade e vento
                col3, col4 = st.columns(2)
                
                with col3:
                    st.markdown("**💧 Umidade Relativa**")
                    st.line_chart(df_chart.set_index('Data')['Umidade Relativa'])
                
                with col4:
                    st.markdown("**💨 Velocidade do Vento**")
                    st.line_chart(df_chart.set_index('Data')['Velocidade Vento'])
            else:
                st.info("Nenhum dado disponível para gráficos.")
        except Exception as e:
            st.error(f"Erro ao gerar gráficos: {str(e)}")

    def _show_monthly_data_preview(self):
        """Mostra preview dos dados mensais"""
        try:
            st.markdown("#### 📋 Dados de Análise Mensal")
            
            # Seletor de mês
            available_months = list(self.dados_processados.keys())
            if available_months:
                selected_month = st.selectbox("Selecione o mês para visualizar:", available_months)
                
                if selected_month in self.dados_processados:
                    month_data = self.dados_processados[selected_month]['monthly_data']
                    
                    # Preparar dados para tabela
                    table_data = []
                    for dia, stats in month_data.items():
                        table_data.append({
                            'Dia': dia,
                            'Temp Min': round(stats['Temp']['min'], 2),
                            'Temp Max': round(stats['Temp']['max'], 2),
                            'Temp Média': round(stats['Temp']['avg'], 2),
                            'Rad Solar 1 (kW)': round(stats['Pir1']['avg'] / 1000, 3),
                            'Rad Solar 2 (kW)': round(stats['Pir2']['avg'] / 1000, 3),
                            'Umidade (%)': round(stats['RH']['avg'], 2),
                            'Vento (m/s)': round(stats['Ane']['avg'], 2),
                            'Outliers Total': stats['Temp']['outliers'] + stats['Pir1']['outliers'] + stats['RH']['outliers']
                        })
                    
                    df_monthly = pd.DataFrame(table_data)
                    df_monthly = df_monthly.sort_values('Dia')
                    st.dataframe(df_monthly, use_container_width=True)
            else:
                st.info("Nenhum dado mensal disponível.")
        except Exception as e:
            st.error(f"Erro ao mostrar dados mensais: {str(e)}")

    def _show_hourly_data_preview_real(self):
        """
        🔧 NOVA FUNÇÃO: Mostra preview dos dados horários reais
        Focada apenas em horas que realmente existem nos dados
        """
        try:
            st.markdown("#### ⏰ Dados de Análise Diária (Apenas Dados Reais)")
            
            # Seletores
            available_months = list(self.dados_processados.keys())
            if available_months:
                col1, col2 = st.columns(2)
                
                with col1:
                    selected_month = st.selectbox("Mês:", available_months, key="hourly_month")
                
                with col2:
                    if selected_month in self.dados_processados:
                        available_days = list(self.dados_processados[selected_month]['daily_data'].keys())
                        selected_day = st.selectbox("Dia:", sorted(available_days), key="hourly_day")
                
                if selected_month in self.dados_processados and selected_day in self.dados_processados[selected_month]['daily_data']:
                    day_data = self.dados_processados[selected_month]['daily_data'][selected_day]
                    
                    # 🔧 NOVA LÓGICA: Estatísticas de dados reais
                    total_horas_reais = len(day_data)
                    horas_disponiveis = sorted(day_data.keys())
                    
                    # Detectar lacunas (horas faltantes)
                    if horas_disponiveis:
                        primeiro_hour = int(horas_disponiveis[0].split(':')[0])
                        ultimo_hour = int(horas_disponiveis[-1].split(':')[0])
                        horas_esperadas = ultimo_hour - primeiro_hour + 1
                        lacunas = horas_esperadas - total_horas_reais
                    else:
                        lacunas = 0
                    
                    # Mostrar estatísticas dos dados reais
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h4>⏰ Horas com Dados Reais</h4>
                            <h2>{total_horas_reais}</h2>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        if horas_disponiveis:
                            periodo = f"{horas_disponiveis[0]} - {horas_disponiveis[-1]}"
                        else:
                            periodo = "N/A"
                        st.markdown(f"""
                        <div class="metric-card">
                            <h4>📅 Período de Dados</h4>
                            <h2>{periodo}</h2>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h4>⚠️ Lacunas Detectadas</h4>
                            <h2>{lacunas}</h2>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # 🔧 TABELA DE DADOS REAIS: Apenas horas que existem
                    hourly_table = []
                    for hour_str in sorted(day_data.keys()):
                        data = day_data[hour_str]
                        
                        hourly_table.append({
                            'Hora': hour_str,
                            'Temperatura': data['Temperatura'],
                            'Piranômetro 1': data['Piranometro_1'],
                            'Piranômetro 2': data['Piranometro_2'],
                            'Piranômetro Albedo': data['Piranometro_Alab'],
                            'Umidade Relativa': data['Umidade_Relativa'],
                            'Velocidade Vento': data['Velocidade_Vento'],
                            'Fonte': '📊 Dados Reais'
                        })
                    
                    df_hourly = pd.DataFrame(hourly_table)
                    
                    # Mostrar tabela
                    st.markdown("**📋 Tabela de Dados Reais (Apenas Horas Existentes)**")
                    st.dataframe(df_hourly, use_container_width=True)
                    
                    # Gráfico horário
                    if len(df_hourly) > 1:
                        st.markdown("**📊 Variação Horária (Dados Reais)**")
                        
                        # Preparar dados para gráfico
                        df_hourly['Hora_num'] = df_hourly['Hora'].str[:2].astype(int)
                        df_hourly = df_hourly.sort_values('Hora_num')
                        
                        chart_cols = st.columns(2)
                        
                        with chart_cols[0]:
                            st.markdown("*Temperatura e Umidade*")
                            temp_humidity = df_hourly.set_index('Hora')[['Temperatura', 'Umidade Relativa']]
                            st.line_chart(temp_humidity)
                        
                        with chart_cols[1]:
                            st.markdown("*Radiação Solar*")
                            radiation = df_hourly.set_index('Hora')[['Piranômetro 1', 'Piranômetro 2', 'Piranômetro Albedo']]
                            st.line_chart(radiation)
                    
                    # Informação adicional sobre dados reais
                    if lacunas > 0:
                        st.warning(f"⚠️ {lacunas} horas faltantes detectadas no período de dados")
                    
                    st.success(f"✅ {total_horas_reais} horas com dados reais processadas dos arquivos .dat")
                else:
                    st.info("Selecione um mês e dia para visualizar os dados horários reais.")
            else:
                st.info("Nenhum dado horário disponível.")
        except Exception as e:
            st.error(f"Erro ao mostrar dados horários: {str(e)}")


def main():
    # Cabeçalho principal
    st.markdown("""
    <div class="main-header">
        <h1>🌤️ Medições Usina Geradora Floriano</h1>
        <p>Processador de Dados Reais - Análises Mensais e Diárias Baseadas em Dados Reais</p>
    </div>
    """, unsafe_allow_html=True)

    # Inicializar o processador
    if 'processor' not in st.session_state:
        st.session_state.processor = RealWeatherProcessor()

    # Sidebar com instruções
    with st.sidebar:
        st.markdown("### 📋 Instruções")
        st.markdown("""
        **Passo 1:** Faça upload do arquivo Excel anual
        
        **Passo 2:** Faça upload dos arquivos .dat
        
        **Passo 3:** Clique em "Processar Dados Reais"
        
        **Passo 4:** Baixe o arquivo Excel atualizado
        """)
        
        st.markdown("---")
        st.markdown("### ℹ️ Sobre os Dados Reais")
        st.markdown("""
        **🎯 FOCO EM DADOS REAIS:**
        - ✅ **Sem preenchimento artificial**: Apenas dados que existem nos .dat
        - ✅ **Múltiplos dias por arquivo**: Suporta .dat com vários dias
        - ✅ **Horários não consecutivos**: Processa gaps naturalmente
        - ✅ **Baseado em timestamps reais**: Data/hora dos registros
        """)
        
        st.markdown("---")
        st.markdown("### 📊 Exemplo dos Seus Dados")
        st.markdown("""
        **Arquivos .dat (352, 353, 354, 355):**
        - **Dia 20/06**: 10:10-23:50 ✅
        - **Dia 21/06**: 00:00-10:00 + 10:10-23:50 ✅
        - **Dia 22/06**: 00:00-10:00 + 10:10-23:50 ✅
        - **Dia 23/06**: 00:00-10:00 + 10:10-23:50 ✅
        - **Dia 24/06**: 00:00-10:00 ✅
        """)

    # Layout principal
    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("### 📊 Upload do Excel Anual")
        excel_file = st.file_uploader(
            "Selecione o arquivo Excel anual",
            type=['xlsx', 'xls'],
            help="Arquivo Excel com as abas de análise mensal e diária"
        )

    with col2:
        st.markdown("### 📁 Upload dos Arquivos .dat")
        dat_files = st.file_uploader(
            "Selecione os arquivos .dat",
            type=['dat'],
            accept_multiple_files=True,
            help="Arquivos de dados meteorológicos (.dat) - Apenas dados reais serão processados"
        )

    # Botão de processamento
    if excel_file and dat_files:
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🎯 Processar Dados Reais", use_container_width=True):
                with st.spinner("Processando apenas dados reais dos arquivos .dat..."):
                    # Processar arquivos .dat
                    success = st.session_state.processor.process_dat_files(dat_files)
                    
                    if success:
                        st.success("✅ Arquivos .dat processados - apenas dados reais!")
                        
                        # Mostrar resumo
                        summary_result = st.session_state.processor.show_summary()
                        if summary_result and len(summary_result) == 3:
                            summary_data, total_days, total_hours = summary_result
                            
                            st.markdown("### 📊 Resumo dos Dados Reais Processados")
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.markdown(f"""
                                <div class="metric-card">
                                    <h4>📅 Total de Meses</h4>
                                    <h2>{len(summary_data)}</h2>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            with col2:
                                st.markdown(f"""
                                <div class="metric-card">
                                    <h4>📊 Total de Dias</h4>
                                    <h2>{total_days}</h2>
                                </div>
                                """, unsafe_allow_html=True)
                                
                            with col3:
                                st.markdown(f"""
                                <div class="metric-card">
                                    <h4>⏰ Horas com Dados Reais</h4>
                                    <h2>{total_hours}h</h2>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            # Tabela de resumo
                            df_summary = pd.DataFrame(summary_data)
                            st.dataframe(df_summary, use_container_width=True)
                            
                            # Preview detalhada dos dados
                            try:
                                st.session_state.processor.show_data_preview()
                            except Exception as e:
                                st.error(f"Erro ao mostrar preview dos dados: {str(e)}")
                                st.info("Os dados foram processados com sucesso, mas houve um problema na visualização da preview.")
                        
                        # Atualizar Excel
                        st.markdown("### 🔄 Atualizando Excel com Dados Reais...")
                        excel_file.seek(0)  # Reset file pointer
                        success, message = st.session_state.processor.update_excel_file(excel_file)
                        
                        if success:
                            st.success(f"✅ {message}")
                            
                            # Botão de download
                            updated_excel = st.session_state.processor.get_updated_excel_file()
                            if updated_excel:
                                st.markdown("### 📥 Download do Arquivo Atualizado")
                                st.download_button(
                                    label="📥 Baixar Excel Atualizado (Dados Reais)",
                                    data=updated_excel,
                                    file_name=f"analise_anual_dados_reais_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                        else:
                            st.error(f"❌ {message}")
                    else:
                        st.error("❌ Erro ao processar arquivos .dat")

    # Informações adicionais
    if not excel_file or not dat_files:
        st.markdown("---")
        st.markdown("### 🔍 Aguardando Arquivos")
        
        missing = []
        if not excel_file:
            missing.append("📊 Arquivo Excel anual")
        if not dat_files:
            missing.append("📁 Arquivos .dat")
        
        st.info(f"Por favor, faça upload dos seguintes arquivos: {', '.join(missing)}")

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>🌤️ Processador de Dados Meteorológicos | Usina Geradora Floriano</p>
        <p><strong>🎯 FOCO:</strong> Sistema otimizado para processar apenas dados reais dos arquivos .dat!</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
