import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os
from datetime import datetime, timedelta
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
    Realiza automaticamente análises mensais E diárias com consolidação inteligente
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

    def _get_custom_date_for_timestamp(self, timestamp):
        """
        Determina a qual 'dia customizado' pertence um timestamp
        Dia customizado: 10:10 de um dia até 10:00 do dia seguinte
        """
        if timestamp.hour >= 10 and timestamp.minute >= 10:
            # Se é 10:10 ou depois, pertence ao dia atual
            return timestamp.date()
        else:
            # Se é antes de 10:10, pertence ao dia anterior
            return (timestamp - timedelta(days=1)).date()

    def process_dat_files(self, dat_files):
        """Processa múltiplos arquivos .dat consolidando por TIMESTAMP com lógica inteligente"""
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_files = len(dat_files)
        self.file_processing_info = []
        
        # Lista para armazenar todos os DataFrames
        all_dataframes = []
        
        # ETAPA 1: Ler todos os arquivos
        for i, uploaded_file in enumerate(dat_files):
            status_text.text(f"🔄 Lendo {i+1}/{total_files}: {uploaded_file.name}")
            try:
                uploaded_file.seek(0)
                data = pd.read_csv(uploaded_file, skiprows=4, parse_dates=[0])
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
                all_dataframes.append(data)
                
                # Calcular dias únicos neste arquivo para informação
                unique_dates = data.index.map(self._get_custom_date_for_timestamp).nunique()
                
                self.file_processing_info.append({
                    'arquivo': uploaded_file.name,
                    'registros': len(data),
                    'periodo_inicio': data.index.min().strftime('%Y-%m-%d %H:%M'),
                    'periodo_fim': data.index.max().strftime('%Y-%m-%d %H:%M'),
                    'dias_span': (data.index.max() - data.index.min()).days + 1,
                    'dias_processados': unique_dates,
                    'status': '✅ Lido'
                })
                
            except Exception as e:
                self.file_processing_info.append({
                    'arquivo': uploaded_file.name,
                    'registros': 0,
                    'periodo_inicio': 'N/A',
                    'periodo_fim': 'N/A',
                    'dias_span': 0,
                    'dias_processados': 0,
                    'status': f'❌ Erro: {str(e)}'
                })
            progress_bar.progress((i + 1) / total_files * 0.4)  # 40% para leitura
        
        if not all_dataframes:
            status_text.text("❌ Nenhum arquivo foi lido com sucesso")
            return False
        
        # ETAPA 2: Consolidação inteligente
        status_text.text("🔄 Consolidando dados...")
        df_consolidado = pd.concat(all_dataframes, sort=True)
        
        # Remover duplicatas mantendo a média para timestamps iguais
        duplicate_mask = df_consolidado.index.duplicated(keep=False)
        if duplicate_mask.any():
            status_text.text("🔄 Tratando duplicatas...")
            # Para timestamps duplicados, calcular a média
            df_consolidado = df_consolidado.groupby(df_consolidado.index).mean()
        
        # Ordenar por timestamp
        df_consolidado.sort_index(inplace=True)
        progress_bar.progress(0.6)  # 60% após consolidação
        
        # ETAPA 3: Reamostragem inteligente para preenchimento
        status_text.text("🔄 Aplicando reamostragem inteligente...")
        df_final = self._smart_resample(df_consolidado)
        progress_bar.progress(0.8)  # 80% após reamostragem
        
        # ETAPA 4: Processamento mensal e diário
        status_text.text("🔄 Processando dados mensais e diários...")
        dias_processados_total = self._process_monthly_and_daily_data(df_final)
        
        # Atualizar info de dias processados total (para compatibilidade)
        total_unique_days = df_final.index.map(self._get_custom_date_for_timestamp).nunique()
        for info in self.file_processing_info:
            if "✅" in info['status']:
                # Manter o valor individual já calculado
                pass
        
        progress_bar.progress(1.0)
        status_text.text("✅ Processamento concluído!")
        
        self._show_file_processing_summary()
        return not df_final.empty

    def _smart_resample(self, df):
        """
        Reamostragem inteligente: cria dados horários apenas para preencher lacunas,
        sem sobrescrever dados existentes
        """
        # Identificar todas as horas que já têm dados
        existing_hours = set()
        for timestamp in df.index:
            hour_timestamp = timestamp.replace(minute=0, second=0, microsecond=0)
            existing_hours.add(hour_timestamp)
        
        # Criar índice horário completo para o período
        start_hour = df.index.min().replace(minute=0, second=0, microsecond=0)
        end_hour = df.index.max().replace(minute=0, second=0, microsecond=0)
        
        # Reamostrar para média horária
        df_hourly = df.resample('H').mean()
        
        # Manter apenas dados que realmente existem (remover NaN completos)
        df_hourly = df_hourly.dropna(how='all')
        
        # Combinar dados originais de 10 em 10 minutos com dados horários apenas onde necessário
        combined_data = []
        
        for timestamp in df.index:
            # Manter todos os dados originais
            combined_data.append(df.loc[timestamp])
        
        # Adicionar dados horários apenas para horas que não têm dados próximos
        for hour_timestamp in df_hourly.index:
            # Verificar se já existe algum dado próximo dessa hora (±30min)
            window_start = hour_timestamp - timedelta(minutes=30)
            window_end = hour_timestamp + timedelta(minutes=30)
            
            nearby_data = df[(df.index >= window_start) & (df.index <= window_end)]
            
            # Se não há dados próximos, adicionar o dado horário
            if nearby_data.empty and not pd.isna(df_hourly.loc[hour_timestamp]).all():
                new_row = df_hourly.loc[hour_timestamp].copy()
                combined_data.append(new_row)
        
        # Se não há dados para adicionar, retornar dados originais
        if not combined_data:
            return df
        
        # Reconstituir DataFrame
        result_df = pd.concat([pd.DataFrame([row]) for row in combined_data])
        result_df = result_df.sort_index()
        result_df = result_df[~result_df.index.duplicated(keep='first')]
        
        return result_df

    def _process_monthly_and_daily_data(self, data):
        """
        Processa DataFrame consolidado usando a definição customizada de dia (10:10→10:00)
        """
        if data.empty:
            return 0
        
        # Adicionar coluna de data customizada
        data['custom_date'] = data.index.map(self._get_custom_date_for_timestamp)
        
        dias_processados = 0
        
        # Agrupar por data customizada
        for custom_date, group_data in data.groupby('custom_date'):
            # Determinar mês e ano baseado na data customizada
            mes_numero = custom_date.month
            ano = custom_date.year
            dia_numero = custom_date.day
            
            dataset_key = f"{ano}-{mes_numero:02d}"
            
            if dataset_key not in self.dados_processados:
                self.dados_processados[dataset_key] = {
                    'monthly_data': {},
                    'daily_data': {}
                }
            
            # Remover coluna auxiliar para cálculos
            group_data_clean = group_data.drop('custom_date', axis=1)
            
            # ANÁLISE MENSAL: estatísticas diárias
            stats = self._calculate_daily_statistics(group_data_clean)
            self.dados_processados[dataset_key]['monthly_data'][dia_numero] = stats
            
            # ANÁLISE DIÁRIA: dados horários inteligentes
            if dia_numero not in self.dados_processados[dataset_key]['daily_data']:
                self.dados_processados[dataset_key]['daily_data'][dia_numero] = {}
            
            # Processar dados horários existentes (sem forçar 24 horas)
            for timestamp, row in group_data_clean.iterrows():
                hora_str = timestamp.strftime('%H:00')
                
                # Só adicionar se não existe ou se vai melhorar o dado existente
                if hora_str not in self.dados_processados[dataset_key]['daily_data'][dia_numero]:
                    self.dados_processados[dataset_key]['daily_data'][dia_numero][hora_str] = {
                        'Temperatura': round(row['Temp_Avg'], 2) if not pd.isna(row['Temp_Avg']) else 0,
                        'Piranometro_1': round(row['Pir1_Avg'] / 1000, 3) if not pd.isna(row['Pir1_Avg']) else 0,
                        'Piranometro_2': round(row['Pir2_Avg'] / 1000, 3) if not pd.isna(row['Pir2_Avg']) else 0,
                        'Piranometro_Alab': round(row['PirALB_Avg'] / 1000, 3) if not pd.isna(row['PirALB_Avg']) else 0,
                        'Umidade_Relativa': round(row['RH_Avg'], 2) if not pd.isna(row['RH_Avg']) else 0,
                        'Velocidade_Vento': round(row['Ane_Avg'], 2) if not pd.isna(row['Ane_Avg']) else 0
                    }
            
            dias_processados += 1
        
        return dias_processados

    def _process_hourly_data_for_day(self, day_data):
        """Processa dados horários para um dia específico - MANTIDO PARA COMPATIBILIDADE"""
        day_data['hour'] = day_data.index.hour
        hourly_averages = {}
        
        for hour in range(24):
            hour_data = day_data[day_data['hour'] == hour]
            if len(hour_data) > 0:
                hourly_averages[f"{hour:02d}:00"] = {
                    'Temperatura': round(hour_data['Temp_Avg'].mean(), 2),
                    'Piranometro_1': round(hour_data['Pir1_Avg'].mean() / 1000, 3),
                    'Piranometro_2': round(hour_data['Pir2_Avg'].mean() / 1000, 3),
                    'Piranometro_Alab': round(hour_data['PirALB_Avg'].mean() / 1000, 3),
                    'Umidade_Relativa': round(hour_data['RH_Avg'].mean(), 2),
                    'Velocidade_Vento': round(hour_data['Ane_Avg'].mean(), 2)
                }
        
        return hourly_averages

    def _calculate_daily_statistics(self, data):
        """Calcula estatísticas diárias para análise mensal"""
        stats = {}
        variables = ['Temp', 'Pir1', 'Pir2', 'PirALB', 'RH', 'Ane', 'Batt', 'LoggTemp', 'LitBatt']
        
        for var in variables:
            stats[var] = {
                'min': data[f'{var}_Min'].min() if not data[f'{var}_Min'].isna().all() else 0,
                'max': data[f'{var}_Max'].max() if not data[f'{var}_Max'].isna().all() else 0,
                'avg': data[f'{var}_Avg'].mean() if not data[f'{var}_Avg'].isna().all() else 0,
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
                
                # ANÁLISE DIÁRIA
                aba_diaria = self._find_sheet(wb.sheetnames, mes_numero, "Diaria")
                if aba_diaria:
                    try:
                        ws_diaria = wb[aba_diaria]
                        dias_diario = self._update_daily_data(ws_diaria, month_data['daily_data'])
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
                return True, f"Sucesso! Análise Mensal: {sucesso_mensal} dias, Análise Diária: {sucesso_diario} dias"
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

    def _update_daily_data(self, ws, daily_data):
        """Atualiza dados da análise diária"""
        dias_atualizados = 0
        
        for dia_numero, day_hourly_data in daily_data.items():
            for hour_str, hour_data in day_hourly_data.items():
                hour_num = int(hour_str[:2])
                row_num = hour_num + 3  # 00:00 = linha 3
                
                if row_num < 3:
                    continue
                
                for variable, value in hour_data.items():
                    col_letter = self._get_column_for_variable_and_day(variable, dia_numero)
                    if col_letter and row_num != 2:
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
                'Dias Processados': dias_no_mes
            })
        
        return summary_data, total_days

    def show_data_preview(self):
        """Mostra preview detalhada dos dados processados"""
        if not self.dados_processados:
            return
        
        st.markdown("---")
        st.markdown("### 🔍 Preview dos Dados Processados")
        
        # Tabs para diferentes visualizações
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Estatísticas Gerais", "📈 Gráficos", "📋 Dados Mensais", "⏰ Dados Horários"])
        
        with tab1:
            self._show_general_statistics()
        
        with tab2:
            self._show_charts()
        
        with tab3:
            self._show_monthly_data_preview()
        
        with tab4:
            self._show_hourly_data_preview()

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

    def _show_hourly_data_preview(self):
        """Mostra preview dos dados horários"""
        try:
            st.markdown("#### ⏰ Dados de Análise Diária (Horários)")
            
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
                    
                    # Preparar dados horários
                    hourly_table = []
                    for hour, data in day_data.items():
                        hourly_table.append({
                            'Hora': hour,
                            'Temperatura': data['Temperatura'],
                            'Piranômetro 1': data['Piranometro_1'],
                            'Piranômetro 2': data['Piranometro_2'],
                            'Piranômetro Albedo': data['Piranometro_Alab'],
                            'Umidade Relativa': data['Umidade_Relativa'],
                            'Velocidade Vento': data['Velocidade_Vento']
                        })
                    
                    df_hourly = pd.DataFrame(hourly_table)
                    
                    # Mostrar tabela
                    st.dataframe(df_hourly, use_container_width=True)
                    
                    # Gráfico horário
                    st.markdown("**📊 Variação Horária**")
                    
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
                else:
                    st.info("Selecione um mês e dia para visualizar os dados horários.")
            else:
                st.info("Nenhum dado horário disponível.")
        except Exception as e:
            st.error(f"Erro ao mostrar dados horários: {str(e)}")


def main():
    # Cabeçalho principal
    st.markdown("""
    <div class="main-header">
        <h1>🌤️ Medições Usina Geradora Floriano</h1>
        <p>Processador Completo de Dados Meteorológicos - Consolidação Inteligente</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Inicializar o processador
    if 'processor' not in st.session_state:
        st.session_state.processor = CompleteWeatherProcessor()
    
    # Sidebar com instruções
    with st.sidebar:
        st.markdown("### 📋 Instruções")
        st.markdown("""
        **Passo 1:** Faça upload do arquivo Excel anual
        
        **Passo 2:** Faça upload dos arquivos .dat (múltiplos)
        
        **Passo 3:** Clique em "Processar Dados"
        
        **Passo 4:** Baixe o arquivo Excel atualizado
        """)
        
        st.markdown("---")
        st.markdown("### ℹ️ Consolidação Inteligente")
        st.markdown("""
        **Novos recursos:**
        - ✅ Unifica múltiplos arquivos .dat
        - ✅ Trata duplicatas com média automática
        - ✅ Reamostragem só para preenchimento
        - ✅ Definição customizada de dia (10:10→10:00)
        - ✅ Não força 24 horas nem zeros desnecessários
        """)
        
        st.markdown("---")
        st.markdown("### 🔧 Sobre")
        st.markdown("""
        Este aplicativo processa dados meteorológicos e atualiza automaticamente:
        - **Análises Mensais**: Estatísticas diárias
        - **Análises Diárias**: Dados horários inteligentes
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
            "Selecione os arquivos .dat (múltiplos)",
            type=['dat'],
            accept_multiple_files=True,
            help="Arquivos de dados meteorológicos (.dat) - podem ser de diferentes períodos"
        )
    
    # Informações sobre os arquivos carregados
    if dat_files:
        st.markdown("### 📋 Arquivos .dat Carregados")
        files_info = []
        for file in dat_files:
            files_info.append({
                'Arquivo': file.name,
                'Tamanho': f"{file.size / 1024:.1f} KB"
            })
        df_files = pd.DataFrame(files_info)
        st.dataframe(df_files, use_container_width=True)
    
    # Botão de processamento
    if excel_file and dat_files:
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🚀 Processar Dados com Consolidação Inteligente", use_container_width=True):
                with st.spinner("Processando dados..."):
                    # Processar arquivos .dat
                    success = st.session_state.processor.process_dat_files(dat_files)
                    
                    if success:
                        st.success("✅ Arquivos .dat processados e consolidados com sucesso!")
                        
                        # Mostrar resumo
                        summary_data, total_days = st.session_state.processor.show_summary()
                        if summary_data:
                            st.markdown("### 📊 Resumo dos Dados Processados")
                            col1, col2 = st.columns(2)
                            
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
                        st.markdown("### 🔄 Atualizando Excel...")
                        excel_file.seek(0)  # Reset file pointer
                        success, message = st.session_state.processor.update_excel_file(excel_file)
                        
                        if success:
                            st.success(f"✅ {message}")
                            
                            # Informações sobre abas atualizadas
                            if st.session_state.processor.abas_mensais_atualizadas or st.session_state.processor.abas_diarias_atualizadas:
                                st.markdown("### 📑 Abas Atualizadas")
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    if st.session_state.processor.abas_mensais_atualizadas:
                                        st.markdown("**📊 Análises Mensais:**")
                                        for aba in st.session_state.processor.abas_mensais_atualizadas:
                                            st.markdown(f"- {aba}")
                                
                                with col2:
                                    if st.session_state.processor.abas_diarias_atualizadas:
                                        st.markdown("**⏰ Análises Diárias:**")
                                        for aba in st.session_state.processor.abas_diarias_atualizadas:
                                            st.markdown(f"- {aba}")
                            
                            # Botão de download
                            updated_excel = st.session_state.processor.get_updated_excel_file()
                            if updated_excel:
                                st.markdown("### 📥 Download do Arquivo Atualizado")
                                st.download_button(
                                    label="📥 Baixar Excel Atualizado",
                                    data=updated_excel,
                                    file_name=f"analise_anual_consolidado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
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
        
        if not dat_files:
            st.markdown("""
            **💡 Dica:** Você pode carregar múltiplos arquivos .dat de uma vez! 
            O sistema irá consolidá-los automaticamente, tratando:
            - Dados de diferentes períodos do dia
            - Sobreposições de timestamp (usando média)
            - Preenchimento inteligente de horários
            """)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>🌤️ Processador de Dados Meteorológicos | Usina Geradora Floriano</p>
        <p><small>Versão com Consolidação Inteligente de Múltiplos Arquivos</small></p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
