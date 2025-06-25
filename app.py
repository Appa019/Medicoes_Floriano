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
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
warnings.filterwarnings('ignore')

# Configuração da página
st.set_page_config(
    page_title="Medições Usina Geradora Floriano",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado com as cores da CSN
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #00529C 0%, #231F20 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        position: relative;
    }
    
    .logo-container {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 2rem;
        margin-bottom: 1rem;
    }
    
    .logo-img {
        height: 80px;
        width: auto;
        background: white;
        padding: 10px;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.3);
    }
    
    .header-text {
        flex: 1;
        text-align: center;
    }
    
    .main-header h1 {
        color: white;
        font-size: 2.8rem;
        margin: 0;
        font-weight: bold;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .main-header p {
        color: #E8E8E8;
        font-size: 1.3rem;
        margin: 0.5rem 0 0 0;
    }
    
    .stButton > button {
        background: linear-gradient(45deg, #00529C, #0066CC);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.7rem 1.5rem;
        font-weight: bold;
        font-size: 1.1rem;
        transition: all 0.3s ease;
        box-shadow: 0 3px 10px rgba(0,82,156,0.3);
    }
    
    .stButton > button:hover {
        background: linear-gradient(45deg, #231F20, #404040);
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(0,82,156,0.4);
    }
    
    .success-box {
        background: linear-gradient(135deg, #d4edda, #c3e6cb);
        border: 1px solid #c3e6cb;
        border-radius: 10px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .error-box {
        background: linear-gradient(135deg, #f8d7da, #f5c6cb);
        border: 1px solid #f5c6cb;
        border-radius: 10px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .info-box {
        background: linear-gradient(135deg, #d1ecf1, #bee5eb);
        border: 1px solid #bee5eb;
        border-radius: 10px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .warning-box {
        background: linear-gradient(135deg, #fff3cd, #ffeaa7);
        border: 1px solid #ffeaa7;
        border-radius: 10px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .metric-card {
        background: linear-gradient(135deg, #ffffff, #f8f9fa);
        padding: 1.5rem;
        border-radius: 12px;
        border-left: 5px solid #00529C;
        margin: 0.5rem 0;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        transition: transform 0.2s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0,0,0,0.15);
    }
    
    .metric-card h4 {
        color: #00529C;
        margin: 0 0 0.5rem 0;
        font-size: 1rem;
        font-weight: 600;
    }
    
    .metric-card h2 {
        color: #231F20;
        margin: 0;
        font-size: 2rem;
        font-weight: bold;
    }
    
    .stats-container {
        background: linear-gradient(135deg, #f8f9fa, #e9ecef);
        padding: 2rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    .chart-container {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        margin: 1rem 0;
        box-shadow: 0 3px 10px rgba(0,0,0,0.1);
        border: 1px solid #e9ecef;
    }
    
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #f8f9fa, #ffffff);
    }
    
    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .stExpander {
        border-radius: 10px;
        border: 1px solid #e9ecef;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    
    /* Responsividade para dispositivos móveis */
    @media (max-width: 768px) {
        .logo-container {
            flex-direction: column;
            gap: 1rem;
        }
        
        .main-header h1 {
            font-size: 2.2rem;
        }
        
        .logo-img {
            height: 60px;
        }
    }
</style>
""", unsafe_allow_html=True)

class ExactWeatherProcessor:
    """
    Processador de dados meteorológicos com busca EXATA tipo PROCV
    NÃO faz médias ou inferências - apenas busca dados pontuais com tolerância de ±10 minutos
    """
    def __init__(self):
        self.consolidated_data = {}  # {timestamp: {variavel: valor}}
        self.processed_sheets = []
        self.conflicts_detected = []
        self.excel_path = None
        
        # Mapeamento de colunas para análise diária
        self.column_mapping = {
            'Temperatura': {'start_num': 2},        # B até AF (2-32)
            'Piranometro_1': {'start_num': 33},     # AG até BK (33-63)
            'Piranometro_2': {'start_num': 64},     # BL até CP (64-94)
            'Piranometro_Alab': {'start_num': 95},  # CQ até DU (95-125)
            'Umidade_Relativa': {'start_num': 126}, # DV até EZ (126-156)
            'Velocidade_Vento': {'start_num': 157}  # FA até GE (157-187)
        }

    def process_dat_files(self, dat_files):
        """Processa múltiplos arquivos .dat consolidando por TIMESTAMP exato"""
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_files = len(dat_files)
        self.file_processing_info = []
        self.conflicts_detected = []
        
        # ETAPA 1: Ler todos os arquivos e consolidar
        for i, uploaded_file in enumerate(dat_files):
            status_text.text(f"Processando {i+1}/{total_files}: {uploaded_file.name}")
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
                
                # Consolidar dados timestamp por timestamp
                for _, row in data.iterrows():
                    timestamp = row['TIMESTAMP']
                    
                    # Extrair dados das variáveis
                    new_data = {
                        'Temperatura': round(row['Temp_Avg'], 2) if not pd.isna(row['Temp_Avg']) else None,
                        'Piranometro_1': round(row['Pir1_Avg'] / 1000, 3) if not pd.isna(row['Pir1_Avg']) else None,
                        'Piranometro_2': round(row['Pir2_Avg'] / 1000, 3) if not pd.isna(row['Pir2_Avg']) else None,
                        'Piranometro_Alab': round(row['PirALB_Avg'] / 1000, 3) if not pd.isna(row['PirALB_Avg']) else None,
                        'Umidade_Relativa': round(row['RH_Avg'], 2) if not pd.isna(row['RH_Avg']) else None,
                        'Velocidade_Vento': round(row['Ane_Avg'], 2) if not pd.isna(row['Ane_Avg']) else None
                    }
                    
                    # Verificar se já existe dados para este timestamp
                    if timestamp in self.consolidated_data:
                        # CONFLITO DETECTADO!
                        conflict_info = {
                            'timestamp': timestamp,
                            'arquivo_anterior': 'dados_anteriores',
                            'arquivo_atual': uploaded_file.name,
                            'dados_anteriores': self.consolidated_data[timestamp].copy(),
                            'dados_novos': new_data.copy()
                        }
                        self.conflicts_detected.append(conflict_info)
                        
                        # Usar último arquivo (sobrescrever)
                        self.consolidated_data[timestamp] = new_data
                    else:
                        # Novo timestamp, adicionar
                        self.consolidated_data[timestamp] = new_data
                
                # Info do arquivo processado
                self.file_processing_info.append({
                    'arquivo': uploaded_file.name,
                    'registros': len(data),
                    'periodo_inicio': data['TIMESTAMP'].min().strftime('%Y-%m-%d %H:%M'),
                    'periodo_fim': data['TIMESTAMP'].max().strftime('%Y-%m-%d %H:%M'),
                    'status': 'Processado com sucesso'
                })
                
            except Exception as e:
                self.file_processing_info.append({
                    'arquivo': uploaded_file.name,
                    'registros': 0,
                    'periodo_inicio': 'N/A',
                    'periodo_fim': 'N/A',
                    'status': f'Erro: {str(e)}'
                })
            
            progress_bar.progress((i + 1) / total_files)
        
        status_text.text("Consolidação concluída com sucesso!")
        
        # Mostrar conflitos se detectados
        if self.conflicts_detected:
            self._show_conflicts()
        
        # Mostrar resumo do processamento
        self._show_file_processing_summary()
        
        return len(self.consolidated_data) > 0

    def _show_conflicts(self):
        """Mostra conflitos detectados entre arquivos"""
        st.markdown("---")
        st.markdown("### Conflitos Detectados")
        
        st.markdown(f"""
        <div class="warning-box">
            <h4>{len(self.conflicts_detected)} conflito(s) encontrado(s)</h4>
            <p>Timestamps idênticos em múltiplos arquivos. Usando dados do último arquivo processado.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Mostrar detalhes dos conflitos
        with st.expander("Ver Detalhes dos Conflitos"):
            for i, conflict in enumerate(self.conflicts_detected[:10]):  # Mostrar só os primeiros 10
                st.markdown(f"**Conflito {i+1}: {conflict['timestamp']}**")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("*Dados Anteriores:*")
                    st.json(conflict['dados_anteriores'])
                with col2:
                    st.markdown(f"*Dados Novos ({conflict['arquivo_atual']}):*")
                    st.json(conflict['dados_novos'])
                st.markdown("---")
            
            if len(self.conflicts_detected) > 10:
                st.info(f"Mostrando apenas os primeiros 10 conflitos de {len(self.conflicts_detected)} total.")

    def _find_closest_timestamp(self, target_time, available_timestamps):
        """
        Busca o timestamp mais próximo dentro da tolerância de ±10 minutos
        
        Args:
            target_time: datetime alvo (ex: 2025-06-22 10:00:00)
            available_timestamps: lista de timestamps disponíveis
            
        Returns:
            timestamp mais próximo ou None se nenhum estiver dentro da tolerância
        """
        target_time = pd.to_datetime(target_time)
        tolerance = timedelta(minutes=10)
        
        min_diff = timedelta.max
        closest_timestamp = None
        
        for ts in available_timestamps:
            ts_converted = pd.to_datetime(ts)
            diff = abs(ts_converted - target_time)
            
            # Verifica se está dentro da tolerância e é mais próximo
            if diff <= tolerance and diff < min_diff:
                min_diff = diff
                closest_timestamp = ts  # Retorna o timestamp original, não o convertido
        
        return closest_timestamp

    def update_excel_file(self, excel_file):
        """
        Atualiza Excel com dados exatos usando lógica PROCV
        """
        if not self.consolidated_data:
            return False, "Nenhum dado processado!"
        
        try:
            # Salvar arquivo Excel temporariamente
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(excel_file.read())
                self.excel_path = tmp_file.name
            
            wb = load_workbook(self.excel_path)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Agrupar dados por mês
            monthly_data = {}
            for timestamp, data in self.consolidated_data.items():
                year_month = f"{timestamp.year}-{timestamp.month:02d}"
                if year_month not in monthly_data:
                    monthly_data[year_month] = {}
                monthly_data[year_month][timestamp] = data
            
            total_months = len(monthly_data)
            sheets_updated = 0
            total_cells_updated = 0
            
            # Processar cada mês
            for i, (year_month, month_timestamps) in enumerate(monthly_data.items()):
                year, month = year_month.split('-')
                month_num = int(month)
                
                status_text.text(f"Processando mês {month}/{year}...")
                
                # Buscar aba correspondente
                sheet_name = self._find_daily_analysis_sheet(wb.sheetnames, month_num)
                if not sheet_name:
                    st.warning(f"Aba para mês {month:02d} não encontrada!")
                    continue
                
                ws = wb[sheet_name]
                cells_updated = self._update_daily_analysis_exact(ws, month_timestamps, int(year), month_num)
                
                if cells_updated > 0:
                    sheets_updated += 1
                    total_cells_updated += cells_updated
                    self.processed_sheets.append(sheet_name)
                
                progress_bar.progress((i + 1) / total_months)
            
            # Salvar alterações
            wb.save(self.excel_path)
            status_text.text("Atualização concluída com sucesso!")
            
            if sheets_updated > 0:
                return True, f"Sucesso! {sheets_updated} aba(s) atualizada(s), {total_cells_updated} célula(s) preenchida(s)"
            else:
                return False, "Nenhuma aba compatível encontrada para atualização"
                
        except Exception as e:
            return False, f"Erro durante atualização: {e}"

    def _find_daily_analysis_sheet(self, sheet_names, month_num):
        """Encontra aba de análise diária para o mês"""
        month_str = f"{month_num:02d}"
        target_pattern = f"{month_str}-Analise Diaria"
        
        # Busca exata primeiro
        if target_pattern in sheet_names:
            return target_pattern
        
        # Busca por padrão similar
        for sheet_name in sheet_names:
            if month_str in sheet_name and "Analise Diaria" in sheet_name:
                return sheet_name
        
        return None

    def _update_daily_analysis_exact(self, ws, month_timestamps, year, month):
        """
        Atualiza análise diária usando busca EXATA tipo PROCV
        """
        cells_updated = 0
        
        # Para cada horário da planilha (00:00 a 23:00)
        for hour in range(24):
            row_num = hour + 3  # Linha 3 = 00:00, Linha 4 = 01:00, etc.
            
            # Para cada dia do mês (1 a 31)
            for day in range(1, 32):
                # Construir timestamp alvo
                try:
                    target_datetime = datetime(year, month, day, hour, 0, 0)
                except ValueError:
                    # Dia inválido para o mês (ex: 31 de fevereiro)
                    continue
                
                # Buscar timestamp mais próximo dentro da tolerância
                available_timestamps = list(month_timestamps.keys())
                closest_timestamp = self._find_closest_timestamp(target_datetime, available_timestamps)
                
                if closest_timestamp is None:
                    # Nenhum dado dentro da tolerância - deixar vazio
                    continue
                
                # Obter dados do timestamp encontrado
                data = month_timestamps[closest_timestamp]
                
                # Atualizar cada variável
                for variable, value in data.items():
                    if value is None:
                        continue
                    
                    col_letter = self._get_column_for_variable_and_day(variable, day)
                    if col_letter:
                        try:
                            ws[f'{col_letter}{row_num}'] = value
                            cells_updated += 1
                        except Exception:
                            # Continua processamento mesmo com erro
                            pass
        
        return cells_updated

    def _get_column_for_variable_and_day(self, variable, day_number):
        """
        Calcula letra da coluna para análise diária
        """
        if variable not in self.column_mapping:
            return None
        
        # Verificar se o dia está no range válido (1-31)
        if day_number < 1 or day_number > 31:
            return None
        
        start_col_num = self.column_mapping[variable]['start_num']
        target_col_num = start_col_num + (day_number - 1)
        
        # Verificar se a coluna está dentro dos limites válidos
        if target_col_num > 187:  # Última coluna GE = 187
            return None
            
        col_letter = get_column_letter(target_col_num)
        return col_letter

    def _show_file_processing_summary(self):
        """Mostra resumo detalhado do processamento"""
        if hasattr(self, 'file_processing_info') and self.file_processing_info:
            st.markdown("---")
            st.markdown("### Resumo do Processamento")
            
            # Criar DataFrame com as informações
            df_files = pd.DataFrame(self.file_processing_info)
            
            # Calcular totais
            total_records = df_files['registros'].sum()
            total_files_success = len([f for f in self.file_processing_info if 'sucesso' in f['status'].lower()])
            total_timestamps = len(self.consolidated_data)
            
            # Mostrar métricas gerais
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>Arquivos Processados</h4>
                    <h2>{total_files_success}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>Registros Lidos</h4>
                    <h2>{total_records:,}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>Timestamps Únicos</h4>
                    <h2>{total_timestamps:,}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>Conflitos Detectados</h4>
                    <h2>{len(self.conflicts_detected)}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            # Tabela detalhada
            st.markdown("#### Detalhes por Arquivo")
            df_display = df_files.copy()
            df_display.columns = ['Arquivo', 'Registros', 'Início', 'Fim', 'Status']
            df_display['Registros'] = df_display['Registros'].apply(lambda x: f"{x:,}" if x > 0 else "0")
            
            st.dataframe(df_display, use_container_width=True)

    def get_updated_excel_file(self):
        """Retorna o arquivo Excel atualizado"""
        if self.excel_path and os.path.exists(self.excel_path):
            with open(self.excel_path, 'rb') as f:
                return f.read()
        return None

    def show_data_preview_and_charts(self):
        """Mostra preview dos dados consolidados com gráficos para conferência"""
        if not self.consolidated_data:
            return
        
        st.markdown("---")
        st.markdown("### Análise dos Dados Consolidados")
        
        # Converter para DataFrame para visualização
        preview_data = []
        for timestamp, data in self.consolidated_data.items():
            row = {'Timestamp': timestamp}
            row.update(data)
            preview_data.append(row)
        
        if preview_data:
            df_preview = pd.DataFrame(preview_data)
            df_preview = df_preview.sort_values('Timestamp')
            
            # Estatísticas gerais
            st.markdown("#### Estatísticas Gerais")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                first_timestamp = min(self.consolidated_data.keys())
                last_timestamp = max(self.consolidated_data.keys())
                period_days = (last_timestamp - first_timestamp).days + 1
                st.metric("Período Total", f"{period_days} dias")
            
            with col2:
                timestamps_per_day = len(self.consolidated_data) / period_days if period_days > 0 else 0
                st.metric("Registros/Dia", f"{timestamps_per_day:.1f}")
            
            with col3:
                # Agrupar por mês
                months = set()
                for ts in self.consolidated_data.keys():
                    months.add(f"{ts.year}-{ts.month:02d}")
                st.metric("Meses Cobertos", len(months))
            
            # Gráficos para conferência das variáveis
            self._create_variable_charts(df_preview)
            
            # Preview da tabela de dados
            st.markdown("#### Preview dos Dados (Primeiros 100 registros)")
            st.dataframe(df_preview.head(100), use_container_width=True)
            
        else:
            st.info("Nenhum dado disponível para preview.")

    def _create_variable_charts(self, df):
        """Cria gráficos para conferência das variáveis meteorológicas"""
        st.markdown("#### Gráficos de Conferência das Variáveis")
        
        # Preparar dados para gráficos
        df_clean = df.dropna()
        
        if len(df_clean) == 0:
            st.warning("Não há dados suficientes para gerar gráficos.")
            return
        
        # Gráfico 1: Temperatura ao longo do tempo
        with st.container():
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.markdown("**Temperatura (°C)**")
            
            fig_temp = px.line(df_clean, x='Timestamp', y='Temperatura',
                              title='Variação da Temperatura ao Longo do Tempo',
                              color_discrete_sequence=['#00529C'])
            fig_temp.update_layout(
                xaxis_title="Data/Hora",
                yaxis_title="Temperatura (°C)",
                height=400,
                showlegend=False
            )
            st.plotly_chart(fig_temp, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Gráfico 2: Piranômetros (Radiação Solar)
        with st.container():
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.markdown("**Radiação Solar (kW/m²)**")
            
            fig_pir = go.Figure()
            
            if 'Piranometro_1' in df_clean.columns:
                fig_pir.add_trace(go.Scatter(x=df_clean['Timestamp'], y=df_clean['Piranometro_1'],
                                           mode='lines', name='Piranômetro 1', line=dict(color='#FF6B35')))
            
            if 'Piranometro_2' in df_clean.columns:
                fig_pir.add_trace(go.Scatter(x=df_clean['Timestamp'], y=df_clean['Piranometro_2'],
                                           mode='lines', name='Piranômetro 2', line=dict(color='#F7931E')))
            
            if 'Piranometro_Alab' in df_clean.columns:
                fig_pir.add_trace(go.Scatter(x=df_clean['Timestamp'], y=df_clean['Piranometro_Alab'],
                                           mode='lines', name='Piranômetro Albedo', line=dict(color='#FFD23F')))
            
            fig_pir.update_layout(
                title='Radiação Solar - Comparação dos Piranômetros',
                xaxis_title="Data/Hora",
                yaxis_title="Radiação Solar (kW/m²)",
                height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_pir, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Gráfico 3: Umidade e Velocidade do Vento
        with st.container():
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.markdown("**Umidade Relativa e Velocidade do Vento**")
            
            # Criar subplots com eixos Y duplos
            fig_combined = make_subplots(specs=[[{"secondary_y": True}]])
            
            if 'Umidade_Relativa' in df_clean.columns:
                fig_combined.add_trace(
                    go.Scatter(x=df_clean['Timestamp'], y=df_clean['Umidade_Relativa'],
                             mode='lines', name='Umidade Relativa (%)', line=dict(color='#4A90E2')),
                    secondary_y=False,
                )
            
            if 'Velocidade_Vento' in df_clean.columns:
                fig_combined.add_trace(
                    go.Scatter(x=df_clean['Timestamp'], y=df_clean['Velocidade_Vento'],
                             mode='lines', name='Velocidade do Vento (m/s)', line=dict(color='#50C878')),
                    secondary_y=True,
                )
            
            # Configurar eixos Y
            fig_combined.update_xaxes(title_text="Data/Hora")
            fig_combined.update_yaxes(title_text="Umidade Relativa (%)", secondary_y=False)
            fig_combined.update_yaxes(title_text="Velocidade do Vento (m/s)", secondary_y=True)
            
            fig_combined.update_layout(
                title='Umidade Relativa e Velocidade do Vento',
                height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_combined, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Estatísticas descritivas
        st.markdown("#### Estatísticas Descritivas")
        
        # Selecionar apenas colunas numéricas
        numeric_cols = ['Temperatura', 'Piranometro_1', 'Piranometro_2', 'Piranometro_Alab', 'Umidade_Relativa', 'Velocidade_Vento']
        available_cols = [col for col in numeric_cols if col in df_clean.columns]
        
        if available_cols:
            stats_df = df_clean[available_cols].describe().round(3)
            stats_df.index = ['Contagem', 'Média', 'Desvio Padrão', 'Mínimo', '25%', '50% (Mediana)', '75%', 'Máximo']
            
            st.markdown('<div class="stats-container">', unsafe_allow_html=True)
            st.dataframe(stats_df, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Distribuição dos dados por hora do dia
        st.markdown("#### Distribuição por Hora do Dia")
        
        if len(df_clean) > 0:
            df_clean['Hora'] = df_clean['Timestamp'].dt.hour
            
            # Gráfico de distribuição por hora
            hourly_stats = df_clean.groupby('Hora')[available_cols].mean().reset_index()
            
            fig_hourly = go.Figure()
            
            colors = ['#00529C', '#FF6B35', '#F7931E', '#FFD23F', '#4A90E2', '#50C878']
            
            for i, col in enumerate(available_cols):
                if col in hourly_stats.columns:
                    fig_hourly.add_trace(go.Scatter(
                        x=hourly_stats['Hora'], 
                        y=hourly_stats[col],
                        mode='lines+markers',
                        name=col.replace('_', ' ').title(),
                        line=dict(color=colors[i % len(colors)])
                    ))
            
            fig_hourly.update_layout(
                title='Valores Médios por Hora do Dia',
                xaxis_title="Hora do Dia",
                yaxis_title="Valores Médios",
                height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            
            st.plotly_chart(fig_hourly, use_container_width=True)


def main():
    # Cabeçalho principal com logo da CSN
    st.markdown("""
    <div class="main-header">
        <div class="logo-container">
            <img src="https://upload.wikimedia.org/wikipedia/pt/e/eb/Companhia_Sider%C3%BArgica_Nacional.png" 
                 alt="Logo CSN" class="logo-img">
            <div class="header-text">
                <h1>Medições Usina Geradora Floriano</h1>
                <p>Processador de Dados Meteorológicos - VERSÃO PROCV EXATO</p>
                <p><small>Busca Pontual | Tolerância ±10min | Zero Inferências</small></p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Inicializar o processador
    if 'processor' not in st.session_state:
        st.session_state.processor = ExactWeatherProcessor()
    
    # Sidebar com instruções
    with st.sidebar:
        st.markdown("### Instruções de Uso")
        st.markdown("""
        **Passo 1:** Upload do arquivo Excel anual
        
        **Passo 2:** Upload dos arquivos .dat (múltiplos)
        
        **Passo 3:** Clique em "Processar Dados"
        
        **Passo 4:** Baixe o Excel atualizado
        """)
        
        st.markdown("---")
        st.markdown("### Funcionalidades")
        st.markdown("""
        **Características:**
        - Busca pontual de dados (sem médias)
        - Tolerância de ±10 minutos
        - Zero inferências ou preenchimentos
        - Detecção de conflitos entre arquivos
        - Mapeamento preciso por timestamp
        - Gráficos de conferência das variáveis
        
        **Lógica de Busca:**
        - Para 10:00 → busca entre 09:50 e 10:10
        - Prioriza timestamp mais próximo
        - Deixa vazio se não há dados na tolerância
        """)
        
        st.markdown("---")
        st.markdown("### Mapeamento de Colunas")
        st.markdown("""
        - **Temperatura**: Colunas B-AF (Dias 1-31)
        - **Piranômetro 1**: Colunas AG-BK (Dias 1-31)
        - **Piranômetro 2**: Colunas BL-CP (Dias 1-31)
        - **Piranômetro Albedo**: Colunas CQ-DU (Dias 1-31)
        - **Umidade**: Colunas DV-EZ (Dias 1-31)
        - **Vento**: Colunas FA-GE (Dias 1-31)
        """)
    
    # Layout principal
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown("### Upload do Excel Anual")
        excel_file = st.file_uploader(
            "Selecione o arquivo Excel anual",
            type=['xlsx', 'xls'],
            help="Arquivo Excel com as abas XX-Analise Diaria"
        )
    
    with col2:
        st.markdown("### Upload dos Arquivos .dat")
        dat_files = st.file_uploader(
            "Selecione os arquivos .dat (múltiplos)",
            type=['dat'],
            accept_multiple_files=True,
            help="Arquivos de dados meteorológicos (.dat) com timestamps de 10 em 10 minutos"
        )
    
    # Informações sobre os arquivos carregados
    if dat_files:
        st.markdown("### Arquivos .dat Carregados")
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
            if st.button("Processar Dados PROCV EXATO", use_container_width=True):
                with st.spinner("Processando dados com busca pontual..."):
                    # Processar arquivos .dat
                    success = st.session_state.processor.process_dat_files(dat_files)
                    
                    if success:
                        st.success("Arquivos .dat processados e consolidados com sucesso!")
                        
                        # Mostrar preview dos dados com gráficos
                        st.session_state.processor.show_data_preview_and_charts()
                        
                        # Atualizar Excel
                        st.markdown("### Atualizando Excel com Busca PROCV...")
                        excel_file.seek(0)  # Reset file pointer
                        success, message = st.session_state.processor.update_excel_file(excel_file)
                        
                        if success:
                            st.success(f"{message}")
                            
                            # Informações sobre abas atualizadas
                            if st.session_state.processor.processed_sheets:
                                st.markdown("### Abas Atualizadas")
                                for sheet in st.session_state.processor.processed_sheets:
                                    st.markdown(f"- {sheet}")
                            
                            # Botão de download
                            updated_excel = st.session_state.processor.get_updated_excel_file()
                            if updated_excel:
                                st.markdown("### Download do Arquivo Atualizado")
                                st.download_button(
                                    label="Baixar Excel Atualizado (PROCV)",
                                    data=updated_excel,
                                    file_name=f"analise_procv_exato_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                        else:
                            st.error(f"{message}")
                    else:
                        st.error("Erro ao processar arquivos .dat")
    
    # Informações adicionais
    if not excel_file or not dat_files:
        st.markdown("---")
        st.markdown("### Aguardando Arquivos")
        missing = []
        if not excel_file:
            missing.append("Arquivo Excel anual")
        if not dat_files:
            missing.append("Arquivos .dat")
        
        st.info(f"Por favor, faça upload dos seguintes arquivos: {', '.join(missing)}")
        
        if not dat_files:
            st.markdown("""
            **Sobre a Busca PROCV Exata:**
            - Busca dados pontuais sem fazer médias
            - Tolerância de ±10 minutos para cada horário
            - Não preenche dados que não existem
            - Detecta e alerta sobre conflitos entre arquivos
            - Mapeia diretamente timestamp → célula da planilha
            - Gera gráficos para conferência visual dos dados
            """)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>Processador de Dados Meteorológicos | Usina Geradora Floriano</p>
        <p><small>Versão PROCV EXATO - Busca Pontual com Tolerância ±10min</small></p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()

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
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
warnings.filterwarnings('ignore')

# Configuração da página
st.set_page_config(
    page_title="Medições Usina Geradora Floriano",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado com as cores da CSN
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #00529C 0%, #231F20 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    
    .main-header h1 {
        color: white;
        font-size: 2.8rem;
        margin: 0;
        font-weight: bold;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .main-header p {
        color: #E8E8E8;
        font-size: 1.3rem;
        margin: 0.5rem 0 0 0;
    }
    
    .stButton > button {
        background: linear-gradient(45deg, #00529C, #0066CC);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.7rem 1.5rem;
        font-weight: bold;
        font-size: 1.1rem;
        transition: all 0.3s ease;
        box-shadow: 0 3px 10px rgba(0,82,156,0.3);
    }
    
    .stButton > button:hover {
        background: linear-gradient(45deg, #231F20, #404040);
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(0,82,156,0.4);
    }
    
    .success-box {
        background: linear-gradient(135deg, #d4edda, #c3e6cb);
        border: 1px solid #c3e6cb;
        border-radius: 10px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .error-box {
        background: linear-gradient(135deg, #f8d7da, #f5c6cb);
        border: 1px solid #f5c6cb;
        border-radius: 10px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .info-box {
        background: linear-gradient(135deg, #d1ecf1, #bee5eb);
        border: 1px solid #bee5eb;
        border-radius: 10px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .warning-box {
        background: linear-gradient(135deg, #fff3cd, #ffeaa7);
        border: 1px solid #ffeaa7;
        border-radius: 10px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .metric-card {
        background: linear-gradient(135deg, #ffffff, #f8f9fa);
        padding: 1.5rem;
        border-radius: 12px;
        border-left: 5px solid #00529C;
        margin: 0.5rem 0;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        transition: transform 0.2s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0,0,0,0.15);
    }
    
    .metric-card h4 {
        color: #00529C;
        margin: 0 0 0.5rem 0;
        font-size: 1rem;
        font-weight: 600;
    }
    
    .metric-card h2 {
        color: #231F20;
        margin: 0;
        font-size: 2rem;
        font-weight: bold;
    }
    
    .stats-container {
        background: linear-gradient(135deg, #f8f9fa, #e9ecef);
        padding: 2rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    .chart-container {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        margin: 1rem 0;
        box-shadow: 0 3px 10px rgba(0,0,0,0.1);
        border: 1px solid #e9ecef;
    }
    
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #f8f9fa, #ffffff);
    }
    
    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .stExpander {
        border-radius: 10px;
        border: 1px solid #e9ecef;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
</style>
""", unsafe_allow_html=True)

class ExactWeatherProcessor:
    """
    Processador de dados de Medicao
    NÃO faz médias ou inferências - apenas busca dados pontuais com tolerância de ±10 minutos
    """
    def __init__(self):
        self.consolidated_data = {}  # {timestamp: {variavel: valor}}
        self.processed_sheets = []
        self.conflicts_detected = []
        self.excel_path = None
        
        # Mapeamento de colunas para análise diária
        self.column_mapping = {
            'Temperatura': {'start_num': 2},        # B até AF (2-32)
            'Piranometro_1': {'start_num': 33},     # AG até BK (33-63)
            'Piranometro_2': {'start_num': 64},     # BL até CP (64-94)
            'Piranometro_Alab': {'start_num': 95},  # CQ até DU (95-125)
            'Umidade_Relativa': {'start_num': 126}, # DV até EZ (126-156)
            'Velocidade_Vento': {'start_num': 157}  # FA até GE (157-187)
        }

    def process_dat_files(self, dat_files):
        """Processa múltiplos arquivos .dat consolidando por TIMESTAMP exato"""
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_files = len(dat_files)
        self.file_processing_info = []
        self.conflicts_detected = []
        
        # ETAPA 1: Ler todos os arquivos e consolidar
        for i, uploaded_file in enumerate(dat_files):
            status_text.text(f"Processando {i+1}/{total_files}: {uploaded_file.name}")
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
                
                # Consolidar dados timestamp por timestamp
                for _, row in data.iterrows():
                    timestamp = row['TIMESTAMP']
                    
                    # Extrair dados das variáveis
                    new_data = {
                        'Temperatura': round(row['Temp_Avg'], 2) if not pd.isna(row['Temp_Avg']) else None,
                        'Piranometro_1': round(row['Pir1_Avg'] / 1000, 3) if not pd.isna(row['Pir1_Avg']) else None,
                        'Piranometro_2': round(row['Pir2_Avg'] / 1000, 3) if not pd.isna(row['Pir2_Avg']) else None,
                        'Piranometro_Alab': round(row['PirALB_Avg'] / 1000, 3) if not pd.isna(row['PirALB_Avg']) else None,
                        'Umidade_Relativa': round(row['RH_Avg'], 2) if not pd.isna(row['RH_Avg']) else None,
                        'Velocidade_Vento': round(row['Ane_Avg'], 2) if not pd.isna(row['Ane_Avg']) else None
                    }
                    
                    # Verificar se já existe dados para este timestamp
                    if timestamp in self.consolidated_data:
                        # CONFLITO DETECTADO!
                        conflict_info = {
                            'timestamp': timestamp,
                            'arquivo_anterior': 'dados_anteriores',
                            'arquivo_atual': uploaded_file.name,
                            'dados_anteriores': self.consolidated_data[timestamp].copy(),
                            'dados_novos': new_data.copy()
                        }
                        self.conflicts_detected.append(conflict_info)
                        
                        # Usar último arquivo (sobrescrever)
                        self.consolidated_data[timestamp] = new_data
                    else:
                        # Novo timestamp, adicionar
                        self.consolidated_data[timestamp] = new_data
                
                # Info do arquivo processado
                self.file_processing_info.append({
                    'arquivo': uploaded_file.name,
                    'registros': len(data),
                    'periodo_inicio': data['TIMESTAMP'].min().strftime('%Y-%m-%d %H:%M'),
                    'periodo_fim': data['TIMESTAMP'].max().strftime('%Y-%m-%d %H:%M'),
                    'status': 'Processado com sucesso'
                })
                
            except Exception as e:
                self.file_processing_info.append({
                    'arquivo': uploaded_file.name,
                    'registros': 0,
                    'periodo_inicio': 'N/A',
                    'periodo_fim': 'N/A',
                    'status': f'Erro: {str(e)}'
                })
            
            progress_bar.progress((i + 1) / total_files)
        
        status_text.text("Consolidação concluída com sucesso!")
        
        # Mostrar conflitos se detectados
        if self.conflicts_detected:
            self._show_conflicts()
        
        # Mostrar resumo do processamento
        self._show_file_processing_summary()
        
        return len(self.consolidated_data) > 0

    def _show_conflicts(self):
        """Mostra conflitos detectados entre arquivos"""
        st.markdown("---")
        st.markdown("### Conflitos Detectados")
        
        st.markdown(f"""
        <div class="warning-box">
            <h4>{len(self.conflicts_detected)} conflito(s) encontrado(s)</h4>
            <p>Timestamps idênticos em múltiplos arquivos. Usando dados do último arquivo processado.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Mostrar detalhes dos conflitos
        with st.expander("Ver Detalhes dos Conflitos"):
            for i, conflict in enumerate(self.conflicts_detected[:10]):  # Mostrar só os primeiros 10
                st.markdown(f"**Conflito {i+1}: {conflict['timestamp']}**")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("*Dados Anteriores:*")
                    st.json(conflict['dados_anteriores'])
                with col2:
                    st.markdown(f"*Dados Novos ({conflict['arquivo_atual']}):*")
                    st.json(conflict['dados_novos'])
                st.markdown("---")
            
            if len(self.conflicts_detected) > 10:
                st.info(f"Mostrando apenas os primeiros 10 conflitos de {len(self.conflicts_detected)} total.")

    def _find_closest_timestamp(self, target_time, available_timestamps):
        """
        Busca o timestamp mais próximo dentro da tolerância de ±10 minutos
        
        Args:
            target_time: datetime alvo (ex: 2025-06-22 10:00:00)
            available_timestamps: lista de timestamps disponíveis
            
        Returns:
            timestamp mais próximo ou None se nenhum estiver dentro da tolerância
        """
        target_time = pd.to_datetime(target_time)
        tolerance = timedelta(minutes=10)
        
        min_diff = timedelta.max
        closest_timestamp = None
        
        for ts in available_timestamps:
            ts_converted = pd.to_datetime(ts)
            diff = abs(ts_converted - target_time)
            
            # Verifica se está dentro da tolerância e é mais próximo
            if diff <= tolerance and diff < min_diff:
                min_diff = diff
                closest_timestamp = ts  # Retorna o timestamp original, não o convertido
        
        return closest_timestamp

    def update_excel_file(self, excel_file):
        """
        Atualiza Excel com Dados

        """
        if not self.consolidated_data:
            return False, "Nenhum dado processado!"
        
        try:
            # Salvar arquivo Excel temporariamente
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(excel_file.read())
                self.excel_path = tmp_file.name
            
            wb = load_workbook(self.excel_path)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Agrupar dados por mês
            monthly_data = {}
            for timestamp, data in self.consolidated_data.items():
                year_month = f"{timestamp.year}-{timestamp.month:02d}"
                if year_month not in monthly_data:
                    monthly_data[year_month] = {}
                monthly_data[year_month][timestamp] = data
            
            total_months = len(monthly_data)
            sheets_updated = 0
            total_cells_updated = 0
            
            # Processar cada mês
            for i, (year_month, month_timestamps) in enumerate(monthly_data.items()):
                year, month = year_month.split('-')
                month_num = int(month)
                
                status_text.text(f"Processando mês {month}/{year}...")
                
                # Buscar aba correspondente
                sheet_name = self._find_daily_analysis_sheet(wb.sheetnames, month_num)
                if not sheet_name:
                    st.warning(f"Aba para mês {month:02d} não encontrada!")
                    continue
                
                ws = wb[sheet_name]
                cells_updated = self._update_daily_analysis_exact(ws, month_timestamps, int(year), month_num)
                
                if cells_updated > 0:
                    sheets_updated += 1
                    total_cells_updated += cells_updated
                    self.processed_sheets.append(sheet_name)
                
                progress_bar.progress((i + 1) / total_months)
            
            # Salvar alterações
            wb.save(self.excel_path)
            status_text.text("Atualização concluída com sucesso!")
            
            if sheets_updated > 0:
                return True, f"Sucesso! {sheets_updated} aba(s) atualizada(s), {total_cells_updated} célula(s) preenchida(s)"
            else:
                return False, "Nenhuma aba compatível encontrada para atualização"
                
        except Exception as e:
            return False, f"Erro durante atualização: {e}"

    def _find_daily_analysis_sheet(self, sheet_names, month_num):
        """Encontra aba de análise diária para o mês"""
        month_str = f"{month_num:02d}"
        target_pattern = f"{month_str}-Analise Diaria"
        
        # Busca exata primeiro
        if target_pattern in sheet_names:
            return target_pattern
        
        # Busca por padrão similar
        for sheet_name in sheet_names:
            if month_str in sheet_name and "Analise Diaria" in sheet_name:
                return sheet_name
        
        return None

    def _update_daily_analysis_exact(self, ws, month_timestamps, year, month):
        """
        Atualiza análise diária usando busca EXATA
        """
        cells_updated = 0
        
        # Para cada horário da planilha (00:00 a 23:00)
        for hour in range(24):
            row_num = hour + 3  # Linha 3 = 00:00, Linha 4 = 01:00, etc.
            
            # Para cada dia do mês (1 a 31)
            for day in range(1, 32):
                # Construir timestamp alvo
                try:
                    target_datetime = datetime(year, month, day, hour, 0, 0)
                except ValueError:
                    # Dia inválido para o mês (ex: 31 de fevereiro)
                    continue
                
                # Buscar timestamp mais próximo dentro da tolerância
                available_timestamps = list(month_timestamps.keys())
                closest_timestamp = self._find_closest_timestamp(target_datetime, available_timestamps)
                
                if closest_timestamp is None:
                    # Nenhum dado dentro da tolerância - deixar vazio
                    continue
                
                # Obter dados do timestamp encontrado
                data = month_timestamps[closest_timestamp]
                
                # Atualizar cada variável
                for variable, value in data.items():
                    if value is None:
                        continue
                    
                    col_letter = self._get_column_for_variable_and_day(variable, day)
                    if col_letter:
                        try:
                            ws[f'{col_letter}{row_num}'] = value
                            cells_updated += 1
                        except Exception:
                            # Continua processamento mesmo com erro
                            pass
        
        return cells_updated

    def _get_column_for_variable_and_day(self, variable, day_number):
        """
        Calcula letra da coluna para análise diária
        """
        if variable not in self.column_mapping:
            return None
        
        # Verificar se o dia está no range válido (1-31)
        if day_number < 1 or day_number > 31:
            return None
        
        start_col_num = self.column_mapping[variable]['start_num']
        target_col_num = start_col_num + (day_number - 1)
        
        # Verificar se a coluna está dentro dos limites válidos
        if target_col_num > 187:  # Última coluna GE = 187
            return None
            
        col_letter = get_column_letter(target_col_num)
        return col_letter

    def _show_file_processing_summary(self):
        """Mostra resumo detalhado do processamento"""
        if hasattr(self, 'file_processing_info') and self.file_processing_info:
            st.markdown("---")
            st.markdown("### Resumo do Processamento")
            
            # Criar DataFrame com as informações
            df_files = pd.DataFrame(self.file_processing_info)
            
            # Calcular totais
            total_records = df_files['registros'].sum()
            total_files_success = len([f for f in self.file_processing_info if 'sucesso' in f['status'].lower()])
            total_timestamps = len(self.consolidated_data)
            
            # Mostrar métricas gerais
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>Arquivos Processados</h4>
                    <h2>{total_files_success}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>Registros Lidos</h4>
                    <h2>{total_records:,}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>Timestamps Únicos</h4>
                    <h2>{total_timestamps:,}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                st.markdown(f"""
                <div class="metric-card">
                    <h4>Conflitos Detectados</h4>
                    <h2>{len(self.conflicts_detected)}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            # Tabela detalhada
            st.markdown("#### Detalhes por Arquivo")
            df_display = df_files.copy()
            df_display.columns = ['Arquivo', 'Registros', 'Início', 'Fim', 'Status']
            df_display['Registros'] = df_display['Registros'].apply(lambda x: f"{x:,}" if x > 0 else "0")
            
            st.dataframe(df_display, use_container_width=True)

    def get_updated_excel_file(self):
        """Retorna o arquivo Excel atualizado"""
        if self.excel_path and os.path.exists(self.excel_path):
            with open(self.excel_path, 'rb') as f:
                return f.read()
        return None

    def show_data_preview_and_charts(self):
        """Mostra preview dos dados consolidados com gráficos para conferência"""
        if not self.consolidated_data:
            return
        
        st.markdown("---")
        st.markdown("### Análise dos Dados Consolidados")
        
        # Converter para DataFrame para visualização
        preview_data = []
        for timestamp, data in self.consolidated_data.items():
            row = {'Timestamp': timestamp}
            row.update(data)
            preview_data.append(row)
        
        if preview_data:
            df_preview = pd.DataFrame(preview_data)
            df_preview = df_preview.sort_values('Timestamp')
            
            # Estatísticas gerais
            st.markdown("#### Estatísticas Gerais")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                first_timestamp = min(self.consolidated_data.keys())
                last_timestamp = max(self.consolidated_data.keys())
                period_days = (last_timestamp - first_timestamp).days + 1
                st.metric("Período Total", f"{period_days} dias")
            
            with col2:
                timestamps_per_day = len(self.consolidated_data) / period_days if period_days > 0 else 0
                st.metric("Registros/Dia", f"{timestamps_per_day:.1f}")
            
            with col3:
                # Agrupar por mês
                months = set()
                for ts in self.consolidated_data.keys():
                    months.add(f"{ts.year}-{ts.month:02d}")
                st.metric("Meses Cobertos", len(months))
            
            # Gráficos para conferência das variáveis
            self._create_variable_charts(df_preview)
            
            # Preview da tabela de dados
            st.markdown("#### Preview dos Dados (Primeiros 100 registros)")
            st.dataframe(df_preview.head(100), use_container_width=True)
            
        else:
            st.info("Nenhum dado disponível para preview.")

    def _create_variable_charts(self, df):
        """Cria gráficos para conferência das variáveis meteorológicas"""
        st.markdown("#### Gráficos de Conferência das Variáveis")
        
        # Preparar dados para gráficos
        df_clean = df.dropna()
        
        if len(df_clean) == 0:
            st.warning("Não há dados suficientes para gerar gráficos.")
            return
        
        # Gráfico 1: Temperatura ao longo do tempo
        with st.container():
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.markdown("**Temperatura (°C)**")
            
            fig_temp = px.line(df_clean, x='Timestamp', y='Temperatura',
                              title='Variação da Temperatura ao Longo do Tempo',
                              color_discrete_sequence=['#00529C'])
            fig_temp.update_layout(
                xaxis_title="Data/Hora",
                yaxis_title="Temperatura (°C)",
                height=400,
                showlegend=False
            )
            st.plotly_chart(fig_temp, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Gráfico 2: Piranômetros (Radiação Solar)
        with st.container():
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.markdown("**Radiação Solar (kW/m²)**")
            
            fig_pir = go.Figure()
            
            if 'Piranometro_1' in df_clean.columns:
                fig_pir.add_trace(go.Scatter(x=df_clean['Timestamp'], y=df_clean['Piranometro_1'],
                                           mode='lines', name='Piranômetro 1', line=dict(color='#FF6B35')))
            
            if 'Piranometro_2' in df_clean.columns:
                fig_pir.add_trace(go.Scatter(x=df_clean['Timestamp'], y=df_clean['Piranometro_2'],
                                           mode='lines', name='Piranômetro 2', line=dict(color='#F7931E')))
            
            if 'Piranometro_Alab' in df_clean.columns:
                fig_pir.add_trace(go.Scatter(x=df_clean['Timestamp'], y=df_clean['Piranometro_Alab'],
                                           mode='lines', name='Piranômetro Alabiotico', line=dict(color='#FFD23F')))
            
            fig_pir.update_layout(
                title='Radiação Solar - Comparação dos Piranômetros',
                xaxis_title="Data/Hora",
                yaxis_title="Radiação Solar (kW/m²)",
                height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_pir, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Gráfico 3: Umidade e Velocidade do Vento
        with st.container():
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.markdown("**Umidade Relativa e Velocidade do Vento**")
            
            # Criar subplots com eixos Y duplos
            fig_combined = make_subplots(specs=[[{"secondary_y": True}]])
            
            if 'Umidade_Relativa' in df_clean.columns:
                fig_combined.add_trace(
                    go.Scatter(x=df_clean['Timestamp'], y=df_clean['Umidade_Relativa'],
                             mode='lines', name='Umidade Relativa (%)', line=dict(color='#4A90E2')),
                    secondary_y=False,
                )
            
            if 'Velocidade_Vento' in df_clean.columns:
                fig_combined.add_trace(
                    go.Scatter(x=df_clean['Timestamp'], y=df_clean['Velocidade_Vento'],
                             mode='lines', name='Velocidade do Vento (m/s)', line=dict(color='#50C878')),
                    secondary_y=True,
                )
            
            # Configurar eixos Y
            fig_combined.update_xaxes(title_text="Data/Hora")
            fig_combined.update_yaxes(title_text="Umidade Relativa (%)", secondary_y=False)
            fig_combined.update_yaxes(title_text="Velocidade do Vento (m/s)", secondary_y=True)
            
            fig_combined.update_layout(
                title='Umidade Relativa e Velocidade do Vento',
                height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_combined, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Estatísticas descritivas
        st.markdown("#### Estatísticas Descritivas")
        
        # Selecionar apenas colunas numéricas
        numeric_cols = ['Temperatura', 'Piranometro_1', 'Piranometro_2', 'Piranometro_Alab', 'Umidade_Relativa', 'Velocidade_Vento']
        available_cols = [col for col in numeric_cols if col in df_clean.columns]
        
        if available_cols:
            stats_df = df_clean[available_cols].describe().round(3)
            stats_df.index = ['Contagem', 'Média', 'Desvio Padrão', 'Mínimo', '25%', '50% (Mediana)', '75%', 'Máximo']
            
            st.markdown('<div class="stats-container">', unsafe_allow_html=True)
            st.dataframe(stats_df, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Distribuição dos dados por hora do dia
        st.markdown("#### Distribuição por Hora do Dia")
        
        if len(df_clean) > 0:
            df_clean['Hora'] = df_clean['Timestamp'].dt.hour
            
            # Gráfico de distribuição por hora
            hourly_stats = df_clean.groupby('Hora')[available_cols].mean().reset_index()
            
            fig_hourly = go.Figure()
            
            colors = ['#00529C', '#FF6B35', '#F7931E', '#FFD23F', '#4A90E2', '#50C878']
            
            for i, col in enumerate(available_cols):
                if col in hourly_stats.columns:
                    fig_hourly.add_trace(go.Scatter(
                        x=hourly_stats['Hora'], 
                        y=hourly_stats[col],
                        mode='lines+markers',
                        name=col.replace('_', ' ').title(),
                        line=dict(color=colors[i % len(colors)])
                    ))
            
            fig_hourly.update_layout(
                title='Valores Médios por Hora do Dia',
                xaxis_title="Hora do Dia",
                yaxis_title="Valores Médios",
                height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            
            st.plotly_chart(fig_hourly, use_container_width=True)


def main():
    # Cabeçalho principal
    st.markdown("""
    <div class="main-header">
        <h1>Medições Usina Geradora Floriano</h1>
        <p>Processador de Dados de Medicao</p>
        <p><small>Busca Pontual | Tolerância ±10min | Zero Inferências</small></p>
    </div>
    """, unsafe_allow_html=True)
    
    # Inicializar o processador
    if 'processor' not in st.session_state:
        st.session_state.processor = ExactWeatherProcessor()
    
    # Sidebar com instruções
    with st.sidebar:
        st.markdown("### Instruções de Uso")
        st.markdown("""
        **Passo 1:** Upload do arquivo Excel anual
        
        **Passo 2:** Upload dos arquivos .dat (múltiplos)
        
        **Passo 3:** Clique em "Processar Dados"
        
        **Passo 4:** Baixe o Excel atualizado
        """)
        
        st.markdown("---")
        st.markdown("### Funcionalidades")
        st.markdown("""
        **Características:**
        - Busca pontual de dados (sem médias)
        - Tolerância de ±10 minutos
        - Zero inferências ou preenchimentos
        - Detecção de conflitos entre arquivos
        - Mapeamento preciso por timestamp
        - Gráficos de conferência das variáveis
        
        **Lógica de Busca:**
        - Para 10:00 → busca entre 09:50 e 10:10
        - Prioriza timestamp mais próximo
        - Deixa vazio se não há dados na tolerância
        """)
        
        st.markdown("---")
        st.markdown("### Mapeamento de Colunas")
        st.markdown("""
        - **Temperatura**: Colunas B-AF (Dias 1-31)
        - **Piranômetro 1**: Colunas AG-BK (Dias 1-31)
        - **Piranômetro 2**: Colunas BL-CP (Dias 1-31)
        - **Piranômetro Alabiotico**: Colunas CQ-DU (Dias 1-31)
        - **Umidade**: Colunas DV-EZ (Dias 1-31)
        - **Vento**: Colunas FA-GE (Dias 1-31)
        """)
    
    # Layout principal
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown("### Upload do Excel Anual")
        excel_file = st.file_uploader(
            "Selecione o arquivo Excel anual",
            type=['xlsx', 'xls'],
            help="Arquivo Excel com as abas XX-Analise Diaria"
        )
    
    with col2:
        st.markdown("### Upload dos Arquivos .dat")
        dat_files = st.file_uploader(
            "Selecione os arquivos .dat (múltiplos)",
            type=['dat'],
            accept_multiple_files=True,
            help="Arquivos de dados meteorológicos (.dat) com timestamps de 10 em 10 minutos"
        )
    
    # Informações sobre os arquivos carregados
    if dat_files:
        st.markdown("### Arquivos .dat Carregados")
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
            if st.button("Processar Dados", use_container_width=True):
                with st.spinner("Processando dados com busca pontual..."):
                    # Processar arquivos .dat
                    success = st.session_state.processor.process_dat_files(dat_files)
                    
                    if success:
                        st.success("Arquivos .dat processados e consolidados com sucesso!")
                        
                        # Mostrar preview dos dados com gráficos
                        st.session_state.processor.show_data_preview_and_charts()
                        
                        # Atualizar Excel
                        st.markdown("### Atualizando Excel...")
                        excel_file.seek(0)  # Reset file pointer
                        success, message = st.session_state.processor.update_excel_file(excel_file)
                        
                        if success:
                            st.success(f"{message}")
                            
                            # Informações sobre abas atualizadas
                            if st.session_state.processor.processed_sheets:
                                st.markdown("### Abas Atualizadas")
                                for sheet in st.session_state.processor.processed_sheets:
                                    st.markdown(f"- {sheet}")
                            
                            # Botão de download
                            updated_excel = st.session_state.processor.get_updated_excel_file()
                            if updated_excel:
                                st.markdown("### Download do Arquivo Atualizado")
                                st.download_button(
                                    label="Baixar Excel Atualizado",
                                    data=updated_excel,
                                    file_name=f"analise_medicoes_exato_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                        else:
                            st.error(f"{message}")
                    else:
                        st.error("Erro ao processar arquivos .dat")
    
    # Informações adicionais
    if not excel_file or not dat_files:
        st.markdown("---")
        st.markdown("### Aguardando Arquivos")
        missing = []
        if not excel_file:
            missing.append("Arquivo Excel anual")
        if not dat_files:
            missing.append("Arquivos .dat")
        
        st.info(f"Por favor, faça upload dos seguintes arquivos: {', '.join(missing)}")
        
        if not dat_files:
            st.markdown("""
            **Sobre a Busca:**
            - Busca dados pontuais sem fazer médias
            - Tolerância de ±10 minutos para cada horário
            - Não preenche dados que não existem
            - Detecta e alerta sobre conflitos entre arquivos
            - Mapeia diretamente timestamp → célula da planilha
            - Gera gráficos para conferência visual dos dados
            """)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>Processador de Dados Meteorológicos | Usina Geradora Floriano</p>
        <p><small>Versão Busca Pontual com Tolerância ±10min</small></p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()

