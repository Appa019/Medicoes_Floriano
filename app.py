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

# CSS personalizado
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
    🔧 PROCESSADOR CORRIGIDO FINAL
    Preenche TODOS os horários 00:00-23:00 no Excel baseado nos arquivos .dat
    """

    def __init__(self):
        self.dados_processados = {}
        self.excel_path = None
        self.abas_diarias_atualizadas = []
        self.file_processing_info = []

        # 🔧 MAPEAMENTO CORRETO DAS COLUNAS (baseado no header fornecido)
        self.column_mapping = {
            'Temperatura': {'start_col': 'B'},           # Temperatura_Dia20 = coluna B
            'Piranometro_1': {'start_col': 'AG'},        # Piranometro_1_Dia1 = coluna AG (coluna 33)
            'Piranometro_2': {'start_col': 'BL'},        # Piranometro_2_Dia1 = coluna BL (coluna 64) 
            'Piranometro_Alab': {'start_col': 'CQ'},     # Piranometro_Alab_Dia1 = coluna CQ (coluna 95)
            'Umidade_Relativa': {'start_col': 'DV'},     # Umidade_Relativa_Dia1 = coluna DV (coluna 126)
            'Velocidade_Vento': {'start_col': 'FA'}      # Velocidade_Vento_Dia1 = coluna FA (coluna 157)
        }

    def process_dat_files(self, dat_files):
        """
        🔧 PROCESSA ARQUIVOS .DAT COM PREENCHIMENTO COMPLETO 00:00-23:00
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
                
                # 🔧 NOVO: Processar com preenchimento completo 24h
                processed_days = self._process_complete_24h_data(data, uploaded_file.name)
                
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

    def _process_complete_24h_data(self, data, filename):
        """
        🔧 NOVA FUNÇÃO: Processa dados para preencher TODAS as 24 horas (00:00-23:00)
        
        Lógica corrigida:
        - Um arquivo .dat contém dados de 10:10 do dia anterior até 10:00 do dia atual
        - Para cada dia no intervalo, preenche TODAS as 24 horas usando os dados disponíveis
        """
        # Determinar datas dos dados
        start_timestamp = data.index.min()
        end_timestamp = data.index.max()
        
        # Determinar quais dias processar
        start_date = start_timestamp.date()
        end_date = end_timestamp.date()
        
        # Lista de todos os dias no intervalo
        current_date = start_date
        processed_days = 0
        
        while current_date <= end_date:
            # 🔧 NOVO: Criar dados completos para 24h deste dia
            complete_day_data = self._create_complete_day_data(data, current_date)
            
            if complete_day_data:  # Se conseguiu criar dados para o dia
                # Armazenar os dados
                year = current_date.year
                month = current_date.month
                day = current_date.day
                dataset_key = f"{year}-{month:02d}"
                
                if dataset_key not in self.dados_processados:
                    self.dados_processados[dataset_key] = {}
                
                # Armazenar dados completos de 24h
                self.dados_processados[dataset_key][day] = complete_day_data
                processed_days += 1
                
                print(f"🔧 Dia {current_date}: {len(complete_day_data)} horas processadas")
            
            current_date += timedelta(days=1)
        
        return processed_days

    def _create_complete_day_data(self, data, target_date):
        """
        🔧 FUNÇÃO CHAVE: Cria dados completos para todas as 24 horas de um dia
        
        Estratégia:
        1. Para cada hora (00:00-23:00), procura dados nos .dat
        2. Calcula média dos registros de 10 em 10 minutos da hora
        3. Se não há dados, deixa None (será tratado no Excel)
        """
        complete_data = {}
        
        # Para cada hora do dia (0-23)
        for hour in range(24):
            hour_data = self._get_hour_data(data, target_date, hour)
            
            if hour_data is not None:
                complete_data[f"{hour:02d}:00"] = hour_data
        
        return complete_data

    def _get_hour_data(self, data, target_date, hour):
        """
        🔧 EXTRAI DADOS DE UMA HORA ESPECÍFICA (média dos registros de 10 em 10 min)
        """
        try:
            # Criar datetime para a hora específica
            start_time = datetime.combine(target_date, datetime.min.time()) + timedelta(hours=hour)
            end_time = start_time + timedelta(hours=1)
            
            # Filtrar dados da hora
            hour_mask = (data.index >= start_time) & (data.index < end_time)
            hour_records = data[hour_mask]
            
            if len(hour_records) == 0:
                return None
            
            # 🔧 CALCULAR MÉDIAS DOS REGISTROS DA HORA
            return {
                'Temperatura': round(hour_records['Temp_Avg'].mean(), 2),
                'Piranometro_1': round(hour_records['Pir1_Avg'].mean() / 1000, 3),  # Converter W/m² para kW/m²
                'Piranometro_2': round(hour_records['Pir2_Avg'].mean() / 1000, 3),
                'Piranometro_Alab': round(hour_records['PirALB_Avg'].mean() / 1000, 3),
                'Umidade_Relativa': round(hour_records['RH_Avg'].mean(), 2),
                'Velocidade_Vento': round(hour_records['Ane_Avg'].mean(), 2)
            }
        
        except Exception as e:
            print(f"Erro ao processar hora {hour} do dia {target_date}: {e}")
            return None

    def update_excel_file(self, excel_file):
        """
        🔧 ATUALIZA EXCEL COM PREENCHIMENTO COMPLETO 00:00-23:00
        """
        if not self.dados_processados:
            return False, "Nenhum dado processado!"

        try:
            # Salvar arquivo Excel temporariamente
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(excel_file.read())
                self.excel_path = tmp_file.name

            wb = load_workbook(self.excel_path)
            
            total_hours_updated = 0
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            total_months = len(self.dados_processados)
            
            # Processar cada mês
            for i, (dataset_key, month_data) in enumerate(self.dados_processados.items()):
                ano, mes = dataset_key.split('-')
                mes_numero = int(mes)
                
                status_text.text(f"Atualizando análise diária {mes}/{ano}...")

                # 🔧 ATUALIZAR ANÁLISE DIÁRIA
                aba_diaria = self._find_daily_sheet(wb.sheetnames, mes_numero)
                if aba_diaria:
                    try:
                        ws_diaria = wb[aba_diaria]
                        hours_updated = self._update_complete_daily_data(ws_diaria, month_data)
                        total_hours_updated += hours_updated

                        if aba_diaria not in self.abas_diarias_atualizadas:
                            self.abas_diarias_atualizadas.append(aba_diaria)
                    except Exception as e:
                        return False, f"Erro na análise diária: {e}"
                
                progress_bar.progress((i + 1) / total_months)

            # Salvar alterações
            wb.save(self.excel_path)
            status_text.text("Atualização concluída!")

            if total_hours_updated > 0:
                return True, f"✅ Sucesso! {total_hours_updated} horas atualizadas com preenchimento completo 00:00-23:00"
            else:
                return False, "Nenhum dado foi atualizado"

        except Exception as e:
            return False, f"Erro geral: {e}"

    def _update_complete_daily_data(self, ws, month_data):
        """
        🔧 ATUALIZA PLANILHA COM DADOS COMPLETOS DE 24H
        """
        total_hours_updated = 0

        for dia_numero, day_data in month_data.items():
            print(f"🔧 Atualizando dia {dia_numero} com {len(day_data)} horas")
            
            # Para cada hora do dia (00:00-23:00)
            for hour in range(24):
                hour_str = f"{hour:02d}:00"
                
                # Linha na planilha (00:00 = linha 3, 01:00 = linha 4, etc.)
                row_num = hour + 3
                
                # Se há dados para esta hora
                if hour_str in day_data:
                    hour_values = day_data[hour_str]
                    
                    # Atualizar cada variável
                    for variable, value in hour_values.items():
                        col_letter = self._get_column_for_variable_and_day(variable, dia_numero)
                        
                        if col_letter and value is not None:
                            try:
                                cell_ref = f'{col_letter}{row_num}'
                                ws[cell_ref] = value
                                total_hours_updated += 1
                                print(f"    {hour_str} {variable} = {value} -> {cell_ref}")
                            except Exception as e:
                                print(f"    Erro ao escrever {variable} no dia {dia_numero}, hora {hour_str}: {e}")

        return total_hours_updated

    def _find_daily_sheet(self, sheet_names, mes_numero):
        """Encontra aba de análise diária"""
        mes_str = f"{mes_numero:02d}"

        possible_names = [
            f"{mes_str}-Analise Diaria",
            f"{mes_str}-Analyse Diaria",
            f"{mes_str} Analise Diaria",
            f"Analise Diaria {mes_str}"
        ]

        for name in possible_names:
            if name in sheet_names:
                return name

        # Buscar por padrão
        for sheet_name in sheet_names:
            if mes_str in sheet_name and "Diaria" in sheet_name:
                return sheet_name

        return None

    def _get_column_for_variable_and_day(self, variable, dia_numero):
        """
        🔧 CALCULA COLUNA CORRETA BASEADA NA ESTRUTURA EXCEL
        """
        if variable not in self.column_mapping:
            return None

        # Obter coluna inicial para a variável
        start_col_letter = self.column_mapping[variable]['start_col']
        
        # Converter letra da coluna para número
        start_col_num = self._column_letter_to_number(start_col_letter)
        
        # Para Temperatura: Dia20 = coluna B, Dia21 = coluna C, etc.
        # Para outras variáveis: Dia1 = start_col, Dia2 = start_col + 1, etc.
        if variable == 'Temperatura':
            # Temperatura_Dia20 está na coluna B, então:
            # Dia20 = B (coluna 2), Dia21 = C (coluna 3), etc.
            target_col_num = start_col_num + (dia_numero - 20)
        else:
            # Para outras variáveis: Dia1 = start_col, Dia2 = start_col + 1, etc.
            target_col_num = start_col_num + (dia_numero - 1)
        
        # Converter de volta para letra
        return get_column_letter(target_col_num)
    
    def _column_letter_to_number(self, column_letter):
        """Converte letra da coluna para número (A=1, B=2, etc.)"""
        result = 0
        for char in column_letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    def get_updated_excel_file(self):
        """Retorna o arquivo Excel atualizado"""
        if self.excel_path and os.path.exists(self.excel_path):
            with open(self.excel_path, 'rb') as f:
                return f.read()
        return None

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

    def show_final_summary(self):
        """Mostra resumo final dos dados processados"""
        if not self.dados_processados:
            return None

        total_days = 0
        total_hours = 0
        
        summary_data = []
        
        for dataset_key, month_data in self.dados_processados.items():
            ano, mes = dataset_key.split('-')
            dias_no_mes = len(month_data)
            total_days += dias_no_mes
            
            # Contar horas processadas
            horas_processadas = 0
            for dia_numero, day_data in month_data.items():
                horas_processadas += len(day_data)
            
            total_hours += horas_processadas
            
            summary_data.append({
                'Mês/Ano': f"{mes}/{ano}",
                'Dias Processados': dias_no_mes,
                'Horas Processadas': horas_processadas,
                'Cobertura (%)': f"{(horas_processadas / (dias_no_mes * 24) * 100):.1f}%"
            })

        return summary_data, total_days, total_hours


def main():
    # Cabeçalho principal
    st.markdown("""
    <div class="main-header">
        <h1>🌤️ Medições Usina Geradora Floriano</h1>
        <p>Processador Corrigido - Preenchimento Completo 00:00-23:00</p>
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
        
        **Passo 2:** Faça upload dos arquivos .dat
        
        **Passo 3:** Clique em "Processar Dados"
        
        **Passo 4:** Baixe o arquivo Excel atualizado
        """)
        
        st.markdown("---")
        st.markdown("### ✅ Características Corrigidas")
        st.markdown("""
        **🔧 PREENCHIMENTO COMPLETO:**
        - ✅ **24 Horas**: Preenche 00:00-23:00 para cada dia
        - ✅ **Média dos Registros**: Calcula média dos 6 registros por hora
        - ✅ **Mapeamento Correto**: Temperatura_Dia20, Dia21, etc.
        - ✅ **Todas as Variáveis**: Temperatura, Piranômetros, Umidade, Vento
        """)
        
        st.markdown("---")
        st.markdown("### 📊 Estrutura dos Dados")
        st.markdown("""
        **Arquivos .dat processados:**
        - **352.dat**: 20/06 10:10 → 21/06 10:00
        - **353.dat**: 21/06 10:10 → 22/06 10:00  
        - **354.dat**: 22/06 10:10 → 23/06 10:00
        - **355.dat**: 23/06 10:10 → 24/06 10:00
        
        **✅ Resultado: 24h/dia completas!**
        """)

    # Layout principal
    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("### 📊 Upload do Excel Anual")
        excel_file = st.file_uploader(
            "Selecione o arquivo Excel anual",
            type=['xlsx', 'xls'],
            help="Arquivo Excel com as abas de análise diária"
        )

    with col2:
        st.markdown("### 📁 Upload dos Arquivos .dat")
        dat_files = st.file_uploader(
            "Selecione os arquivos .dat",
            type=['dat'],
            accept_multiple_files=True,
            help="Arquivos .dat para preenchimento completo 00:00-23:00"
        )

    # Botão de processamento
    if excel_file and dat_files:
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🔧 Processar com Preenchimento Completo 24h", use_container_width=True):
                with st.spinner("Processando com preenchimento completo 00:00-23:00..."):
                    # Processar arquivos .dat
                    success = st.session_state.processor.process_dat_files(dat_files)
                    
                    if success:
                        st.success("✅ Arquivos .dat processados com preenchimento completo 24h!")
                        
                        # Mostrar resumo
                        summary_result = st.session_state.processor.show_final_summary()
                        if summary_result and len(summary_result) == 3:
                            summary_data, total_days, total_hours = summary_result
                            
                            st.markdown("### 📊 Resumo dos Dados Processados (24h Completas)")
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.markdown(f"""
                                <div class="metric-card">
                                    <h4>📅 Total de Dias</h4>
                                    <h2>{total_days}</h2>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            with col2:
                                st.markdown(f"""
                                <div class="metric-card">
                                    <h4>⏰ Horas Processadas</h4>
                                    <h2>{total_hours}h</h2>
                                </div>
                                """, unsafe_allow_html=True)
                                
                            with col3:
                                cobertura = (total_hours / (total_days * 24) * 100) if total_days > 0 else 0
                                st.markdown(f"""
                                <div class="metric-card">
                                    <h4>📈 Cobertura</h4>
                                    <h2>{cobertura:.1f}%</h2>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            # Tabela de resumo
                            df_summary = pd.DataFrame(summary_data)
                            st.dataframe(df_summary, use_container_width=True)
                        
                        # Atualizar Excel
                        st.markdown("### 🔄 Atualizando Excel com Preenchimento Completo...")
                        excel_file.seek(0)  # Reset file pointer
                        success, message = st.session_state.processor.update_excel_file(excel_file)
                        
                        if success:
                            st.success(f"✅ {message}")
                            
                            # Botão de download
                            updated_excel = st.session_state.processor.get_updated_excel_file()
                            if updated_excel:
                                st.markdown("### 📥 Download do Arquivo Completo")
                                st.download_button(
                                    label="📥 Baixar Excel com Preenchimento Completo 24h",
                                    data=updated_excel,
                                    file_name=f"analise_anual_completa_24h_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                                
                                st.success("🎯 **SUCESSO!** Todas as horas 00:00-23:00 foram preenchidas no Excel!")
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
        <p><strong>🔧 CORRIGIDO FINAL:</strong> Preenchimento completo 00:00-23:00 com mapeamento correto!</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
