import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os
from datetime import datetime, timedelta
import warnings
import tempfile

warnings.filterwarnings('ignore')

# Configuração da página
st.set_page_config(
    page_title="Medições Usina Geradora Floriano",
    page_icon="🌤️",
    layout="wide"
)

# CSS simples
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

class SimpleWeatherProcessor:
    """
    Processador SIMPLES de dados meteorológicos
    Foco em funcionalidade básica sem complexidades
    """

    def __init__(self):
        self.dados_processados = {}
        self.excel_path = None

    def process_dat_files(self, dat_files):
        """
        Processa arquivos .dat de forma simples
        """
        if not dat_files:
            return False
            
        st.info(f"Processando {len(dat_files)} arquivo(s)...")
        
        for uploaded_file in dat_files:
            try:
                # Ler arquivo .dat
                uploaded_file.seek(0)
                data = pd.read_csv(uploaded_file, skiprows=4, parse_dates=[0])

                # Renomear colunas básicas
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
                
                # Processar dados por dia
                self._process_simple_data(data)
                
                st.success(f"✅ {uploaded_file.name}: {len(data)} registros processados")

            except Exception as e:
                st.error(f"❌ Erro ao processar {uploaded_file.name}: {str(e)}")
                continue

        return bool(self.dados_processados)

    def _process_simple_data(self, data):
        """
        Processamento simples dos dados
        """
        # Agrupar por data
        data['date'] = data.index.date
        
        for date in data['date'].unique():
            day_data = data[data['date'] == date]
            
            year = date.year
            month = date.month
            day = date.day
            
            # Criar chave do mês
            month_key = f"{year}-{month:02d}"
            
            if month_key not in self.dados_processados:
                self.dados_processados[month_key] = {}
            
            # Processar dados horários simples
            hourly_data = {}
            
            # Agrupar por hora
            day_data['hour'] = day_data.index.hour
            
            for hour in range(24):
                hour_records = day_data[day_data['hour'] == hour]
                
                if len(hour_records) > 0:
                    hourly_data[f"{hour:02d}:00"] = {
                        'Temperatura': round(hour_records['Temp_Avg'].mean(), 2),
                        'Piranometro_1': round(hour_records['Pir1_Avg'].mean() / 1000, 3),
                        'Piranometro_2': round(hour_records['Pir2_Avg'].mean() / 1000, 3),
                        'Piranometro_Alab': round(hour_records['PirALB_Avg'].mean() / 1000, 3),
                        'Umidade_Relativa': round(hour_records['RH_Avg'].mean(), 2),
                        'Velocidade_Vento': round(hour_records['Ane_Avg'].mean(), 2)
                    }
            
            # Armazenar dados do dia
            self.dados_processados[month_key][day] = hourly_data

    def update_excel_file(self, excel_file):
        """
        Atualiza Excel de forma simples
        """
        if not self.dados_processados:
            return False, "Nenhum dado processado!"

        try:
            # Salvar arquivo temporariamente
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(excel_file.read())
                self.excel_path = tmp_file.name

            wb = load_workbook(self.excel_path)
            total_updated = 0
            
            # Processar cada mês
            for month_key, month_data in self.dados_processados.items():
                year, month = month_key.split('-')
                month_num = int(month)
                
                # Encontrar aba diária
                sheet_name = self._find_daily_sheet(wb.sheetnames, month_num)
                
                if sheet_name:
                    ws = wb[sheet_name]
                    updated = self._update_simple_daily_data(ws, month_data)
                    total_updated += updated

            # Salvar arquivo
            wb.save(self.excel_path)
            
            if total_updated > 0:
                return True, f"✅ {total_updated} células atualizadas!"
            else:
                return False, "Nenhum dado foi atualizado"

        except Exception as e:
            return False, f"Erro: {str(e)}"

    def _find_daily_sheet(self, sheet_names, month_num):
        """
        Encontra aba diária simples
        """
        month_str = f"{month_num:02d}"
        
        # Procurar padrões simples
        for name in sheet_names:
            if month_str in name and "Diaria" in name:
                return name
        
        return None

    def _update_simple_daily_data(self, ws, month_data):
        """
        Atualiza dados na planilha de forma simples
        """
        total_updated = 0
        
        # Mapeamento simples de colunas (baseado na estrutura do documento)
        column_mapping = {
            'Temperatura': 2,  # Coluna B
            'Piranometro_1': 33,  # Coluna AG
            'Piranometro_2': 64,  # Coluna BL
            'Piranometro_Alab': 95,  # Coluna CQ
            'Umidade_Relativa': 126,  # Coluna DV
            'Velocidade_Vento': 157  # Coluna FA
        }
        
        for day_num, day_data in month_data.items():
            for hour_str, hour_data in day_data.items():
                # Calcular linha (00:00 = linha 3, 01:00 = linha 4, etc.)
                hour_num = int(hour_str[:2])
                row_num = hour_num + 3
                
                # Atualizar cada variável
                for variable, value in hour_data.items():
                    if variable in column_mapping:
                        # Calcular coluna para o dia específico
                        base_col = column_mapping[variable]
                        
                        # Ajuste especial para temperatura (Dia20, Dia21, etc.)
                        if variable == 'Temperatura':
                            col_num = base_col + (day_num - 20)
                        else:
                            col_num = base_col + (day_num - 1)
                        
                        # Verificar se a coluna é válida
                        if col_num > 0:
                            try:
                                col_letter = get_column_letter(col_num)
                                cell_ref = f'{col_letter}{row_num}'
                                ws[cell_ref] = value
                                total_updated += 1
                            except:
                                pass  # Ignorar erros silenciosamente
        
        return total_updated

    def get_updated_excel_file(self):
        """
        Retorna arquivo Excel atualizado
        """
        if self.excel_path and os.path.exists(self.excel_path):
            with open(self.excel_path, 'rb') as f:
                return f.read()
        return None

    def show_summary(self):
        """
        Mostra resumo simples
        """
        if not self.dados_processados:
            return None, 0, 0

        total_days = 0
        total_hours = 0
        summary_data = []
        
        for month_key, month_data in self.dados_processados.items():
            year, month = month_key.split('-')
            days_count = len(month_data)
            
            hours_count = 0
            for day_data in month_data.values():
                hours_count += len(day_data)
            
            total_days += days_count
            total_hours += hours_count
            
            summary_data.append({
                'Mês/Ano': f"{month}/{year}",
                'Dias': days_count,
                'Horas': hours_count
            })

        return summary_data, total_days, total_hours


def main():
    # Cabeçalho
    st.markdown("""
    <div class="main-header">
        <h1>🌤️ Medições Usina Geradora Floriano</h1>
        <p>Processador SIMPLES de Dados Meteorológicos</p>
    </div>
    """, unsafe_allow_html=True)

    # Inicializar processador
    if 'processor' not in st.session_state:
        st.session_state.processor = SimpleWeatherProcessor()

    # Instruções simples
    st.markdown("### 📋 Como usar:")
    st.markdown("""
    1. Faça upload do arquivo Excel
    2. Faça upload dos arquivos .dat
    3. Clique em "Processar"
    4. Baixe o arquivo atualizado
    """)

    # Upload de arquivos
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 📊 Excel Anual")
        excel_file = st.file_uploader(
            "Arquivo Excel",
            type=['xlsx', 'xls']
        )

    with col2:
        st.markdown("### 📁 Arquivos .dat")
        dat_files = st.file_uploader(
            "Arquivos .dat",
            type=['dat'],
            accept_multiple_files=True
        )

    # Processamento
    if excel_file and dat_files:
        st.markdown("---")
        
        if st.button("🚀 Processar Dados", use_container_width=True, type="primary"):
            with st.spinner("Processando..."):
                # Processar .dat
                success = st.session_state.processor.process_dat_files(dat_files)
                
                if success:
                    # Mostrar resumo
                    summary_data, total_days, total_hours = st.session_state.processor.show_summary()
                    
                    if summary_data:
                        st.success("✅ Dados processados!")
                        
                        # Métricas simples
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            st.markdown(f"""
                            <div class="metric-card">
                                <h4>📅 Dias</h4>
                                <h2>{total_days}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col2:
                            st.markdown(f"""
                            <div class="metric-card">
                                <h4>⏰ Horas</h4>
                                <h2>{total_hours}</h2>
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
                        
                        # Tabela resumo
                        df_summary = pd.DataFrame(summary_data)
                        st.dataframe(df_summary, use_container_width=True)
                    
                    # Atualizar Excel
                    st.info("Atualizando Excel...")
                    excel_file.seek(0)
                    success, message = st.session_state.processor.update_excel_file(excel_file)
                    
                    if success:
                        st.success(message)
                        
                        # Download
                        updated_excel = st.session_state.processor.get_updated_excel_file()
                        if updated_excel:
                            st.download_button(
                                label="📥 Baixar Excel Atualizado",
                                data=updated_excel,
                                file_name=f"medicoes_floriano_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                    else:
                        st.error(message)
                else:
                    st.error("❌ Erro ao processar arquivos")
    else:
        st.info("📤 Aguardando upload dos arquivos...")

    # Footer simples
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666;">
        <p>🌤️ Processador Simples | Usina Geradora Floriano</p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
