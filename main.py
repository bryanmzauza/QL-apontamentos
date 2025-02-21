import pandas as pd
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime

class QLApontamentos():
    """
    Classe para manipulação e análise de apontamentos da Quimlabor Jr.
    Inclui carregamento, limpeza, processamento de dados e exportação para PDF.
    """
    def __init__(self):
        self.data = None

    def get_data(self, file_path):
        """
        Carrega um arquivo Excel e retorna um DataFrame.
        :param file_path: Caminho do arquivo Excel.
        :return: DataFrame com os dados carregados.
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"O arquivo '{file_path}' não foi encontrado.")
        
        try:
            self.data = pd.read_excel(file_path)
        except Exception as e:
            raise ValueError(f"Erro ao carregar o arquivo Excel: {e}")
        
        return self.data

    def clean_data(self, data):
        """
        Limpa e estrutura os dados.
        - Renomeia a coluna 'Membro - name' para 'Membro'.
        - Mantém apenas colunas relevantes.
        - Separa múltiplos membros em linhas distintas.
        - Converte colunas de data para o formato datetime.
        :param data: DataFrame com os dados brutos.
        :return: DataFrame limpo e estruturado.
        """
        data.rename(columns={'Membro - name': 'Membro'}, inplace=True)
        data = data[['Membro', 'Categoria', 'Reunião', 'Atribuição Interna', 'Data - start', 'Data - end', 'Duração', 'Detalhamento']]
        data['Membro'] = data['Membro'].str.split(r'[;]')
        data = data.explode('Membro')
        data['Membro'] = data['Membro'].str.strip()
        data['Data - start'] = pd.to_datetime(data['Data - start'], errors='coerce')
        data['Data - end'] = pd.to_datetime(data['Data - end'], errors='coerce')
        return data

    def sum_hours(self, data, start_date, end_date):
        """
        Calcula a soma das horas de cada membro dentro do período especificado.
        :param data: DataFrame de apontamentos limpos.
        :param start_date: Data inicial do filtro.
        :param end_date: Data final do filtro.
        :return: DataFrame com total de horas por membro.
        """
        mask = (data['Data - start'] >= start_date) & (data['Data - start'] <= end_date)
        filtered_data = data[mask]
        total_hours = filtered_data.groupby('Membro')['Duração'].sum().reset_index()
        total_hours = total_hours.sort_values(by='Duração', ascending=False)
        return total_hours
    
    def format_hours(self, data):
        """
        Formata a coluna 'Duração' para exibir horas e minutos de forma legível.
        :param data: DataFrame com soma de horas por membro.
        :return: DataFrame formatado com horas e minutos.
        """
        data['Duração'] = data['Duração'].apply(lambda x: f"{int(x)}h{int((x - int(x)) * 60):02d}min")
        data.rename(columns={'Duração': 'Total de Horas'}, inplace=True)
        return data

    def add_expected_hours(self, data, start_date):
        """
        Adiciona uma coluna com a quantidade de horas esperadas.
        :param data: DataFrame com soma de horas por membro.
        :param start_date: Data inicial do cálculo.
        :return: DataFrame com a coluna de horas esperadas adicionada.
        """
        now = datetime.now()
        weeks = (now - start_date).days // 7
        expected_hours = (weeks - 1) * 6  # Subtrai um devido ao periodo de férias em janeiro/25
        data['Horas Esperadas'] = expected_hours
        return data

    def export_to_pdf(self, data, file_path, title="Relatório de Apontamentos Quimlabor Jr."):
        """
        Exporta o DataFrame para um arquivo PDF, formatando os dados em forma de tabela.
        :param data: DataFrame com as informações formatadas.
        :param file_path: Caminho para salvar o PDF.
        :param title: Título do relatório.
        """
        with PdfPages(file_path) as pdf:
            fig, ax = plt.subplots(figsize=(8.5, 11))
            ax.axis('tight')
            ax.axis('off')
            fig.suptitle(title, fontsize=16, weight='bold', ha='center', va='top', y=0.95)
            table = ax.table(
                cellText=data.values, 
                colLabels=data.columns, 
                cellLoc='center', 
                loc='upper center'
            )
            table.auto_set_font_size(False)
            table.set_fontsize(12)
            table.auto_set_column_width(col=list(range(len(data.columns))))
            for i, key in enumerate(table.get_celld().keys()):
                cell = table.get_celld()[key]
                cell.set_height(0.05)
            ax.set_position([0.05, 0.2, 0.9, 0.7])
            now = datetime.now().strftime("%H:%M:%S do dia %d/%m/%Y")
            fig.text(0.75, 0.25, f"Relatório criado às {now}", ha='center', va='bottom', fontsize=10)
            pdf.savefig(fig)
            plt.close()

if __name__ == "__main__":
    """
    Bloco principal de execução do código.
    """
    data_manager = QLApontamentos()
    file_path = "data/210225.xlsx"
    pdf_path = "relatorio.pdf"
    try:
        df = data_manager.get_data(file_path)
        df = data_manager.clean_data(df)
        start_date = pd.Timestamp("2025-01-01")
        end_date = pd.Timestamp("2025-06-30")
        total_hours = data_manager.sum_hours(df, start_date, end_date)
        total_hours = data_manager.format_hours(total_hours)
        total_hours = data_manager.add_expected_hours(total_hours, start_date)
        print(df)
        print(total_hours)
        data_manager.export_to_pdf(total_hours, pdf_path)
        print(f"Relatório exportado para '{pdf_path}'.")
    except Exception as e:
        print(f"Erro: {e}")
