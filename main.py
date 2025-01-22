import pandas as pd
import os

class QLApontamentos():
    def __init__(self):
        self.data = None

    # Load the Excel file and return a DataFrame
    def get_data(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"O arquivo '{file_path}' não foi encontrado.")
        
        try:
            self.data = pd.read_excel(file_path)

        except Exception as e:
            raise ValueError(f"Erro ao carregar o arquivo Excel: {e}")
        
        return self.data


    # Filter the DataFrame to keep only the specified columns
    def clean_data(self, data):

        # Filter the DataFrame to keep only the specified columns
        data = data[['Membro - name', 'Categoria', 'Reunião', 'Atribuição Interna', 'Data - start', 'Data - end', 'Duração', 'Detalhamento']]

        # Split 'Membro - name' by semicolons, expand rows, and trim whitespace
        data['Membro - name'] = data['Membro - name'].str.split(r'[;]')
        data = data.explode('Membro - name')
        data['Membro - name'] = data['Membro - name'].str.strip()

        # Convert 'Data - start' and 'Data - end' to datetime
        data['Data - start'] = pd.to_datetime(data['Data - start'], errors='coerce')
        data['Data - end'] = pd.to_datetime(data['Data - end'], errors='coerce')

        return data


    # Sum the hours for each member within the specified date range
    def sum_hours(self, data, start_date, end_date):
        # Filter by the date range
        mask = (data['Data - start'] >= start_date) & (data['Data - start'] <= end_date)
        filtered_data = data[mask]

        # Sum the duration for each member
        total_hours = filtered_data.groupby('Membro - name')['Duração'].sum().reset_index()

        # Sort by total hours in descending order
        total_hours = total_hours.sort_values(by='Duração', ascending=False)

        return total_hours

if __name__ == "__main__":
    data_manager = QLApontamentos()
    file_path = "data/220125.xlsx"
    try:
        df = data_manager.get_data(file_path)
        df = data_manager.clean_data(df)
        
        # Specify the date range (e.g., from 2025-01-01 to 2025-01-31)
        start_date = pd.Timestamp("2025-01-01")
        end_date = pd.Timestamp("2025-01-31")

        total_hours = data_manager.sum_hours(df, start_date, end_date)

        # Print the result
        print(df)
        print(total_hours)
    except Exception as e:
        print(f"Erro: {e}")
    