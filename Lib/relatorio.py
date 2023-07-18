import os 
import sys 
from datetime import datetime
import copy

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from files import file_colaboradores, file_cliente, file_pedidos
from config import columns

def show_error(func):
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                print(f"Ocorreu um erro na função '{func.__name__}'")
                print(f"Tipo do erro: {type(e).__name__}")
                print(f"Mensagem de erro: {str(e)}")
        return wrapper



class Excel:
    def __init__(self, file) -> None:
        self._file = file 

    @show_error
    def rename_columns(self, new_columns: dict):
        self.new_file = self._file.rename(columns=new_columns)
        return self.new_file


class Dates:
    def __init__(self) -> None:
        pass

    @show_error
    def get_weekday_dates(self, start_date: datetime, end_date: datetime, periods: int) -> dict:
        semana_dias = {
            0: 'segunda',
            1: 'terca',
            2: 'quarta',
            3: 'quinta',
            4: 'sexta'
        }

        start_date = datetime.strptime(start_date, "%d/%m/%Y").date()
        end_date = datetime.strptime(end_date, "%d/%m/%Y").date()
        dates = pd.date_range(start=start_date, end=end_date, periods=periods)
        
        dias_semana = {value: [] for value in semana_dias.values()}

        for data in dates:
            dia_semana = semana_dias.get(data.weekday(), [])
            if dia_semana != []:
                dias_semana[dia_semana].append(data.date())
        return dias_semana
    
  
class Relatorio:
    def __init__(self, file_cliente: pd.DataFrame, file_pedidos: pd.DataFrame, file_colaboradores: pd.DataFrame) -> None:
        self._file_client = file_cliente
        self._file_pedidos = file_pedidos
        self._file_colaboradores = file_colaboradores

    @show_error
    def filter_seller(self) -> pd.DataFrame:
        seller_data = copy.copy(self._file_colaboradores)

        file_vendedor = seller_data.filter(['codigo', 'nome']).loc[seller_data['funcao'] == 'VENDEDOR EXTERNO']
        return file_vendedor.sort_values('codigo').reset_index(drop=True)

    @show_error
    def clientForSellers(self) -> dict:
        sellers = self.filter_seller()
        client_data = copy.copy(self._file_client)
        
        frames_vendedores = {value: [] for value in sellers['nome']}

        for vendedor in sellers['nome']:
            frames_vendedores[vendedor] = client_data.filter(['codigo', 'nome_fantasia', 'dia_semana', 'nome_vendedor']).loc[client_data['nome_vendedor'] == vendedor].sort_values('dia_semana').reset_index(drop=True)

        return frames_vendedores

    @show_error
    def client_for_seller_for_day(self) -> dict:
        clients = self.clientForSellers()
        sellers = self.filter_seller()

        dias_semana = {
            2: 'segunda',
            3: 'terca',
            4: 'quarta',
            5: 'quinta',
            6: 'sexta',
        }

        seller_client_day = {
            vendedor: {dia: [] for dia in dias_semana.values()}
            for vendedor in sellers['nome']
        }

        for vendedor in sellers['nome']:
            for key, value in dias_semana.items():
                seller_client_day[vendedor][value] = clients[vendedor].loc[clients[vendedor]['dia_semana'] == key]
        return seller_client_day, dias_semana

    def increase_data(self, start_date, end_date, periods):
        datas = Dates()
        datas_analise = datas.get_weekday_dates(start_date, end_date, periods)
        sellers = self.filter_seller()
        clients_for_day, dias = self.client_for_seller_for_day()

        for vendedor in sellers['nome']:
            wb = Workbook()
            for key_data, value_data in dias.items():
                clients_for_day[vendedor][value_data]     
                planilha = wb.create_sheet(title=value_data)
                for data in datas_analise[value_data]:
                    clients_for_day[vendedor][value_data][data] = np.nan
                for row in dataframe_to_rows(clients_for_day[vendedor][value_data], index=False, header=True):
                    planilha.append(row)

            save_directory = "../files-sellers/"
            os.makedirs(save_directory, exist_ok=True)
            wb.save(os.path.join(save_directory, f"{vendedor}.xlsx"))               

        

    def file_orders(self, ):
        data_orders = copy.copy(self._file_pedidos)
        client_for_day = self.client_for_seller_for_day()
        




if __name__ == "__main__":
    arquivo_clientes = Excel(file_cliente)
    new_file_cliente = arquivo_clientes.rename_columns(new_columns=columns['cliente'])
    
    arquivo_pedidos = Excel(file_pedidos)
    new_file_pedido = arquivo_pedidos.rename_columns(new_columns=columns['pedidos'])

    arquivo_colaboradores = Excel(file_colaboradores)
    new_file_colaborador = arquivo_colaboradores.rename_columns(new_columns=columns['colaboradores'])

    datas = Dates()
    datas_analise = datas.get_weekday_dates(start_date="01/07/2023", end_date="31/07/2023", periods=31)

    relatorio_analise = Relatorio(file_cliente=new_file_cliente, file_pedidos=new_file_pedido, file_colaboradores=new_file_colaborador)
    relatorio_analise.filter_seller()
    relatorio_analise.clientForSellers()
    relatorio_analise.client_for_seller_for_day()
    relatorio_analise.file_orders()
    relatorio_analise.increase_data(start_date="01/07/2023", end_date="31/07/2023", periods=31)

    

