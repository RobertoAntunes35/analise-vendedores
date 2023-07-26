import os 
import sys 
from datetime import datetime
import copy

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from files import file_colaboradores, file_cliente, file_pedidos
from config import columns, father_path


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
    def __init__(self, file_cliente: pd.DataFrame, file_pedidos: pd.DataFrame, file_colaboradores: pd.DataFrame, start_date, end_date, periods) -> None:
        self._file_client = file_cliente
        self._file_pedidos = file_pedidos
        self._file_colaboradores = file_colaboradores

        self._start_date = start_date
        self._end_date = end_date
        self._periods = periods



    @show_error
    def convert_seller(self, value):
        sellers = self.filter_seller()
        if isinstance(value, int):
            for i in (sellers.loc[sellers['codigo'] == value]['nome']):
                return str(i)
        if isinstance(value, str):
            for i in sellers.loc[sellers['nome'] == value]['codigo']:
                return int(i)



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

    def increase_data(self):
        datas = Dates()
        datas_analise = datas.get_weekday_dates(self._start_date, self._end_date, self._periods)
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
                clients_for_day[vendedor][value_data][data] = pd.to_datetime(clients_for_day[vendedor][value_data][data])

            save_directory = "../files-sellers/"
            os.makedirs(save_directory, exist_ok=True)
            name_file_vendedor = vendedor.replace(" ", "_")
            wb.remove(wb.active)
            wb.save(os.path.join(save_directory, f"{name_file_vendedor}.xlsx"))               

    @show_error    
    def return_sellers_to_folder(self, path):
        pass  

    @show_error

    @show_error
    def file_orders(self):
        data_orders = copy.copy(self._file_pedidos)
        client_for_day, days = self.client_for_seller_for_day()
        datas = Dates()
        datas_semana = datas.get_weekday_dates(self._start_date, self._end_date, self._periods)


        path_sellers_full = os.path.join(father_path, "files-sellers")

        # Loop para pegar os vendedores
        for seller in os.listdir(path_sellers_full):
            name_seller = seller.split('.')[0].replace("_", " ")
            
            value_analise_vendedor = self.convert_seller(name_seller)
            
            frame_orders_seller = data_orders.loc[data_orders['codigo_vendedor'] == value_analise_vendedor]

            # Loop para pegar os dias da semana
            for valor, dia in days.items():
                
                frame_data_sellers = pd.read_excel(f'{path_sellers_full}/{seller}', sheet_name=dia)
                
                nomes_fantasia_em_cadastro = [str(value) for value in frame_data_sellers['nome_fantasia']]
                nomes_fantasia_positivados = [str(value) for value in frame_orders_seller['nome_fantasia']]
                
                # Loop para verificar a positivação do cliente 
                for nome_fantasia_analise in nomes_fantasia_em_cadastro:    
                    if nome_fantasia_analise in nomes_fantasia_positivados:
                        
                        # Loop para verificar todas as segundas, tercas, quartas, quintas e sexta do mes
                        for key_week_day, value_week_day in datas_semana.items():
                            print(type(value_week_day))
                            frame_do_cliente_positivado = frame_orders_seller.loc[(frame_orders_seller['nome_fantasia'] == nome_fantasia_analise) & (frame_orders_seller['natureza_opereracao'] != 3), ['nome_fantasia', 'valor_pedido', 'data_importacao', 'natureza_opereracao']]
                            

if __name__ == "__main__":
    arquivo_clientes = Excel(file_cliente)
    new_file_cliente = arquivo_clientes.rename_columns(new_columns=columns['cliente'])
    
    arquivo_pedidos = Excel(file_pedidos)
    new_file_pedido = arquivo_pedidos.rename_columns(new_columns=columns['pedidos'])

    arquivo_colaboradores = Excel(file_colaboradores)
    new_file_colaborador = arquivo_colaboradores.rename_columns(new_columns=columns['colaboradores'])


    relatorio_analise = Relatorio(file_cliente=new_file_cliente, file_pedidos=new_file_pedido, file_colaboradores=new_file_colaborador, start_date="01/07/2023", end_date="31/07/2023", periods=31)
    relatorio_analise.filter_seller()
    relatorio_analise.clientForSellers()
    relatorio_analise.client_for_seller_for_day()
    relatorio_analise.increase_data()
    relatorio_analise.file_orders()


