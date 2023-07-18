import os 
import sys 

import pandas as pd
import numpy as np 

from config import usually_path

name_file_cliente = "D01_Cliente.xls"
name_file_pedidos = "Pedidos.xls"
name_file_colaboradores = "D20_Vendedor.xls"


if os.path.exists(usually_path + name_file_cliente) and os.path.exists(usually_path + name_file_pedidos) and os.path.exists(usually_path + name_file_colaboradores):
    try:
        file_cliente = pd.read_excel(usually_path + name_file_cliente)
        file_pedidos = pd.read_excel(usually_path + name_file_pedidos)
        file_colaboradores = pd.read_excel(usually_path + name_file_colaboradores)

        print("Importação realizada com sucesso.")

    except ValueError as err:
        sys.exit(err)

else:
    print("Erro ao procurar Arquivos.")











