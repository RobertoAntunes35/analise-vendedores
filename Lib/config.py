import os 
from datetime import datetime 

import pandas as pd

# Importação de Arquivos
abs_path= os.path.abspath(__file__)

father_path = os.path.dirname(os.path.dirname(abs_path))

folder_files = "static-files/"

usually_path = os.path.join(father_path, folder_files)

# Configurações
semana_dias = {
        0: 'segunda',
        1: 'terca',
        2: 'quarta',
        3: 'quinta',
        4: 'sexta'
    }

# Rename Columns
columns = {
    'pedidos': {
        'Texto56':'razao_social',
        'Texto14':'nome_fantasia',
        'CODCLI':'codigo',
        'Combinação22':'codigo_vendedor',
        'VALPED':'valor_pedido',
        'Data_Importacao':'data_importacao',
        'Combinação22':'codigo_vendedor',
        'Texto36': 'natureza_opereracao'
        },
    'cliente': {
        'D01_Cod_Cliente':'codigo',
        'D01_Nome':'razao_social',
        'Fantasia':'nome_fantasia',
        'xregiao':'dia_semana',
        'D01_Vendedor':'nome_vendedor',
    },
    'colaboradores': {
        'D03_Salao':'codigo',
        'D03_Descricao':'nome',
        'Cargo':'funcao'
    }
}




