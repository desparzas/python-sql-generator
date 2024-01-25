import os
import pandas as pd


def main():
    # read the excel file, and select the sheet
    df = pd.read_excel('./data/1442 - (240122).CO - Cargar Transacciones 2024.xlsx', sheet_name='Cargar Transacciones')
    # hacer un subdataframe con solo las 5 primeras filas
    columns_type_mapping = {
        'channel': 'VARCHAR',
        'fecha_emision': 'DATE',
        'tx_status': 'VARCHAR',
        'producto_original': 'VARCHAR',
        'airline': 'VARCHAR',
        # 'kam': 'VARCHAR',
        'idagencia': 'VARCHAR',
        'hh_provider': 'VARCHAR',
        'pais_destino': 'VARCHAR',
        'destino': 'VARCHAR',
        'tx_code': 'VARCHAR',
        'gb': 'NUMERIC',
        'hotel': 'VARCHAR'
    }
    # imprimir las filas del dataframe
    insert_queries = []
    for _, row in df.iterrows():
        values = []
        for col, datatype in columns_type_mapping.items():
            # si es varchar y no es nulo, entonces poner comillas simples

            if pd.notnull(row[col]) and datatype == 'VARCHAR' and row[col] != 'NULL' and row[col] != 'nan':
                if type(row[col]) == str:
                    if "'" in row[col]:
                        s = row[col].replace("'", "''")
                        values.append(f"'{s}'")
                    else:
                        values.append(f"'{row[col]}'")
                else:
                    values.append(f"'{row[col]}'")
                    
            elif pd.notnull(row[col]) and datatype == 'DATE' and row[col] != 'NULL' and row[col] != 'nan':
                values.append(f"'{row[col].strftime('%Y-%m-%d')}'")
            elif pd.notnull(row[col]) and datatype == 'NUMERIC' and row[col] != 'NULL' and row[col] != 'nan':
                # limitar a 2 decimales
                values.append(f"{row[col]}")
            else:
                values.append('NULL')
        values_str = ', '.join(values)
        columns_str = ', '.join(columns_type_mapping.keys())

        insert_query = f"INSERT INTO despegar.transaction ({columns_str}) VALUES ({values_str});"
        insert_queries.append(insert_query)

    
    
    # guardar las consultas en un archivo, encoding UTF-8
    with open('./scripts_sql/insert_transaction_2024.sql', 'w', encoding='utf-8') as f:
        f.write('\n'.join(insert_queries))
        


if __name__ == "__main__":
    main()