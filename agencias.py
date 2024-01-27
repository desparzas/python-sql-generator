import pandas as pd


def main():
    # read the excel file, and select the sheet
    df = pd.read_excel('./data/1442 - (240122).CO - Cargar agencias.xlsx', sheet_name='Cargar Agencias')
    # hacer un subdataframe con solo las 5 primeras filas
    columns_type_mapping = {
        'environment': 'varchar', 
        'country': 'varchar', 
        'id': 'varchar', 
        'agency_name': 'varchar', 
        'agency_status': 'varchar', 
        'segmento': 'varchar', 
        'fecha_activacion': 'date', 
        'tax_id_country': 'varchar', 
        'tax_id_type': 'int', 
        'tax_id': 'varchar', 
        'phone_country': 'varchar', 
        'phone_number': 'varchar', 
        'phone_label': 'varchar', 
        'email': 'varchar', 
        'email_label': 'varchar', 
        'kam': 'varchar'
    }
    
    columns_to_insert = [
        'environment',
        'country',
        'id',
        'agency_name',
        'agency_status',
        'segmento',
        'fecha_activacion'
    ]


    # filtrar las columnas que se van a insertar
    columns_type_mapping = {k: v for k, v in columns_type_mapping.items() if k in columns_to_insert}
    print(columns_type_mapping)
    print(df.columns)
    print(df.head())
    insert_queries = []
    for _, row in df.iterrows():
        values = []
        for col, datatype in columns_type_mapping.items():
            # si es varchar y no es nulo, entonces poner comillas simples
            if pd.notnull(row[col]) and datatype.upper() == 'VARCHAR' and row[col] != 'NULL' and row[col] != 'nan':
                # verificar que no tenga comillas simples dentro de la cadena
                
                if type(row[col]) == str:
                    if "'" in row[col]:
                        s = row[col].replace("'", "''")
                        values.append(f"'{s}'")
                    else:
                        values.append(f"'{row[col]}'")
                else:
                    values.append(f"'{row[col]}'")
            elif pd.notnull(row[col]) and datatype.upper() == 'DATE' and row[col] != 'NULL' and row[col] != 'nan':
                values.append(f"'{row[col].strftime('%Y-%m-%d')}'")
            elif pd.notnull(row[col]) and datatype.upper() == 'NUMERIC' and row[col] != 'NULL' and row[col] != 'nan':
                # limitar a 2 decimales
                values.append(f"{row[col]}")
            else:
                values.append('NULL')
        values_str = ', '.join(values)
        columns_str = ', '.join(columns_type_mapping.keys())

        insert_query = f"INSERT INTO despegar.agencia ({columns_str}) VALUES ({values_str});"
        insert_queries.append(insert_query)

    
    
    # guardar las consultas en un archivo, encoding UTF-8
    with open('./scripts_sql/insert_agencias_240122_colombia.sql', 'w', encoding='utf-8') as f:
        f.write('\n'.join(insert_queries))
        


if __name__ == "__main__":
    main()