import pandas as pd

# Leer el archivo XLSX
archivo_xlsx = 'cuentas_credicel_Original.xlsx'
df = pd.read_excel(archivo_xlsx)


def actualizar_valores(row):
    # Iterar a través de las rows del DataFrame
    # for indice, row in df.iterrows():
        # Verificar si AI contiene un número y si riesgo está en blanco
    if pd.notna(row['AI']):
        row['status_cuenta'] = row['riesgo']
        row['puntos'] = row['porc_enganche']
        row['riesgo'] = row['porc_tasa']
        row['porc_enganche'] = row['score_buro']
        row['porc_tasa'] = row['razones_buro']
        row['score_buro'] = row['semana_actual']
        row['razones_buro'] = row['codigo_postal']
        row['semana_actual'] = row['AH']
        row['codigo_postal'] = row['AI']

    # Verificar si AJ contiene un número y si codigo_postal está en blanco
    if pd.notna(row['AJ']):
        row['puntos'] = row['porc_tasa']
        row['riesgo'] = row['score_buro']
        row['porc_enganche'] = row['razones_buro']
        row['porc_tasa'] = row['semana_actual']
        row['score_buro'] = row['codigo_postal']
        row['razones_buro'] = row['AH']
        row['semana_actual'] = row['AI']
        row['codigo_postal'] = row['AJ']
    return row


# Aplicar la función a lo largo de las rows (axis=1) del DataFrame
df = df.apply(actualizar_valores, axis=1)

# Eliminar las columnas "AI" y "AJ"
df = df.drop(['AI', 'AJ'], axis=1)

# Guardar el DataFrame actualizado en un nuevo archivo XLSX
df.to_excel('Credicel_Limpio_Celdas.xlsx', index=False)
