import pandas as pd
import time

inicio = time.time()

# Rutas
ruta_ataque_curso = r'C:\Users\a33300\OneDrive - BNP Paribas\1 REMARKETING\BI\2_ATAQUE\MISCELLANEUS\ob505_2025.csv'
ruta_ataque_antes = r'C:\Users\a33300\OneDrive - BNP Paribas\1 REMARKETING\BI\2_ATAQUE\MISCELLANEUS\ANTERIORES\ob505_2024.csv'

archivo_completo = "facturas_ordenadas.csv"
archivo_ultimas = "facturas_ultimas.csv"
archivo_huerfanos = "facturas_abonos_huerfanos.csv"

columnas_fechas = ['INVOICE_DATE', 'DATE_REGISTERED', 'SALE_PROCEEDS_DATE', 'FECHA_PAGO']

# Carga de datos
df_curso = pd.read_csv(ruta_ataque_curso, sep=';', encoding='utf-8', parse_dates=columnas_fechas, decimal=',')
df_anteriores = pd.read_csv(ruta_ataque_antes, sep=';', encoding='utf-8', parse_dates=columnas_fechas, decimal=',')

df_facturas = pd.concat([df_curso, df_anteriores])
df = df_facturas.copy()

# Función para ordenar y emparejar pagos con abonos
def ordenar_matricula(grupo):
    grupo = grupo.sort_values(by='INVOICE_DATE').reset_index(drop=True)
    usados = set()
    resultado = []
    huerfanos = []

    for i, row in grupo.iterrows():
        if i in usados:
            continue

        # Detectar factura de abono por 'A' en INVOICE_NUMBER
        es_abono = isinstance(row['INVOICE_NUMBER'], str) and 'A' in row['INVOICE_NUMBER']

        if not es_abono and row['ITEM_VALUE'] > 0:
            resultado.append(row)
            usados.add(i)

            # Buscar abono correspondiente
            mask = (
                (grupo['REGISTRATION'] == row['REGISTRATION']) &
                (grupo['COMPANY_NAME'] == row['COMPANY_NAME']) &
                (grupo['ITEM_VALUE'] == -row['ITEM_VALUE']) &
                (grupo['TAX_AMOUNT'] == -row['TAX_AMOUNT']) &
                (grupo['INVOICE_DATE'] > row['INVOICE_DATE']) &
                (grupo.index != i)
            )

            posibles = grupo[mask]
            for j, abono in posibles.iterrows():
                if j not in usados and isinstance(abono['INVOICE_NUMBER'], str) and 'A' in abono['INVOICE_NUMBER']:
                    resultado.append(abono)
                    usados.add(j)
                    break  # Solo emparejar el primer abono válido

        elif es_abono and row['ITEM_VALUE'] < 0:
            # Abono sin factura previa
            huerfanos.append(row)
            resultado.append(row)
            usados.add(i)

    return pd.DataFrame(resultado), pd.DataFrame(huerfanos)

# Procesar por matrícula
ordenados = []
huerfanos = []

for reg, grupo in df.groupby("REGISTRATION"):
    ordenado, hu = ordenar_matricula(grupo)
    ordenados.append(ordenado)
    if not hu.empty:
        huerfanos.append(hu)

df_ordenado = pd.concat(ordenados).reset_index(drop=True)
df_huerfanos = pd.concat(huerfanos).reset_index(drop=True) if huerfanos else pd.DataFrame()

# Guardar resultados
df_ordenado.to_csv(archivo_completo, index=False, sep=';')
df_ordenado.groupby("REGISTRATION", as_index=False).tail(1).to_csv(archivo_ultimas, index=False, sep=';')
if not df_huerfanos.empty:
    df_huerfanos.to_csv(archivo_huerfanos, index=False, sep=';')

# Mensajes finales
print("✅ Archivo completo guardado en:", archivo_completo)
print("✅ Archivo últimas facturas guardado en:", archivo_ultimas)
if not df_huerfanos.empty:
    print("⚠️ Archivo de abonos huérfanos guardado en:", archivo_huerfanos)

fin = time.time()
print(f"⏱️ Tiempo de ejecución: {fin - inicio:.2f} segundos")
