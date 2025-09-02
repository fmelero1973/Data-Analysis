import pandas as pd

# Archivos
archivo_entrada = "facturas.csv"
archivo_completo = "facturas_ordenadas.csv"
archivo_ultimas = "facturas_ultimas.csv"
archivo_huerfanos = "facturas_abonos_huerfanos.csv"

# Leer archivo
df = pd.read_csv(archivo_entrada)

# Normalizar fecha
df['fecha_factura'] = pd.to_datetime(df['fecha_factura'], dayfirst=True, errors='coerce')

def ordenar_matricula(grupo):
    grupo = grupo.sort_values(by=['fecha_factura']).reset_index(drop=True)
    usados = set()
    resultado = []
    huerfanos = []

    for i, row in grupo.iterrows():
        if i in usados:
            continue

        if row['importe'] > 0:  # factura de pago
            resultado.append(row)
            # Buscar abono correspondiente
            mask = (
                (grupo['registration'] == row['registration']) &
                (grupo['cliente'] == row['cliente']) &
                (grupo['importe'] == -row['importe']) &
                (grupo['fecha_factura'] >= row['fecha_factura'])
            )
            posibles = grupo[mask]
            if not posibles.empty:
                abono = posibles.iloc[0]
                abono_idx = abono.name
                if abono_idx not in usados:
                    resultado.append(abono)
                    usados.add(abono_idx)

        elif row['importe'] < 0:  # abono sin factura previa en este archivo
            huerfanos.append(row)
            resultado.append(row)

        usados.add(i)

    return pd.DataFrame(resultado), pd.DataFrame(huerfanos)

# Procesar todas las matrículas
ordenados = []
huerfanos = []
for reg, grupo in df.groupby("registration"):
    ordenado, hu = ordenar_matricula(grupo)
    ordenados.append(ordenado)
    if not hu.empty:
        huerfanos.append(hu)

df_ordenado = pd.concat(ordenados).reset_index(drop=True)
df_huerfanos = pd.concat(huerfanos).reset_index(drop=True) if huerfanos else pd.DataFrame()

# Guardar archivos
df_ordenado.to_csv(archivo_completo, index=False)
df_ordenado.groupby("registration", as_index=False).tail(1).to_csv(archivo_ultimas, index=False)
if not df_huerfanos.empty:
    df_huerfanos.to_csv(archivo_huerfanos, index=False)

print("✅ Archivo completo guardado en:", archivo_completo)
print("✅ Archivo últimas facturas guardado en:", archivo_ultimas)
if not df_huerfanos.empty:
    print("⚠️ Archivo de abonos huérfanos guardado en:", archivo_huerfanos)