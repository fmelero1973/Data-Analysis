import pandas as pd
import numpy as np
from unidecode import unidecode
import datetime
from pathlib import Path
from comunidad_python.obiee.arval_obiee import descargar_obiee_data, obtener_filtros_obiee_data
from melvive.funciones import imp_mensaje_inicial, imp_mensaje_final, imp_mensaje_ko, mensaje_imprimir
from melvive.funciones import ejecutar_con_log, alerta_usuario, alerta_ok_usuario

# FUNCIONES AUXILIARES ---------------------------------------------------------
def añadir_motivo(df: pd.DataFrame, columna: str, condicion: pd.Series, motivo: str) -> None:
    """Añade un motivo de exclusión a las filas que cumplen una condición."""
    df.loc[condicion, columna] = df.loc[condicion, columna].apply(lambda x: x + [motivo])

# VARIABLES --------------------------------------------------------------------
nombre_script = 'GENERAR CARGADOR RELEASE'
ob_des = 'ob713'

ruta_stock = Path(r'C:\Users\a33300\OneDrive - BNP Paribas\1 REMARKETING\BI\1_DESCARGA ARCHIVOS\RELEASE\STOCK_B2C\STOCK.xlsx')
extension_parcial = datetime.datetime.now().strftime("%Y%m%d_CARGADOR_RELEASE")
ruta_descarga = Path(r'C:\Users\a33300\OneDrive - BNP Paribas\1 REMARKETING\BI\2_ATAQUE\RELEASE')

nombre_archivo_csv = ruta_descarga / f'{extension_parcial}.csv'
nombre_archivo_excluidos = ruta_descarga / f'{extension_parcial}_EXCLUIDOS.csv'

# RUN --------------------------------------------------------------------------
print(f'\nSE INICIA SCRIPT {nombre_script}\n')

def mi_funcion_main():
    try:
        # Descarga OBIEE
        df_filtros_obiee = obtener_filtros_obiee_data(ob_des)
        report_params = [
            {'name': 'VP_START_DATE', 'value': '01/01/2020'},
            {'name': 'VP_END_DATE', 'value': '31/12/2030'}
        ]
        df_ob_713 = descargar_obiee_data(ob_des, report_params=report_params)

        # Filtrado OBIEE
        df_ob_713_filtrado = df_ob_713[
            (df_ob_713['LEASE_STATUS'] != 'ACTIVE') &
            (df_ob_713['QUALIFYING_VEHICLE'] == 'N') &
            (df_ob_713['CUSTOMER_QUALIFYING_VEHICLE'] == 'Y')
        ]

        # Cargar stock original
        df_stock_original = pd.read_excel(ruta_stock)
        df_stock_original.columns = [unidecode(col).upper() for col in df_stock_original.columns]
        df_stock_original['MOTIVOS_EXCLUSION'] = [[] for _ in range(len(df_stock_original))]

        # Exclusiones por DMP
        añadir_motivo(df_stock_original, 'MOTIVOS_EXCLUSION', df_stock_original['DMP'] == 0, 'DMP igual a 0')
        añadir_motivo(df_stock_original, 'MOTIVOS_EXCLUSION', df_stock_original['DMP'].isna(), 'DMP nulo')

        # Stock válido para cruce
        df_stock_valido = df_stock_original[
            (df_stock_original['DMP'] != 0) & (df_stock_original['DMP'].notna())
        ].copy()
        df_stock_valido['DMP_NOIVA_GASTOS'] = np.ceil(df_stock_valido['DMP'] / 1.21 + 1000).astype(int)
        df_stock_valido = df_stock_valido[['MATRICULA', 'DMP_NOIVA_GASTOS', 'COLOR DE PINTURA']]

        # Exclusiones por condiciones OBIEE
        matriculas_obiee = set(df_ob_713['REGISTRATION'])

        añadir_motivo(df_stock_original, 'MOTIVOS_EXCLUSION',
                      df_stock_original['MATRICULA'].isin(df_stock_valido.loc[~df_stock_valido['MATRICULA'].isin(matriculas_obiee), 'MATRICULA']),
                      'No está en OBIEE')

        añadir_motivo(df_stock_original, 'MOTIVOS_EXCLUSION',
                      df_stock_original['MATRICULA'].isin(df_ob_713.loc[df_ob_713['LEASE_STATUS'] == 'ACTIVE', 'REGISTRATION']),
                      'LEASE_STATUS = ACTIVE')

        añadir_motivo(df_stock_original, 'MOTIVOS_EXCLUSION',
                      df_stock_original['MATRICULA'].isin(df_ob_713.loc[df_ob_713['QUALIFYING_VEHICLE'] != 'N', 'REGISTRATION']),
                      'QUALIFYING_VEHICLE ≠ N')

        añadir_motivo(df_stock_original, 'MOTIVOS_EXCLUSION',
                      df_stock_original['MATRICULA'].isin(df_ob_713.loc[df_ob_713['CUSTOMER_QUALIFYING_VEHICLE'] != 'Y', 'REGISTRATION']),
                      'CUSTOMER_QUALIFYING_VEHICLE ≠ Y')

        # Generar DataFrame de excluidos
        df_excluidos_stock = df_stock_original[df_stock_original['MOTIVOS_EXCLUSION'].apply(lambda x: len(x) > 0)].copy()
        df_excluidos_stock['MOTIVOS_EXCLUSION'] = df_excluidos_stock['MOTIVOS_EXCLUSION'].apply(lambda x: ', '.join(sorted(set(x))))

        # Cruce final
        df_cruce = pd.merge(
            df_ob_713_filtrado,
            df_stock_valido,
            left_on='REGISTRATION', right_on='MATRICULA',
            how='inner'
        )

        df_cruce['MTR'] = 'N'
        df_cruce['QUOTE'] = None
        df_cruce['INT'] = None
        df_cruce['INTERNAL_COLOUR'] = df_cruce['INTERNAL_COLOUR'].fillna('Serie').str.title().replace('.', 'Serie')

        df_cruce = df_cruce[['REGISTRATION', 'DMP_NOIVA_GASTOS', 'CHASSIS_NUMBER',
                             'COLOR DE PINTURA', 'INTERNAL_COLOUR', 'CO2_EMISSIONS', 'MTR', 'QUOTE', 'INT']]

        # Guardar archivos
        df_cruce.to_csv(nombre_archivo_csv, sep=';', encoding='latin1', header=False, index=False)
        df_excluidos_stock.to_csv(nombre_archivo_excluidos, sep=';', encoding='latin1', index=False)

        # Resumen
        num_incluidos = len(df_cruce)
        num_excluidos = len(df_excluidos_stock)
        resumen_filas = {
            'archivo_cargador': str(nombre_archivo_csv),
            'archivo_excluidos': str(nombre_archivo_excluidos),
            'filas_incluidas': num_incluidos,
            'filas_excluidas': num_excluidos
        }

        mensaje_imprimir(
            f"{nombre_script} OK\n\n"
            f"Cargador: {resumen_filas['archivo_cargador']} ({resumen_filas['filas_incluidas']} filas)\n"
            f"Excluidos: {resumen_filas['archivo_excluidos']} ({resumen_filas['filas_excluidas']} filas)\n"
        )

        return resumen_filas

    except Exception as e:
        raise e

if __name__ == '__main__':
    imp_mensaje_inicial(nombre_script)
    ok = ejecutar_con_log(mi_funcion_main)
    if ok:
        imp_mensaje_final(nombre_script)
        alerta_ok_usuario(nombre_script)
    else:
        imp_mensaje_ko(nombre_script)
        alerta_usuario(nombre_script)
