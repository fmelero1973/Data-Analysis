# --- DETECCIÓN DE FILAS EXCLUIDAS DEL STOCK ----------------------------------

# Cargar archivo original sin filtrar
df_stock_original = pd.read_excel(ruta_stock)

# Normalizar columnas como en el flujo principal
df_stock_original.columns = [unidecode(col).upper() for col in df_stock_original.columns]

# Crear columna de motivos de exclusión como lista
df_stock_original['MOTIVOS_EXCLUSION'] = [[] for _ in range(len(df_stock_original))]

# Evaluar exclusión por DMP = 0
mask_dmp_cero = df_stock_original['DMP'] == 0
df_stock_original.loc[mask_dmp_cero, 'MOTIVOS_EXCLUSION'] = df_stock_original.loc[mask_dmp_cero, 'MOTIVOS_EXCLUSION'].apply(lambda x: x + ['DMP igual a 0'])

# Evaluar exclusión por DMP nulo
mask_dmp_nulo = df_stock_original['DMP'].isna()
df_stock_original.loc[mask_dmp_nulo, 'MOTIVOS_EXCLUSION'] = df_stock_original.loc[mask_dmp_nulo, 'MOTIVOS_EXCLUSION'].apply(lambda x: x + ['DMP nulo'])

# Filtrar stock válido como en tu código
df_stock_valido = df_stock_original[
    (~mask_dmp_cero) & (~mask_dmp_nulo)
].copy()

# Normalizar columnas y calcular DMP_NOIVA_GASTOS como en tu flujo
df_stock_valido['DMP_NOIVA_GASTOS'] = np.ceil(df_stock_valido['DMP'] / 1.21 + 1000).astype(int)

# Seleccionar columnas necesarias para el cruce
df_stock_valido = df_stock_valido[['MATRICULA', 'DMP_NOIVA_GASTOS', 'COLOR DE PINTURA']]

# Realizar cruce como en tu código
df_cruce = pd.merge(df_ob_713,
                    df_stock_valido,
                    left_on='REGISTRATION', right_on='MATRICULA',
                    how='inner')

# Detectar matrículas que no aparecen en el cruce
matriculas_cruzadas = set(df_cruce['MATRICULA'])
mask_no_cruce = ~df_stock_valido['MATRICULA'].isin(matriculas_cruzadas)

# Añadir motivo de exclusión por no estar en OBIEE o no cumplir filtros
df_stock_original.loc[df_stock_original['MATRICULA'].isin(df_stock_valido.loc[mask_no_cruce, 'MATRICULA']),
                      'MOTIVOS_EXCLUSION'] = df_stock_original.loc[df_stock_original['MATRICULA'].isin(df_stock_valido.loc[mask_no_cruce, 'MATRICULA']),
                                                                    'MOTIVOS_EXCLUSION'].apply(lambda x: x + ['No cruza con OBIEE'])

# Filtrar solo excluidos
df_excluidos_stock = df_stock_original[df_stock_original['MOTIVOS_EXCLUSION'].apply(lambda x: len(x) > 0)].copy()

# Convertir lista de motivos en texto separado por comas
df_excluidos_stock['MOTIVOS_EXCLUSION'] = df_excluidos_stock['MOTIVOS_EXCLUSION'].apply(lambda x: ', '.join(x))
