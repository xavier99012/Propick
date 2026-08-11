
const PY_CODE = String.raw`
import pandas as pd
import math
import csv
import json

def ejecutar_motor():
    resultado = {"success": False, "error": None, "materiales_faltantes": [], "resumen": [], "detalle": [], "total_pallets": 0, "total_cajas": 0, "total_auxiliares": 0}
    try:
        comprimido = pd.read_excel('/tmp/propick.xlsx')
        comprimido['Material'] = comprimido['Material'].astype(str).str.strip()
    except Exception as e:
        resultado["error"] = f"No se pudo leer el archivo Propick: {e}"
        return json.dumps(resultado)

    try:
        master = None
        last_err = None
        for sep in [',', ';', '\t', '|']:
            try:
                candidate = pd.read_csv('/tmp/master.csv', sep=sep, engine='python', on_bad_lines='skip', dtype={'Codigo': str})
                if 'Codigo' in candidate.columns:
                    master = candidate
                    break
            except Exception as e:
                last_err = e
        if master is None:
            raise last_err if last_err else Exception("No se detectó el delimitador correcto (se esperaba la columna 'Codigo')")
        master['Codigo'] = master['Codigo'].astype(str).str.strip()
    except Exception as e:
        resultado["error"] = f"No se pudo leer MasterData: {e}"
        return json.dumps(resultado)

    materiales = comprimido['Material'].unique()
    materiales_master = master['Codigo'].unique()
    materiales_faltantes = [mat for mat in materiales if mat not in materiales_master]

    if materiales_faltantes:
        resultado["error"] = "Hay materiales que no están en MasterData"
        resultado["materiales_faltantes"] = [str(m) for m in materiales_faltantes]
        return json.dumps(resultado)

    agrupadoC = comprimido.groupby(['Material', 'Ruta'], as_index=False).agg({'Cantidad en UMA': 'sum'})
    otras_columnasC = comprimido.drop(columns=['Cantidad en UMA']).drop_duplicates(subset=['Material', 'Ruta'])
    resultadoC = pd.merge(agrupadoC, otras_columnasC, on=['Material', 'Ruta'], how='left')
    resultadoC = resultadoC.sort_values(by='Ruta').reset_index(drop=True)

    resultadoC['UxC'] = resultadoC['Material'].map(master.set_index('Codigo')['UnitBox'])
    resultadoC['CxP'] = resultadoC['Material'].map(master.set_index('Codigo')['BoxPallet'])
    resultadoC['CxZ'] = resultadoC['Material'].map(master.set_index('Codigo')['ZoneCode'])
    resultadoC['Zona'] = resultadoC['Material'].map(master.set_index('Codigo')['NameZone'])
    resultadoC['Volumen'] = resultadoC['Material'].map(master.set_index('Codigo')['Volumen'])
    resultadoC['Grupo'] = resultadoC['Material'].map(master.set_index('Codigo')['Grupo'])

    resultadoC = resultadoC[resultadoC['CxZ'] != "NO"].copy()
    resultadoC[['Cantidad en UMA', 'UxC', 'CxP']] = resultadoC[['Cantidad en UMA', 'UxC', 'CxP']].apply(pd.to_numeric)

    def calcular_valores(row):
        cantidad = row['Cantidad en UMA']
        uxc = row['UxC']
        cxp = row['CxP']
        if cantidad >= uxc * cxp:
            pallets = math.floor(cantidad / (uxc * cxp))
        else:
            pallets = 0
        if cantidad >= uxc:
            cajas = math.floor((cantidad - (math.floor(cantidad / (uxc * cxp)) * uxc * cxp)) / uxc)
        else:
            cajas = 0
        unidades = cantidad % uxc
        cajas_totales = round(cantidad / uxc, 2)
        if cantidad >= uxc:
            cajas_adb = cajas + (unidades / uxc)
        else:
            cajas_adb = cantidad / uxc
        pallets_dist = round(cajas_adb / cxp, 4)
        return pd.Series([pallets, cajas, unidades, cajas_totales, cajas_adb, pallets_dist],
                         index=['Pallets', 'Cajas', 'Unidades', 'CajasTotales', 'CajasAdB', 'PalletsDist'])

    resultadoC[['Pallets', 'Cajas', 'Unidades', 'CajasTotales', 'CajasAdB', 'PalletsDist']] = resultadoC.apply(calcular_valores, axis=1)

    df = resultadoC
    skus_df = resultadoC

    try:
        with open('/tmp/personal.csv', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            auxiliares = [{"Nombre": row["NombreAuxiliar"]} for row in reader]
    except Exception as e:
        resultado["error"] = f"No se pudo leer Asignación de Personal: {e}"
        return json.dumps(resultado)

    num_auxiliares = len(auxiliares)
    if num_auxiliares == 0:
        resultado["error"] = "No se definieron auxiliares"
        return json.dumps(resultado)

    carga_total = df["CajasAdB"].sum()
    carga_ideal = carga_total / num_auxiliares
    margen = 0.06
    carga_maxima = carga_ideal * (1 + margen)

    asignaciones = {aux["Nombre"]: {"Rutas": [], "Cajas": 0, "Zonas": set()} for aux in auxiliares}

    df['Zona'] = df['Zona'].astype(str)
    df_rutas_agrupadas = df.groupby(['Ruta', 'Zona'], as_index=False)['CajasAdB'].sum()

    zonas_ordenadas = df_rutas_agrupadas.groupby("Zona")["CajasAdB"].sum().sort_values(ascending=False).index.tolist()
    rutas_por_zona = {
        zona: df_rutas_agrupadas[df_rutas_agrupadas["Zona"] == zona].sort_values(by="CajasAdB", ascending=False).to_dict("records")
        for zona in zonas_ordenadas
    }

    aux_index = 0
    for zona in zonas_ordenadas:
        rutas_zona = rutas_por_zona[zona]
        while rutas_zona:
            if aux_index >= len(auxiliares) - 1:
                break
            aux = auxiliares[aux_index]
            nombre_aux = aux["Nombre"]
            carga_actual = asignaciones[nombre_aux]["Cajas"]
            asignada = False
            for i, ruta in enumerate(rutas_zona):
                if carga_actual + ruta["CajasAdB"] <= carga_maxima:
                    asignaciones[nombre_aux]["Rutas"].append({"Zona": ruta["Zona"], "Ruta": ruta["Ruta"], "Cajas": ruta["CajasAdB"]})
                    asignaciones[nombre_aux]["Cajas"] += ruta["CajasAdB"]
                    asignaciones[nombre_aux]["Zonas"].add(ruta["Zona"])
                    rutas_zona.pop(i)
                    asignada = True
                    break
                else:
                    break
            if not asignada:
                faltante = carga_maxima - carga_actual
                mejor_ruta = None
                mejor_diff = float('inf')
                for i, ruta in enumerate(rutas_zona):
                    cajas = ruta["CajasAdB"]
                    if cajas <= faltante:
                        diff = faltante - cajas
                        if diff < mejor_diff:
                            mejor_diff = diff
                            mejor_ruta = (i, ruta)
                if mejor_ruta:
                    idx, ruta = mejor_ruta
                    asignaciones[nombre_aux]["Rutas"].append({"Zona": ruta["Zona"], "Ruta": ruta["Ruta"], "Cajas": ruta["CajasAdB"]})
                    asignaciones[nombre_aux]["Cajas"] += ruta["CajasAdB"]
                    asignaciones[nombre_aux]["Zonas"].add(ruta["Zona"])
                    rutas_zona.pop(idx)
                else:
                    aux_index += 1
            rutas_por_zona[zona] = rutas_zona

    if aux_index >= len(auxiliares) - 1:
        aux_index = len(auxiliares) - 1
    aux_final = auxiliares[aux_index]["Nombre"]

    for zona, rutas_restantes in rutas_por_zona.items():
        for r in rutas_restantes:
            asignaciones[aux_final]["Rutas"].append({"Zona": r["Zona"], "Ruta": r["Ruta"], "Cajas": r["CajasAdB"]})
            asignaciones[aux_final]["Cajas"] += r["CajasAdB"]
            asignaciones[aux_final]["Zonas"].add(r["Zona"])

    resultados = []
    for aux in asignaciones:
        for r in asignaciones[aux]["Rutas"]:
            resultados.append({
                "Auxiliar": aux,
                "Zonas": r["Zona"],
                "Rutas": str(r["Ruta"]),
                "CajasAdB": r["Cajas"]
            })

    asignaciones_df = pd.DataFrame(resultados)


    skus_df['Zona'] = skus_df['Zona'].astype(str)
    skus_df['Ruta'] = skus_df['Ruta'].astype(str)
    skus_df['Grupo'] = skus_df['Grupo'].astype(str)
    skus_df['Volumen'] = skus_df['Volumen'].astype(str)
    skus_df['VolumenNum'] = pd.to_numeric(skus_df['Volumen'].str.extract(r'(\d+)', expand=False), errors='coerce')
    asignaciones_df['Rutas'] = asignaciones_df['Rutas'].astype(str)

    merged_df = pd.merge(skus_df, asignaciones_df, left_on=['Zona', 'Ruta'], right_on=['Zonas', 'Rutas'], how='inner')

    zona4_df = merged_df[merged_df['Zona'] == 'Zona 4'].copy()
    otras_zonas_df = merged_df[merged_df['Zona'] != 'Zona 4'].copy()

    def ordenar_skus(df):
        if 'VolumenNum' not in df.columns:
            df['Volumen'] = df['Volumen'].astype(str)
            df['VolumenNum'] = pd.to_numeric(df['Volumen'].str.extract(r'(\d+)', expand=False), errors='coerce')
        df['Prioridad'] = df['Grupo'].apply(lambda x: 20000 if x == 'GRB' else -1 if x == 'PRB' else 0)
        df['VolumenOrden'] = df['Prioridad'] + df['VolumenNum'].fillna(0)
        return df.sort_values(by=['Ruta', 'VolumenOrden'], ascending=[True, False])

    otras_zonas_df = ordenar_skus(otras_zonas_df)
    zona4_df = zona4_df.sort_values(by=['Ruta', 'Zona'])

    full_df = pd.concat([otras_zonas_df, zona4_df], ignore_index=True)
    pallets_montacarga_df = full_df[full_df['Pallets'] > 0].copy()
    pallet_id = 1
    pallets = []
    incompletos = []

    for (aux, ruta, zona), group in full_df.groupby(['Auxiliar', 'Ruta', 'Zona']):
        group = group.reset_index(drop=True)
        prb = group[group['Grupo'] == 'PRB']
        for _, row in prb.iterrows():
            pallets.append({
                'Auxiliar': aux, 'Ruta': ruta, 'Zona': zona,
                'Material': row['Material'], 'Grupo': row['Grupo'], 'Volumen': row['Volumen'],
                'PalletsNum': pallet_id, 'Número de material': row['Número de material'],
                'Transporte': row['Transporte'], 'CajasAdB': row['CajasAdB_x'],
                'Cantidad en UMA': row.get('Cantidad en UMA'), 'PsEx': row.get('PsEx'),
                'UxC': row.get('UxC'), 'CxP': row.get('CxP'), 'CxZ': row.get('CxZ'),
                'Pallets': row.get('Pallets'), 'Cajas': row.get('Cajas'), 'Unidades': row.get('Unidades'),
                'CajasTotales': row.get('CajasTotales'), 'PalletsDist': row.get('PalletsDist')
            })
            pallet_id += 1

        if zona == 'Zona 4':
            for _, row in group[group['Grupo'] != 'PRB'].iterrows():
                pallets.append({
                    'Auxiliar': aux, 'Ruta': ruta, 'Zona': zona,
                    'Material': row['Material'], 'Grupo': row['Grupo'], 'Volumen': row['Volumen'],
                    'PalletsNum': pallet_id, 'Número de material': row['Número de material'],
                    'Transporte': row['Transporte'], 'CajasAdB': row['CajasAdB_x'],
                    'Cantidad en UMA': row.get('Cantidad en UMA'), 'PsEx': row.get('PsEx'),
                    'UxC': row.get('UxC'), 'CxP': row.get('CxP'), 'CxZ': row.get('CxZ'),
                    'Pallets': row.get('Pallets'), 'Cajas': row.get('Cajas'), 'Unidades': row.get('Unidades'),
                    'CajasTotales': row.get('CajasTotales'), 'PalletsDist': row.get('PalletsDist')
                })
            pallet_id += 1
            continue

        restante = group[group['Grupo'] != 'PRB'].copy().reset_index(drop=True)
        actual_pallet = []
        total_pallet = 0

        for _, row in restante.iterrows():
            pal_dist = row['PalletsDist']
            if total_pallet + pal_dist <= 1 + margen:
                actual_pallet.append(row)
                total_pallet += pal_dist
            else:
                if 1 <= total_pallet <= 1 + margen:
                    for r in actual_pallet:
                        pallets.append({
                            'Auxiliar': aux, 'Ruta': ruta, 'Zona': zona,
                            'Material': r['Material'], 'Grupo': r['Grupo'], 'Volumen': r['Volumen'],
                            'PalletsNum': pallet_id, 'Número de material': r['Número de material'],
                            'Transporte': r['Transporte'], 'CajasAdB': r['CajasAdB_x'],
                            'Cantidad en UMA': r.get('Cantidad en UMA'), 'PsEx': r.get('PsEx'),
                            'UxC': r.get('UxC'), 'CxP': r.get('CxP'), 'CxZ': r.get('CxZ'),
                            'Pallets': r.get('Pallets'), 'Cajas': r.get('Cajas'), 'Unidades': r.get('Unidades'),
                            'CajasTotales': r.get('CajasTotales'), 'PalletsDist': r.get('PalletsDist')
                        })
                    pallet_id += 1
                else:
                    for r in actual_pallet:
                        incompletos.append({
                            'Ruta': ruta, 'Zona': zona, 'Material': r['Material'], 'Grupo': r['Grupo'],
                            'Volumen': r['Volumen'], 'PalletsDist': r['PalletsDist'],
                            'Número de material': r['Número de material'], 'Transporte': r['Transporte'],
                            'CajasAdB': r['CajasAdB_x'], 'Cantidad en UMA': r.get('Cantidad en UMA'),
                            'PsEx': r.get('PsEx'), 'UxC': r.get('UxC'), 'CxP': r.get('CxP'), 'CxZ': r.get('CxZ'),
                            'Pallets': r.get('Pallets'), 'Cajas': r.get('Cajas'), 'Unidades': r.get('Unidades'),
                            'CajasTotales': r.get('CajasTotales')
                        })
                actual_pallet = [row]
                total_pallet = pal_dist

        if actual_pallet:
            if 1 <= total_pallet <= 1 + margen:
                for r in actual_pallet:
                    pallets.append({
                        'Auxiliar': aux, 'Ruta': ruta, 'Zona': zona,
                        'Material': r['Material'], 'Grupo': r['Grupo'], 'Volumen': r['Volumen'],
                        'PalletsNum': pallet_id, 'Número de material': r['Número de material'],
                        'Transporte': r['Transporte'], 'CajasAdB': r['CajasAdB_x'],
                        'Cantidad en UMA': r.get('Cantidad en UMA'), 'PsEx': r.get('PsEx'),
                        'UxC': r.get('UxC'), 'CxP': r.get('CxP'), 'CxZ': r.get('CxZ'),
                        'Pallets': r.get('Pallets'), 'Cajas': r.get('Cajas'), 'Unidades': r.get('Unidades'),
                        'CajasTotales': r.get('CajasTotales'), 'PalletsDist': r.get('PalletsDist')
                    })
                pallet_id += 1
            else:
                for r in actual_pallet:
                    incompletos.append({
                        'Ruta': ruta, 'Zona': zona, 'Material': r['Material'], 'Grupo': r['Grupo'],
                        'Volumen': r['Volumen'], 'PalletsDist': r['PalletsDist'],
                        'Número de material': r['Número de material'], 'Transporte': r['Transporte'],
                        'CajasAdB': r['CajasAdB_x'], 'Cantidad en UMA': r.get('Cantidad en UMA'),
                        'PsEx': r.get('PsEx'), 'UxC': r.get('UxC'), 'CxP': r.get('CxP'), 'CxZ': r.get('CxZ'),
                        'Pallets': r.get('Pallets'), 'Cajas': r.get('Cajas'), 'Unidades': r.get('Unidades'),
                        'CajasTotales': r.get('CajasTotales')
                    })

    incompletos_df = pd.DataFrame(incompletos)
    pallets_rearmados = []

    if not incompletos_df.empty:
        incompletos_df = ordenar_skus(incompletos_df)
        for ruta, grupo in incompletos_df.groupby('Ruta'):
            grupo = grupo.reset_index(drop=True)
            usados = [False] * len(grupo)
            pallets_ruta = []

            while not all(usados):
                actual_pallet = []
                total_pallet = 0
                progreso = False
                for i, row in grupo.iterrows():
                    if usados[i]:
                        continue
                    pal_dist = row['PalletsDist']
                    if total_pallet + pal_dist <= 1 + margen:
                        actual_pallet.append(row.to_dict())
                        total_pallet += pal_dist
                        usados[i] = True
                        progreso = True
                    if 1 <= total_pallet <= 1 + margen:
                        break
                if actual_pallet:
                    pallets_ruta.append(actual_pallet)
                elif not progreso:
                    break

            pallets_finales = []
            incompletos_tmp = []
            for pallet in pallets_ruta:
                total = sum(item['PalletsDist'] for item in pallet)
                if total < 1:
                    incompletos_tmp.append((pallet, total))
                else:
                    pallets_finales.append(pallet)

            usados_comb = [False] * len(incompletos_tmp)
            for i in range(len(incompletos_tmp)):
                if usados_comb[i]:
                    continue
                base_pallet, base_total = incompletos_tmp[i]
                usados_comb[i] = True
                combinado = base_pallet.copy()
                total = base_total
                for j in range(i + 1, len(incompletos_tmp)):
                    if usados_comb[j]:
                        continue
                    otro_pallet, otro_total = incompletos_tmp[j]
                    if total + otro_total <= 1.20:
                        combinado.extend(otro_pallet)
                        total += otro_total
                        usados_comb[j] = True
                pallets_finales.append(combinado)

            pallets_ruta = pallets_finales

            def unir_pallets_pequenos(pallets, umbral=0.15):
                volumenes = [sum(row['PalletsDist'] for row in pallet) for pallet in pallets]
                i = 0
                while i < len(pallets):
                    if volumenes[i] < umbral:
                        pallet_pequeno = pallets.pop(i)
                        volumen_pequeno = volumenes.pop(i)
                        if not pallets:
                            break
                        idx_menor = volumenes.index(min(volumenes))
                        pallets[idx_menor].extend(pallet_pequeno)
                        volumenes[idx_menor] += volumen_pequeno
                        i = 0
                    else:
                        i += 1
                return pallets

            pallets_ruta = unir_pallets_pequenos(pallets_ruta)
            for pallet_num, pallet in enumerate(pallets_ruta, start=1):
                for row in pallet:
                    row['PalletsNum'] = pallet_num
                pallets_rearmados.append(pallet)

    ultimo_pallet_por_ruta = {}
    for p in pallets:
        ruta = p['Ruta']
        pid = p['PalletsNum']
        if ruta not in ultimo_pallet_por_ruta:
            ultimo_pallet_por_ruta[ruta] = pid
        else:
            ultimo_pallet_por_ruta[ruta] = max(ultimo_pallet_por_ruta[ruta], pid)

    pallets_montacarga = []
    for (ruta, zona), grupo in pallets_montacarga_df.groupby(['Ruta', 'Zona']):
        pid = ultimo_pallet_por_ruta.get(ruta, 0) + 1
        for _, row in grupo.iterrows():
            cantidad_pm = int(row['Pallets'])
            for _ in range(cantidad_pm):
                pallets_montacarga.append({
                    'Auxiliar': row['Auxiliar'], 'Ruta': ruta, 'Zona': zona,
                    'Material': row['Material'], 'Grupo': row['Grupo'], 'Volumen': row['Volumen'],
                    'PalletsNum': pid, 'Número de material': row['Número de material'], 'Transporte': row['Transporte'],
                    'CajasAdB': 0, 'Cantidad en UMA': row.get('Cantidad en UMA'), 'PsEx': row.get('PsEx'),
                    'UxC': row.get('UxC'), 'CxP': row.get('CxP'), 'CxZ': row.get('CxZ'),
                    'Pallets': row.get('Pallets'), 'Cajas': row.get('Cajas'), 'Unidades': row.get('Unidades'),
                    'CajasTotales': row.get('CajasTotales'), 'PalletsDist': 0
                })
                pid += 1
        ultimo_pallet_por_ruta[ruta] = pid - 1

    auxiliares_df = asignaciones_df[['Auxiliar']].drop_duplicates().reset_index(drop=True)
    pallets_df = pd.DataFrame(pallets)
    carga_actual = pallets_df.groupby('Auxiliar')['CajasAdB'].sum().to_dict() if not pallets_df.empty else {}

    for aux in auxiliares_df['Auxiliar']:
        if aux not in carga_actual:
            carga_actual[aux] = 0

    for pallet_group in pallets_rearmados:
        total_cajas_pallet = sum([r['CajasAdB'] for r in pallet_group])
        aux_min = min(carga_actual, key=carga_actual.get)
        for r in pallet_group:
            pallets.append({
                'Auxiliar': aux_min, 'Ruta': r['Ruta'], 'Zona': r['Zona'],
                'Material': r['Material'], 'Grupo': r['Grupo'], 'Volumen': r['Volumen'],
                'PalletsNum': pallet_id, 'Número de material': r['Número de material'], 'Transporte': r['Transporte'],
                'CajasAdB': r['CajasAdB'], 'Cantidad en UMA': r.get('Cantidad en UMA'), 'PsEx': r.get('PsEx'),
                'UxC': r.get('UxC'), 'CxP': r.get('CxP'), 'CxZ': r.get('CxZ'),
                'Pallets': r.get('Pallets'), 'Cajas': r.get('Cajas'), 'Unidades': r.get('Unidades'),
                'CajasTotales': r.get('CajasTotales'), 'PalletsDist': r.get('PalletsDist')
            })
        pallet_id += 1
        carga_actual[aux_min] += total_cajas_pallet

    pallets_df_final = pd.DataFrame(pallets + pallets_montacarga)
    pallets_df_final['PalletsNum'] = pallets_df_final['PalletsNum'].astype(int)

    notas_por_pallet = (
        pallets_df_final.groupby(['Ruta', 'PalletsNum'])['Zona']
        .apply(lambda zonas: ', '.join(zonas.drop_duplicates()))
        .reset_index()
        .rename(columns={'Zona': 'Notas'})
    )
    pallets_df_final = pallets_df_final.merge(notas_por_pallet, on=['Ruta', 'PalletsNum'], how='left')

    def ordenar_skus_final(df):
        if 'VolumenNum' not in df.columns:
            df['Volumen'] = df['Volumen'].astype(str)
            df['VolumenNum'] = pd.to_numeric(df['Volumen'].str.extract(r'(\d+)', expand=False), errors='coerce')
        df['Prioridad'] = df['Grupo'].apply(lambda x: 20000 if x == 'GRB' else 0)
        df['VolumenOrden'] = df['Prioridad'] + df['VolumenNum'].fillna(0)
        return df.sort_values(by=['Ruta', 'VolumenOrden'], ascending=[True, False])

    pallets_df_final = ordenar_skus_final(pallets_df_final)
    pallets_df_final['PalletsNum'] = (
        pallets_df_final.groupby('Ruta')['PalletsNum'].transform(lambda x: pd.factorize(x)[0] + 1)
    )
    pallets_df_final.drop(columns='Prioridad', inplace=True)
    pallets_df_final.drop(columns='VolumenNum', inplace=True)
    pallets_df_final.drop(columns='VolumenOrden', inplace=True)
    pallets_df_final = pallets_df_final.sort_values(by=['Ruta', 'PalletsNum']).reset_index(drop=True)
    pallets_df_final['Tipo'] = pallets_df_final['CajasAdB'].apply(lambda x: 'Montacargas' if x == 0 else 'Auxiliar')

    columnas_almacenar = ['Auxiliar', 'Ruta', 'Zona', 'Material', 'Grupo', 'Volumen', 'PalletsNum', 'Número de material', 'Transporte', 'CajasAdB', 'Cantidad en UMA', 'PsEx']
    pallets_df_final[columnas_almacenar].to_csv('/tmp/salida.csv', index=False, sep=';', quoting=csv.QUOTE_MINIMAL)

    resumen = []
    for aux, group_aux in pallets_df_final.groupby('Auxiliar'):
        total_auxiliar = 0
        zonas_resumen = []

        for zona, group_zona in group_aux.groupby('Zona'):
            rutas_resumen = []
            for ruta, group_ruta in group_zona.groupby('Ruta'):
                cajas_ruta = float(group_ruta['CajasAdB'].sum())
                total_auxiliar += cajas_ruta
                rutas_resumen.append({"ruta": str(ruta), "cajas": round(cajas_ruta, 2)})
            zonas_resumen.append({"zona": str(zona), "rutas": rutas_resumen})
        resumen.append({"auxiliar": str(aux), "total_cajas": round(total_auxiliar, 2), "zonas": zonas_resumen})

    detalle = []
    for _, row in pallets_df_final.iterrows():
        detalle.append({
            "auxiliar": str(row['Auxiliar']),

            "ruta": str(row['Ruta']),
            "zona": str(row['Zona']),
            "pallet": int(row['PalletsNum']),
            "tipo": str(row['Tipo']),
            "material": str(row['Material']),
            "nombre": str(row.get('Número de material', '')),
            "grupo": str(row.get('Grupo', '')),
            "cajas": round(float(row['CajasAdB']), 2) if row['CajasAdB'] else 0,
            "notas": str(row.get('Notas', ''))
        })

    resultado["success"] = True
    resultado["resumen"] = resumen
    resultado["detalle"] = detalle
    resultado["total_pallets"] = int(pallets_df_final.groupby(['Ruta','PalletsNum']).ngroups)
    resultado["total_cajas"] = round(float(pallets_df_final['CajasAdB'].sum()), 2)
    resultado["total_auxiliares"] = int(pallets_df_final['Auxiliar'].nunique())
    return json.dumps(resultado)
`;

class Component extends DCLogic {
  static LOGIN_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/1272c8747e47415e87d333ed83156f22/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=d5JYG1CL_sZQ_mAScwvjGgyBFceczTSdr0f9QSSuVEw';
  static GET_PRODUCTOS_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/39902af87b044795a3decb034901e384/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=NolDc3GAE7W4AEYEQTqt-iyyFv3s8VOTG0UvaaOswKo';
  static GET_USUARIOS_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0f8e3c5f1faa40c28a35f674b196da9e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=8v-lMooFg6qPG3PLkPNVbcKGkZzvqV9B-xj0_uZhvaw';
  static POST_PRODUCTO_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/57303143ba1c4607a6275266e701f140/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=nxra48UTQ0ZBvGTarjo4JQg4wNBFNZK2G44GAfa5BkU';
  static POST_USUARIO_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7816c6906db1497893b33fdd0ef8e8df/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=mAdRa5NKJtVeQ2M3yRSPr5hY91TyGttwVm1_g7xJF8c';
  static PUT_USUARIO_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/bdad055f65c64a4c8935bc5f00fd1d31/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=3n7SHApRfmrQzw01Ut13lwlIgn7F8RN0d84tKGQ3Keo';
  static DELETE_PRODUCTO_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/971f36a3c5b54f519bf1e0e8e43d29b7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=m6y424PEQsqtOg5tBVqq7JQFlIq5DmIahf-u5waYq6s';
  static DELETE_USUARIO_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/66d1a1e1933a4cf585eefb68a2ca59b8/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=095mNHzkqksV7esvvsEtWtvK3vapDUFSWhrnwW68r8I';
  static PUT_PRODUCTO_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0b3e56d95f044a9cad36411986c5d595/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=VT-4OCavE_BZU7N5hIB1sp1mkKtIE22qy4C-_sBWFzI';
  static DELETE_ASIGNACIONES_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/b99c3dcb2e1c4b9696a4736ca4ca5bbe/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=KAxhbdOOpSAU0knlMfT336tYJgawt4oyCzs3f5IsLiE';
  static GET_SUBIDA_ARME_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/cb6a67d259644ff09f2c9a50ee4a0837/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=FSshi7EUd80wGzcXprp4yfCCyCcKN5JI3Xdh-PkT0TQ';
  static POST_SUBIDA_ARME_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/f1cdc46a4aa44fc1b2dff8cc53a1be92/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=qwjzjuk0rZqTWmOCmjDOMzCYqzKfNx8iyBIJ8Jner-4';
  static GET_ASIGNACIONES_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/59edb00f78ed41058315a5e55acf510c/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=dxNpgb0H4TQqmRNh16kKtyk6dPzPf2RsCOwgQ7Cu3ro';
  static POST_ASIGNACIONES_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/61f68a198f0e487ea621e05b7f2fdae2/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=kRjPv11-pf1bAGurqFpWWuaU9afZGOF7jWSan1i8kN0';
  static GET_PROCESO_ARMADO_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/c6c3095b0a054c05981f44e8cee52d9e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=kCsXTQUVxdoT-UMjTxs_NExyLoLlustAcMNtUvmB0TA';
  static POST_PROCESO_ARMADO_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/b483d66c79b945cbb5d94d0d7c1a2783/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=yzTAyZE9q6YEUDe22wtjw8BeK8J932D-DpNX2vELKKo';
  static PUT_PROCESO_ARMADO_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/249e7f1de7c240da89f1c7c7da641395/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=xfUNhgc0CXFqqtxnp6C7sUOq41UD3tVVwoImVPw_7vM';
  static GET_VERIFICACION_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/132bdd3b58a043939f9b1a5f856e2fd3/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=GGtKtKhcPsZGTyJU7alQ6lvK4XU5PTxxH9tq9BOV6ok';
  static PUT_VERIFICACION_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/dc76a2d5153941039579ef123c977eca/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=cSpsImUbFb9f7KuHQ_xydLT2QNSCqsODgPxBC7QQpjs';
  static POST_VERIFICACION_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/48919f5040d44b4985d5623047e71cae/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=CInwEVttIOmOF7g2o_XGivo4aZwSl9-2iElE4-oOXPA';
  static GET_RESUMEN_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/db67e9dda7434a0eaee2bc4b2ea7a19a/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=nKIT5DB67Nw4cptms7gQjrQTakaFxwsyk2683icHcZw';
  static POST_RESUMEN_FLOW_URL = 'https://defaulte9193073ba8b4e388647ba66872708.21.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d3318cef98ff4d2793858d8c16a44529/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=FbiCRjLLCyuC2uazyCaoSPKUbj-wsWK7bGodWZ3ce6s';
  state = {
    screen: 'splash',
    tab: 'armado',
    usuarios: [],
    usuariosSaving: false,
    addUsuarioOpen: false,
    newUsuarioForm: { nombre: '', apellido: '', role: 'auxiliar' },
    deleteUsuarioConfirm: null,
    productoEditIdx: null,
    productoEditBackup: null,
    productosExpanded: {},
    usuarioEditIdx: null,
    usuarioEditBackup: null,
    addProductoOpen: false,
    newProducto: { Codigo: '', Nombre_Material: '', IsPicked: '', Categoria: '', Grupo: '', Volumen: '', UnitBox: '', BoxPallet: '' },
    deleteConfirm: null,
    productos: [],
    authUser: null,
    authUsername: '',
    authPassword: '',
    authError: '',
    files: { propick: null, master: null },
    auxRows: [{ Nombre: '' }],
    pyodideReady: false,
    statusText: 'Preparando motor Python en el navegador (una sola vez)...',
    consoleLines: [],
    resultsVisible: false,
    stats: { pallets: 0, cajas: 0, aux: 0 },
    auxSummaryRaw: [],
    missingVisible: false,
    missingList: [],
    loginUser: '',
    loginPass: '',
    loginError: '',
    armeVisible: false,
    armeNombre: '',
    armeResumen: '',
    armeRutasRaw: {},
    itemStatus: {},
    qtyDrafts: {},
    filterArmeAuxiliar: 'Todos',
    filterArmeRuta: 'Todas',
    armeDeleteConfirm: null,
    procesarBusy: false,
    procesoHistorial: [],
    armadoPorOperador: {},
    verificacionRegistro: {},
    verifQtyDrafts: {},
    verifFiltroOperador: '',
    verifFiltroRuta: '',
    sqlSyncStatus: null,
    sqlSyncDetalle: ''
  };

  parseDelimLine(line, delim) {
    const result = [];
    let cur = '', inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (inQuotes) {
        if (c === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else inQuotes = false; }
        else cur += c;
      } else {
        if (c === '"') inQuotes = true;
        else if (c === delim) { result.push(cur); cur = ''; }
        else cur += c;
      }
    }
    result.push(cur);
    return result;
  }

  buildAsignacionesRows() {
    if (!this.csvAlmacenado) return [];
    const lines = this.csvAlmacenado.split('\n').filter(l => l.length);
    if (lines.length < 2) return [];
    const headers = this.parseDelimLine(lines[0], ';');
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
      const vals = this.parseDelimLine(lines[i], ';');
      const obj = {};
      headers.forEach((h, idx) => obj[h] = vals[idx]);
      const cajasVal = parseFloat(obj['CajasAdB']) || 0;
      rows.push({
        Auxiliar: obj['Auxiliar'] || '',
        Ruta: String(parseInt(obj['Ruta'], 10) || 0),
        Zona: obj['Zona'] || '',
        PalletNum: parseInt(obj['PalletsNum'], 10) || 0,
        Material: obj['Material'] || '',
        Nombre_Material: obj['Número de material'] || '',
        Grupo: obj['Grupo'] || '',
        Transporte: String(parseInt(obj['Transporte'], 10) || 0),
        Cajas: cajasVal,
        'Cantidad en UMA': parseFloat(obj['Cantidad en UMA']) || 0,
        PsEx: obj['PsEx'] || '',
        Tipo: cajasVal > 0 ? 'Auxiliar' : 'Montacargas'
      });
    }
    return rows;
  }

  generarProcesoId() {
    const d = new Date();
    const pad = n => String(n).padStart(2, '0');
    return 'P-' + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
  }

  async fetchConTimeout(url, body, ms = 15000) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), ms);
    try {
      return await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body), signal: ctrl.signal });
    } finally {
      clearTimeout(timer);
    }
  }

  async guardarEnSQL(resultado) {
    this.setState({ sqlSyncStatus: 'pending', sqlSyncDetalle: 'Guardando en SQL...' });
    const procesoId = this.generarProcesoId();
    this.currentProcesoId = procesoId;
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const fecha = now.getFullYear() + '-' + pad(now.getMonth() + 1) + '-' + pad(now.getDate());
    const hora = pad(now.getHours()) + ':' + pad(now.getMinutes()) + ':' + pad(now.getSeconds());
    const header = {
      ProcesoID: procesoId,
      Fecha: fecha,
      Hora: hora,
      TotalPallets: resultado.total_pallets,
      TotalCajas: resultado.total_cajas,
      TotalAuxiliares: resultado.total_auxiliares,
      UsuarioProceso: (this.state.authUser && this.state.authUser.username) || this.state.authUsername || ''
    };

    this.logMsg('Guardando encabezado en Propick_Subida_de_arme...', 'info');
    try {
      const resp = await this.fetchConTimeout(Component.POST_SUBIDA_ARME_FLOW_URL, header);
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      this.logMsg('✅ Encabezado guardado (ProcesoID ' + procesoId + ')', 'ok');
    } catch (e) {
      const msg = e.name === 'AbortError' ? 'tiempo de espera agotado (posible URL de flujo sin firma/sig o inactivo)' : e.message;
      this.logMsg('❌ Error guardando encabezado en Propick_Subida_de_arme: ' + msg, 'err');
      this.setState({ sqlSyncStatus: 'error', sqlSyncDetalle: 'Falló el guardado del encabezado: ' + msg });
      return;
    }

    const rows = this.buildAsignacionesRows();
    if (rows.length === 0) {
      this.logMsg('⚠️ No hay líneas de asignación para guardar (CSV vacío o no generado).', 'err');
      this.setState({ sqlSyncStatus: 'error', sqlSyncDetalle: 'Encabezado guardado, pero no se encontraron líneas de asignación.' });
      return;
    }

    this.logMsg('Guardando ' + rows.length + ' líneas en Propick_Asignaciones (en un solo lote)...', 'info');
    try {
      const resp = await this.fetchConTimeout(Component.POST_ASIGNACIONES_FLOW_URL, { ProcesoID: procesoId, Rows: rows }, 60000);
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      this.logMsg('✅ ' + rows.length + ' líneas guardadas en Propick_Asignaciones', 'ok');
      this.setState({ sqlSyncStatus: 'ok', sqlSyncDetalle: procesoId + ' · ' + rows.length + ' líneas guardadas' });
    } catch (e) {
      const msg = e.name === 'AbortError' ? 'tiempo de espera agotado' : e.message;
      this.logMsg('⚠️ Error guardando el lote en Propick_Asignaciones: ' + msg, 'err');
      this.setState({ sqlSyncStatus: 'error', sqlSyncDetalle: 'Encabezado guardado, pero falló el lote de líneas — ' + msg });
    }
  }

  async loadProductos() {
    if (this.state.productosLoading) return;
    this.setState({ productosLoading: true, productosError: '' });
    try {
      const resp = await fetch(Component.GET_PRODUCTOS_FLOW_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      const data = await resp.json();
      const rows = Array.isArray(data) ? data : (Array.isArray(data.value) ? data.value : []);
      const productos = rows.map(r => ({
        ItemInternalId: r.ItemInternalId ?? null,
        Codigo: r.Codigo ?? r.codigo ?? '',
        Nombre_Material: r.Nombre_Material ?? r.nombre_material ?? '',
        IsPicked: (r.IsPicked === true || r.IsPicked === 'Si') ? 'Si' : (r.IsPicked === false || r.IsPicked === 'No') ? 'No' : (r.IsPicked ?? r.isPicked ?? ''),
        Categoria: r.Categoria ?? r.categoria ?? '',
        Grupo: r.Grupo ?? r.grupo ?? '',
        Volumen: r.Volumen ?? r.volumen ?? '',
        NameZone: r.NameZone ?? r.nameZone ?? '',
        ZoneCode: r.ZoneCode ?? r.zoneCode ?? '',
        DescriptionZone: r.DescriptionZone ?? r.descriptionZone ?? '',
        UnitBox: r.UnitBox ?? r.unitBox ?? '',
        BoxPallet: r.BoxPallet ?? r.boxPallet ?? ''
      }));
      this.setState({ productos, productosLoading: false });
    } catch (e) {
      this.setState({ productosLoading: false, productosError: 'No se pudieron cargar los productos. Intenta nuevamente.' });
    }
  }

  async getRows(url) {
    const resp = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const data = await resp.json();
    return Array.isArray(data) ? data : (Array.isArray(data.value) ? data.value : []);
  }

  async loadUltimoProceso() {
    if (this.loadingUltimoProceso) return;
    this.loadingUltimoProceso = true;
    try {
      const subidas = await this.getRows(Component.GET_SUBIDA_ARME_FLOW_URL);
      if (!subidas.length) { this.loadingUltimoProceso = false; return; }
      const norm = subidas.map(r => ({
        ProcesoID: r.ProcesoID ?? r.procesoID ?? '',
        Fecha: r.Fecha ?? r.fecha ?? '',
        Hora: r.Hora ?? r.hora ?? ''
      }));
      norm.sort((a, b) => (a.Fecha + a.Hora).localeCompare(b.Fecha + b.Hora));
      const ultimo = norm[norm.length - 1];
      const ultimoIdNorm = String(ultimo.ProcesoID).trim().toLowerCase();
      this.currentProcesoId = ultimo.ProcesoID;

      const asigRows = await this.getRows(Component.GET_ASIGNACIONES_FLOW_URL);
      const detalle = asigRows.filter(r => String(r.ProcesoID ?? r.procesoID ?? '').trim().toLowerCase() === ultimoIdNorm).map(r => ({
        auxiliar: r.Auxiliar ?? r.auxiliar ?? '',
        ruta: String(r.Ruta ?? r.ruta ?? ''),
        zona: r.Zona ?? r.zona ?? '',
        pallet: parseInt(r.PalletNum ?? r.palletNum ?? 0, 10) || 0,
        tipo: r.Tipo ?? r.tipo ?? 'Auxiliar',
        material: r.Material ?? r.material ?? '',
        nombre: r.Nombre_Material ?? r.nombre_material ?? '',
        grupo: r.Grupo ?? r.grupo ?? '',
        cajas: parseFloat(r.Cajas ?? r.cajas ?? 0) || 0,
        notas: r.Zona ?? r.zona ?? ''
      }));
      if (!detalle.length) { this.loadingUltimoProceso = false; return; }

      const porAux = {};
      detalle.forEach(d => {
        if (!porAux[d.auxiliar]) porAux[d.auxiliar] = { total_cajas: 0, zonasMap: {} };
        porAux[d.auxiliar].total_cajas += d.cajas;
        if (!porAux[d.auxiliar].zonasMap[d.zona]) porAux[d.auxiliar].zonasMap[d.zona] = {};
        porAux[d.auxiliar].zonasMap[d.zona][d.ruta] = (porAux[d.auxiliar].zonasMap[d.zona][d.ruta] || 0) + d.cajas;
      });
      const resumen = Object.keys(porAux).map(aux => ({
        auxiliar: aux,
        total_cajas: Math.round(porAux[aux].total_cajas * 100) / 100,
        zonas: Object.keys(porAux[aux].zonasMap).map(zona => ({
          zona,
          rutas: Object.keys(porAux[aux].zonasMap[zona]).map(ruta => ({ ruta, cajas: Math.round(porAux[aux].zonasMap[zona][ruta] * 100) / 100 }))
        }))
      }));
      const totalPallets = new Set(detalle.map(d => d.ruta + '|' + d.pallet)).size;
      const totalCajas = Math.round(detalle.reduce((sum, d) => sum + d.cajas, 0) * 100) / 100;
      const totalAux = new Set(detalle.map(d => d.auxiliar)).size;
      this.lastResultado = { success: true, detalle, resumen, total_pallets: totalPallets, total_cajas: totalCajas, total_auxiliares: totalAux };
      this.setState({ stats: { pallets: totalPallets, cajas: totalCajas, aux: totalAux } });

      const armadoRows = await this.getRows(Component.GET_PROCESO_ARMADO_FLOW_URL);
      const armadoPorOperador = {};
      this.procesoArmadoIds = this.procesoArmadoIds || {};
      armadoRows.filter(r => String(r.ProcesoID ?? r.procesoID ?? '').trim().toLowerCase() === ultimoIdNorm).forEach(r => {
        const operador = r.Auxiliar ?? r.auxiliar ?? '';
        const ruta = String(r.Ruta ?? r.ruta ?? '');
        const pnum = String(r.PalletNum ?? r.palletNum ?? '');
        const material = r.Codigo ?? r.codigo ?? '';
        const key = 'r' + ruta + '_p' + pnum + '_m' + material;
        if (!armadoPorOperador[operador]) armadoPorOperador[operador] = {};
        const estadoRaw = String(r.Estado ?? r.estado ?? '').toLowerCase();
        armadoPorOperador[operador][key] = {
          material,
          nombre: r.NombreMaterial ?? r.nombreMaterial ?? '',
          zona: r.Zona ?? r.zona ?? '',
          cajasPlan: parseFloat(r.CajasPedidas ?? r.cajasPedidas ?? 0) || 0,
          cajasReal: parseFloat(r.CajasReales ?? r.cajasReales ?? 0) || 0,
          estado: estadoRaw === 'confirmado' ? 'confirmado' : estadoRaw === 'problema' ? 'problema' : 'pendiente'
        };
        if (r.InternalID) this.procesoArmadoIds[key] = r.InternalID;
      });

      const verifRows = await this.getRows(Component.GET_VERIFICACION_FLOW_URL);
      const verificacionRegistro = {};
      verifRows.filter(r => String(r.ProcesoID ?? r.procesoID ?? '').trim().toLowerCase() === ultimoIdNorm).forEach(r => {
        const operador = r.Auxiliar ?? r.auxiliar ?? '';
        const ruta = String(r.Ruta ?? r.ruta ?? '');
        const pnum = String(r.PalletNum ?? r.palletNum ?? '');
        const material = r.Codigo ?? r.codigo ?? '';
        const key = 'r' + ruta + '_p' + pnum + '_m' + material;
        if (!verificacionRegistro[operador]) verificacionRegistro[operador] = {};
        const resultado = String(r.Resultado ?? r.resultado ?? '').toLowerCase();
        verificacionRegistro[operador][key] = resultado === 'corregido'
          ? { estado: 'corregido', cantidad: parseFloat(r.CantidadCorregida ?? r.cantidadCorregida ?? 0) || 0 }
          : { estado: 'correcto' };
      });

      this.setState({ armadoPorOperador, verificacionRegistro });
    } catch (e) {
      this.logMsg('⚠️ No se pudo recuperar el último proceso desde SQL: ' + e.message, 'err');
    }
    this.loadingUltimoProceso = false;
  }

  async loadUsuarios() {
    if (this.state.usuariosLoading) return;
    this.setState({ usuariosLoading: true, usuariosError: '' });
    try {
      const resp = await fetch(Component.GET_USUARIOS_FLOW_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      const data = await resp.json();
      const rows = Array.isArray(data) ? data : (Array.isArray(data.value) ? data.value : []);
      const usuarios = rows.map(r => ({
        UserInternalId: r.UserInternalId ?? r.ItemInternalId ?? r.userinternalid ?? r.iteminternalid ?? null,
        nombre: r.Nombre ?? r.nombre ?? '',
        apellido: r.Apellido ?? r.apellido ?? '',
        usuario: r.Usuario ?? r.usuario ?? '',
        password: String(r.Contraseña ?? r.contraseña ?? r.password ?? ''),
        role: String(r.Rol ?? r.rol ?? '').trim().toLowerCase()
      }));
      this.setState({ usuarios, usuariosLoading: false });
    } catch (e) {
      this.setState({ usuariosLoading: false, usuariosError: 'No se pudieron cargar los usuarios. Intenta nuevamente.' });
    }
  }

  componentDidMount() {
    document.body.classList.add('pp-has-app');
    this.initPyodide();
    this._lineId = 0;
    this.downloadUrl = null;
    this.lastResultado = null;
    this.currentProcesoId = null;
    this.verifInicioPallet = {};
    this.resumenEnviado = {};
    this.procesoArmadoIds = {};
    this.loadingUltimoProceso = false;
  }

  scrollConsola() {
    const el = document.querySelector('.pp-console');
    if (el) el.scrollTop = el.scrollHeight;
  }

  logMsg(msg, cls) {
    setTimeout(() => this.scrollConsola(), 0);
    const t = new Date().toLocaleTimeString('es-EC');
    this._lineId += 1;
    this.setState(s => ({ consoleLines: [...s.consoleLines, { id: this._lineId, cls: cls || '', text: '[' + t + '] ' + msg }] }));
  }

  async initPyodide() {
    try {
      this.logMsg('Iniciando entorno Python (Pyodide)...', 'info');
      this.pyodide = await loadPyodide();
      this.logMsg('Cargando pandas...', 'info');
      await this.pyodide.loadPackage(['pandas']);
      try {
        this.logMsg('Cargando openpyxl...', 'info');
        await this.pyodide.loadPackage(['openpyxl']);
      } catch (e) {
        this.logMsg('openpyxl no disponible como paquete nativo, instalando via micropip...', 'info');
        await this.pyodide.loadPackage('micropip');
        const micropip = this.pyodide.pyimport('micropip');
        await micropip.install('openpyxl');
      }
      this.setState({ pyodideReady: true });
      this.logMsg('✅ Motor listo. Puedes procesar cuando cargues los 3 archivos.', 'ok');
    } catch (e) {
      this.logMsg('❌ Error iniciando el motor Python: ' + e, 'err');
      this.setState({ statusText: 'Error al iniciar el motor. Recarga la página.' });
    }
  }

  onFileSelected(key, file) {
    if (!file) return;
    this.setState(s => ({ files: { ...s.files, [key]: file } }));
    this.logMsg('Archivo cargado: ' + file.name, 'info');
  }

  escapeAttr(v) { return String(v); }

  addAuxRow() { this.setState(s => ({ auxRows: [...s.auxRows, { Nombre: '' }] })); }
  removeAuxRow(idx) { this.setState(s => ({ auxRows: s.auxRows.filter((_, i) => i !== idx) })); }
  updateAuxCell(idx, field, value) {
    this.setState(s => {
      const rows = s.auxRows.map((r, i) => i === idx ? { ...r, [field]: value } : r);
      return { auxRows: rows };
    });
  }

  csvField(v) {
    v = (v === undefined || v === null) ? '' : String(v);
    if (v.includes(',') || v.includes('"') || v.includes('\n')) return '"' + v.replace(/"/g, '""') + '"';
    return v;
  }
  buildPersonalCsv() {
    const lines = ['NombreAuxiliar'];
    this.state.auxRows.forEach(row => lines.push(this.csvField(row.Nombre)));
    return lines.join('\n');
  }

  buildMasterCsv() {
    const cols = ['Codigo', 'Nombre_Material', 'IsPicked', 'Categoria', 'Grupo', 'Volumen', 'NameZone', 'ZoneCode', 'DescriptionZone', 'UnitBox', 'BoxPallet'];
    const lines = [cols.join(',')];
    this.state.productos.forEach(p => lines.push(cols.map(c => this.csvField(p[c])).join(',')));
    return lines.join('\n');
  }

  async writeFileToFS(file, path) {
    const buf = await file.arrayBuffer();
    this.pyodide.FS.writeFile(path, new Uint8Array(buf));
  }

  async procesar() {
    this.setState({ procesarBusy: true, missingVisible: false, resultsVisible: false, procesoResultado: null });
    try {
      this.logMsg('Escribiendo archivos en el sistema virtual...', 'info');
      await this.writeFileToFS(this.state.files.propick, '/tmp/propick.xlsx');
      this.pyodide.FS.writeFile('/tmp/master.csv', new TextEncoder().encode(this.buildMasterCsv()));
      this.pyodide.FS.writeFile('/tmp/personal.csv', new TextEncoder().encode(this.buildPersonalCsv()));

      this.logMsg('Ejecutando motor de asignación (pandas)...', 'info');
      await this.pyodide.runPythonAsync(PY_CODE);
      const resultJson = await this.pyodide.runPythonAsync('ejecutar_motor()');
      const resultado = JSON.parse(resultJson);

      if (!resultado.success) {
        if (resultado.materiales_faltantes && resultado.materiales_faltantes.length > 0) {
          this.logMsg('❌ ' + resultado.error, 'err');
          this.setState({ missingVisible: true, missingList: resultado.materiales_faltantes });
        } else {
          this.logMsg('❌ ' + resultado.error, 'err');
        }
        this.setState({ procesarBusy: false, procesoResultado: { ok: false, titulo: 'El proceso no se completó', detalle: resultado.error || 'Revisa el detalle en la consola.' } });
        return;
      }

      this.lastResultado = resultado;
      this.logMsg('✅ Proceso completado: ' + resultado.total_pallets + ' pallets, ' + resultado.total_cajas + ' cajas, ' + resultado.total_auxiliares + ' auxiliares', 'ok');
      this.setState({ procesoResultado: { ok: true, titulo: 'Proceso completado', detalle: resultado.total_pallets + ' pallets · ' + resultado.total_cajas + ' cajas · ' + resultado.total_auxiliares + ' auxiliares' } });

      const csvTexto = new TextDecoder('utf-8').decode(this.pyodide.FS.readFile('/tmp/salida.csv'));
      this.csvAlmacenado = csvTexto; // simula el guardado en base de datos (tabla Asignacion)
      this.logMsg('💾 Asignación guardada en base de datos', 'ok');

      this.setState(s => ({
        stats: { pallets: resultado.total_pallets, cajas: resultado.total_cajas, aux: resultado.total_auxiliares },
        auxSummaryRaw: resultado.resumen,
        resultsVisible: true,
        procesarBusy: false,
        guardadoEnBD: true,
        procesoHistorial: [...s.procesoHistorial, { hora: new Date().toLocaleTimeString('es-EC'), pallets: resultado.total_pallets, cajas: resultado.total_cajas, aux: resultado.total_auxiliares }]
      }));
      this.guardarEnSQL(resultado);
    } catch (e) {
      this.logMsg('❌ Error durante el procesamiento: ' + e, 'err');
      this.setState({ procesarBusy: false, procesoResultado: { ok: false, titulo: 'Error durante el procesamiento', detalle: String(e) } });
    }
  }

  cargarDatosEjemplo() {
    const detalle = [
      { auxiliar: 'Julio De La Cruz', ruta: '101', zona: 'Zona 2', pallet: 1, tipo: 'Auxiliar', material: 'AA004001', nombre: 'Pepsi 355 Ml Lata X 24', grupo: 'Lata', cajas: 24, notas: 'Zona 2' },
      { auxiliar: 'Julio De La Cruz', ruta: '101', zona: 'Zona 2', pallet: 1, tipo: 'Auxiliar', material: 'BA017072', nombre: 'Pepsi Sabor Intenso 355 ML GRB X 24', grupo: 'Lata', cajas: 2.25, notas: 'Zona 2' },
      { auxiliar: 'Julio De La Cruz', ruta: '102', zona: 'Zona 3', pallet: 1, tipo: 'Auxiliar', material: 'BA000798', nombre: 'Manzana 3785 ml Pet x 4', grupo: 'PET', cajas: 1.25, notas: 'Zona 3' },
      { auxiliar: 'Nixon Engracia', ruta: '103', zona: 'Zona 1', pallet: 1, tipo: 'Auxiliar', material: 'BA005969', nombre: 'Mas Naranja 3785 ml Pet x 4', grupo: 'PET', cajas: 2, notas: 'Zona 1' },
      { auxiliar: 'Nixon Engracia', ruta: '103', zona: 'Zona 1', pallet: 2, tipo: 'Montacargas', material: 'BA000786', nombre: 'Tropical Fresa 3785 ml Pet x 4', grupo: 'PET', cajas: 20, notas: 'Zona 1' }
    ];
    const resumen = [
      { auxiliar: 'Julio De La Cruz', total_cajas: 27.5, zonas: [{ zona: 'Zona 2', rutas: [{ ruta: '101', cajas: 26.25 }] }, { zona: 'Zona 3', rutas: [{ ruta: '102', cajas: 1.25 }] }] },
      { auxiliar: 'Nixon Engracia', total_cajas: 22, zonas: [{ zona: 'Zona 1', rutas: [{ ruta: '103', cajas: 22 }] }] }
    ];
    const resultado = { success: true, detalle, resumen, total_pallets: 4, total_cajas: 49.5, total_auxiliares: 2 };
    this.lastResultado = resultado;
    if (!this.currentProcesoId) this.currentProcesoId = 'P-DEMO';

    const key0 = 'r101_p1_mAA004001', key1 = 'r101_p1_mBA017072', key2 = 'r102_p1_mBA000798';
    const armadoJulio = {
      [key0]: { material: 'AA004001', nombre: 'Pepsi 355 Ml Lata X 24', cajasPlan: 24, cajasReal: 24, estado: 'confirmado' },
      [key1]: { material: 'BA017072', nombre: 'Pepsi Sabor Intenso 355 ML GRB X 24', cajasPlan: 2.25, cajasReal: 1, estado: 'problema' },
      [key2]: { material: 'BA000798', nombre: 'Manzana 3785 ml Pet x 4', cajasPlan: 1.25, cajasReal: 1.25, estado: 'confirmado' }
    };

    this.setState(s => ({
      stats: { pallets: resultado.total_pallets, cajas: resultado.total_cajas, aux: resultado.total_auxiliares },
      auxSummaryRaw: resultado.resumen,
      resultsVisible: true,
      armadoPorOperador: { ...s.armadoPorOperador, 'Julio De La Cruz': armadoJulio },
      procesoHistorial: [...s.procesoHistorial, { hora: new Date().toLocaleTimeString('es-EC') + ' (ejemplo)', pallets: resultado.total_pallets, cajas: resultado.total_cajas, aux: resultado.total_auxiliares }]
    }));
    this.logMsg('✅ Datos de ejemplo cargados (demo, no reemplaza un proceso real).', 'ok');
  }



  intentarLogin() {
    if (!this.lastResultado) { this.setState({ loginError: 'Primero debe procesarse un archivo Propick para poder consultar el arme.' }); return; }
    const usuario = this.state.loginUser.trim();
    if (!usuario) { this.setState({ loginError: 'Selecciona un operador.' }); return; }
    const match = this.lastResultado.resumen.find(a => a.auxiliar.trim().toLowerCase() === usuario.toLowerCase());
    if (!match) { this.setState({ loginError: 'Operador no encontrado en asignaciones.' }); return; }
    this.setState({ loginError: '' });
    this.mostrarArme(match);
  }

  mostrarArme(match) {
    const misFilas = this.lastResultado.detalle.filter(d => d.auxiliar === match.auxiliar);
    const porRuta = {};
    misFilas.forEach(f => {
      if (!porRuta[f.ruta]) porRuta[f.ruta] = {};
      if (!porRuta[f.ruta][f.pallet]) porRuta[f.ruta][f.pallet] = { tipo: f.tipo, zona: f.notas || f.zona, items: [] };
      porRuta[f.ruta][f.pallet].items.push(f);
    });
    const existente = this.state.armadoPorOperador[match.auxiliar];
    const itemStatus = existente ? { ...existente } : {};
    Object.keys(porRuta).forEach(ruta => {
      Object.keys(porRuta[ruta]).forEach(pnum => {
        porRuta[ruta][pnum].items.forEach((it, idx) => {
          const key = 'r' + ruta + '_p' + pnum + '_m' + it.material;
          if (!itemStatus[key]) itemStatus[key] = { material: it.material, nombre: it.nombre, zona: it.zona, cajasPlan: it.cajas, cajasReal: it.cajas, estado: 'pendiente' };
        });
      });
    });
    this.setState(s => ({
      armeVisible: true,
      armeNombre: match.auxiliar,
      armeResumen: match.total_cajas + ' cajas totales asignadas',
      armeRutasRaw: porRuta,
      itemStatus,
      qtyDrafts: {},
      armadoPorOperador: { ...s.armadoPorOperador, [match.auxiliar]: itemStatus }
    }));
  }

  confirmarItem(key) {
    const operador = this.state.armeNombre;
    let cajasPlan;
    this.setState(s => {
      cajasPlan = s.itemStatus[key].cajasPlan;
      const itemStatus = { ...s.itemStatus, [key]: { ...s.itemStatus[key], estado: 'confirmado', cajasReal: cajasPlan } };
      return { itemStatus, armadoPorOperador: { ...s.armadoPorOperador, [operador]: itemStatus } };
    });
    this.postProcesoArmado(operador, key, 'Confirmado', cajasPlan);
  }
  marcarProblema(key) {
    this.setState(s => ({ itemStatus: { ...s.itemStatus, [key]: { ...s.itemStatus[key], estado: 'ajustando' } }, qtyDrafts: { ...s.qtyDrafts, [key]: '0' } }));
  }
  setQtyDraft(key, value) { this.setState(s => ({ qtyDrafts: { ...s.qtyDrafts, [key]: value } })); }
  guardarCantidadReal(key) {
    const cantidad = parseFloat(this.state.qtyDrafts[key]) || 0;
    const operador = this.state.armeNombre;
    this.setState(s => {
      const itemStatus = { ...s.itemStatus, [key]: { ...s.itemStatus[key], estado: 'problema', cajasReal: cantidad } };
      return { itemStatus, armadoPorOperador: { ...s.armadoPorOperador, [operador]: itemStatus } };
    });
    this.postProcesoArmado(operador, key, 'Problema', cantidad);
  }

  async postProcesoArmado(operador, key, estado, cajasReal) {
    const m = key.match(/^r(.+)_p(\d+)_m(.+)$/);
    if (!m) return;
    const ruta = m[1], pnum = m[2];
    const it = (this.state.armadoPorOperador[operador] || {})[key] || this.state.itemStatus[key];
    if (!it) return;
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const fecha = now.getFullYear() + '-' + pad(now.getMonth() + 1) + '-' + pad(now.getDate());
    const hora = pad(now.getHours()) + ':' + pad(now.getMinutes()) + ':' + pad(now.getSeconds());
    const payload = {
      ProcesoID: this.currentProcesoId || '',
      IDPallet: ruta + '-' + pnum,
      Auxiliar: operador,
      Ruta: parseInt(ruta, 10) || 0,
      PalletNum: String(pnum),
      Zona: it.zona || '',
      Codigo: it.material || '',
      NombreMaterial: it.nombre || '',
      CajasPedidas: it.cajasPlan || 0,
      CajasReales: cajasReal != null ? cajasReal : (it.cajasReal || 0),
      Estado: estado,
      Fecha: fecha,
      HoraArmado: hora
    };
    const itemInternalId = this.procesoArmadoIds ? this.procesoArmadoIds[key] : null;
    try {
      if (itemInternalId) {
        await this.fetchConTimeout(Component.PUT_PROCESO_ARMADO_FLOW_URL, { ...payload, InternalID: itemInternalId });
      } else {
        const resp = await this.fetchConTimeout(Component.POST_PROCESO_ARMADO_FLOW_URL, payload);
        if (resp.ok) {
          let created = null;
          try { created = await resp.json(); } catch (e) { /* sin cuerpo de respuesta */ }
          if (created && created.InternalID) {
            if (!this.procesoArmadoIds) this.procesoArmadoIds = {};
            this.procesoArmadoIds[key] = created.InternalID;
          }
        }
      }
      this.logMsg('✅ Armado guardado: ' + payload.Codigo + ' (' + estado + ')', 'ok');
    } catch (e) {
      this.logMsg('❌ Error guardando armado de ' + payload.Codigo + ': ' + e.message, 'err');
    }
  }

  cerrarSesion() {
    if (this.state.authUser && this.state.authUser.role === 'auxiliar') { this.logoutFull(); return; }
    this.setState({ armeVisible: false, loginUser: '', loginPass: '', loginError: '' });
  }

  recalcResumen(detalle) {
    const byAux = {};
    detalle.forEach(f => {
      if (!byAux[f.auxiliar]) byAux[f.auxiliar] = { auxiliar: f.auxiliar, total_cajas: 0, zonaMap: {} };
      const a = byAux[f.auxiliar];
      a.total_cajas += f.cajas;
      const zona = f.notas || f.zona;
      if (!a.zonaMap[zona]) a.zonaMap[zona] = {};
      a.zonaMap[zona][f.ruta] = (a.zonaMap[zona][f.ruta] || 0) + f.cajas;
    });
    return Object.values(byAux).map(a => ({
      auxiliar: a.auxiliar,
      total_cajas: Math.round(a.total_cajas * 100) / 100,
      zonas: Object.keys(a.zonaMap).map(zona => ({
        zona,
        rutas: Object.keys(a.zonaMap[zona]).map(ruta => ({ ruta, cajas: Math.round(a.zonaMap[zona][ruta] * 100) / 100 }))
      }))
    }));
  }

  eliminarAsignacionesQueCoinciden(filterFn, mensaje, sqlPayload) {
    if (!this.lastResultado) return;
    const nuevoDetalle = this.lastResultado.detalle.filter(f => !filterFn(f));
    const nuevoResumen = this.recalcResumen(nuevoDetalle);
    this.lastResultado = {
      ...this.lastResultado,
      detalle: nuevoDetalle,
      resumen: nuevoResumen,
      total_pallets: new Set(nuevoDetalle.map(f => f.auxiliar + '|' + f.ruta + '|' + f.pallet)).size,
      total_cajas: Math.round(nuevoDetalle.reduce((s, f) => s + f.cajas, 0) * 100) / 100,
      total_auxiliares: nuevoResumen.length,
    };
    this.setState({ armeDeleteConfirm: null });
    this.logMsg(mensaje, 'ok');
    if (Component.DELETE_ASIGNACIONES_FLOW_URL) {
      this.fetchConTimeout(Component.DELETE_ASIGNACIONES_FLOW_URL, {
        ProcesoID: this.currentProcesoId || '',
        Auxiliar: sqlPayload.auxiliar || '',
        Ruta: sqlPayload.ruta || '',
        PalletNum: sqlPayload.pnum || ''
      }).then(resp => {
        if (!resp.ok) throw new Error('HTTP ' + resp.status);
      }).catch(e => this.logMsg('⚠️ No se pudo borrar en la base (SQL): ' + e.message + '. El cambio quedó solo local.', 'err'));
    }
  }

  onEliminarPallet(auxiliar, ruta, pnum) {
    this.setState({ armeDeleteConfirm: { tipo: 'pallet', auxiliar, ruta, pnum, texto: 'el Pallet #' + pnum + ' de la ruta ' + ruta + ' (' + auxiliar + ')' } });
  }
  onEliminarRuta(auxiliar, ruta) {
    this.setState({ armeDeleteConfirm: { tipo: 'ruta', auxiliar, ruta, texto: 'toda la ruta ' + ruta + ' de ' + auxiliar } });
  }
  onEliminarTodas() {
    this.setState({ armeDeleteConfirm: { tipo: 'todas', texto: 'TODAS las asignaciones' } });
  }
  cancelarEliminarAsignacion() { this.setState({ armeDeleteConfirm: null }); }
  confirmarEliminarAsignacion() {
    const c = this.state.armeDeleteConfirm;
    if (!c) return;
    if (c.tipo === 'pallet') {
      this.eliminarAsignacionesQueCoinciden(f => f.auxiliar === c.auxiliar && String(f.ruta) === String(c.ruta) && String(f.pallet) === String(c.pnum), '🗑️ Pallet eliminado.', { auxiliar: c.auxiliar, ruta: c.ruta, pnum: c.pnum });
    } else if (c.tipo === 'ruta') {
      this.eliminarAsignacionesQueCoinciden(f => f.auxiliar === c.auxiliar && String(f.ruta) === String(c.ruta), '🗑️ Ruta eliminada.', { auxiliar: c.auxiliar, ruta: c.ruta });
    } else if (c.tipo === 'todas') {
      this.lastResultado = null;
      this.setState({ armeDeleteConfirm: null, armadoPorOperador: {} });
      this.logMsg('🗑️ Todas las asignaciones eliminadas.', 'ok');
      if (Component.DELETE_ASIGNACIONES_FLOW_URL) {
        this.fetchConTimeout(Component.DELETE_ASIGNACIONES_FLOW_URL, { ProcesoID: this.currentProcesoId || '', Auxiliar: '', Ruta: '', PalletNum: '' })
          .then(resp => { if (!resp.ok) throw new Error('HTTP ' + resp.status); })
          .catch(e => this.logMsg('⚠️ No se pudo borrar en la base (SQL): ' + e.message + '. El cambio quedó solo local.', 'err'));
      }
    }
  }

  marcarVerifCorrecto(operador, key) {
    this.setState(s => ({ verificacionRegistro: { ...s.verificacionRegistro, [operador]: { ...s.verificacionRegistro[operador], [key]: { estado: 'correcto' } } } }));
    this.postVerificacionLinea(operador, key, 'Correcto', 0);
  }
  marcarVerifIncorrecto(operador, key, cajasReal) {
    this.setState(s => ({
      verificacionRegistro: { ...s.verificacionRegistro, [operador]: { ...s.verificacionRegistro[operador], [key]: { estado: 'ajustando' } } },
      verifQtyDrafts: { ...s.verifQtyDrafts, [key]: String(cajasReal) }
    }));
  }
  setVerifQtyDraft(key, value) { this.setState(s => ({ verifQtyDrafts: { ...s.verifQtyDrafts, [key]: value } })); }
  guardarVerifCantidad(operador, key) {
    const cantidad = parseFloat(this.state.verifQtyDrafts[key]) || 0;
    this.setState(s => ({ verificacionRegistro: { ...s.verificacionRegistro, [operador]: { ...s.verificacionRegistro[operador], [key]: { estado: 'corregido', cantidad } } } }));
    this.postVerificacionLinea(operador, key, 'Corregido', cantidad);
  }

  nombreVerificadorActual() {
    const u = this.state.authUser;
    if (!u) return '';
    return ((u.nombre || '') + ' ' + (u.apellido || '')).trim() || u.username || '';
  }

  async postVerificacionLinea(operador, key, resultadoTxt, cantidadCorregida) {
    const m = key.match(/^r(.+)_p(\d+)_m(.+)$/);
    if (!m) return;
    const ruta = m[1], pnum = m[2];
    const itemStatus = (this.state.armadoPorOperador[operador] || {})[key];
    if (!itemStatus) return;
    const palletKeyId = operador + '|' + ruta + '|' + pnum;
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const horaActual = pad(now.getHours()) + ':' + pad(now.getMinutes()) + ':' + pad(now.getSeconds());
    const fecha = now.getFullYear() + '-' + pad(now.getMonth() + 1) + '-' + pad(now.getDate());
    if (!this.verifInicioPallet[palletKeyId]) this.verifInicioPallet[palletKeyId] = horaActual;

    const payload = {
      ProcesoID: this.currentProcesoId || '',
      IDPallet: ruta + '-' + pnum,
      Auxiliar: operador,
      Ruta: parseInt(ruta, 10) || 0,
      PalletNum: String(pnum),
      Codigo: itemStatus.material || '',
      CantidadArmada: itemStatus.cajasReal || 0,
      Resultado: resultadoTxt,
      CantidadCorregida: cantidadCorregida != null ? cantidadCorregida : 0,
      Verificador: this.nombreVerificadorActual(),
      Fecha: fecha,
      HoraInicioVerificacion: this.verifInicioPallet[palletKeyId],
      HoraVerificacion: horaActual
    };
    try {
      const resp = await this.fetchConTimeout(Component.POST_VERIFICACION_FLOW_URL, payload);
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      this.logMsg('✅ Verificación guardada: ' + payload.Codigo + ' (' + resultadoTxt + ')', 'ok');
    } catch (e) {
      this.logMsg('❌ Error guardando verificación de ' + payload.Codigo + ': ' + e.message, 'err');
    }
    this.checkYPostResumenPallet(operador, ruta, pnum);
  }

  async checkYPostResumenPallet(operador, ruta, pnum) {
    const itemStatus = this.state.armadoPorOperador[operador] || {};
    const verifOp = this.state.verificacionRegistro[operador] || {};
    const prefix = 'r' + ruta + '_p' + pnum + '_m';
    const keys = Object.keys(itemStatus).filter(k => k.startsWith(prefix));
    if (keys.length === 0) return;
    const todasVerificadas = keys.every(k => verifOp[k] && (verifOp[k].estado === 'correcto' || verifOp[k].estado === 'corregido'));
    if (!todasVerificadas) return;
    const palletKeyId = operador + '|' + ruta + '|' + pnum;
    if (this.resumenEnviado[palletKeyId]) return;
    this.resumenEnviado[palletKeyId] = true;

    const cajasPedidas = keys.reduce((sum, k) => sum + (itemStatus[k].cajasPlan || 0), 0);
    const cajasReales = keys.reduce((sum, k) => {
      const v = verifOp[k];
      const val = (v && v.estado === 'corregido') ? v.cantidad : itemStatus[k].cajasReal;
      return sum + (val || 0);
    }, 0);
    const hayCorregidos = keys.some(k => verifOp[k] && verifOp[k].estado === 'corregido');
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const horaActual = pad(now.getHours()) + ':' + pad(now.getMinutes()) + ':' + pad(now.getSeconds());

    const payload = {
      ProcesoID: this.currentProcesoId || '',
      Auxiliar: operador,
      Ruta: parseInt(ruta, 10) || 0,
      PalletNum: parseInt(pnum, 10) || 0,
      TipoPallet: 'Auxiliar',
      NumProductos: keys.length,
      CajasPedidas: cajasPedidas,
      CajasReales: cajasReales,
      HoraInicio: this.verifInicioPallet[palletKeyId] || horaActual,
      HoraTermino: horaActual,
      Completo: 'Si',
      EstadoVerificacion: hayCorregidos ? 'Corregido' : 'Correcto',
      Verificador: this.nombreVerificadorActual(),
      HoraVerificacion: horaActual
    };
    try {
      const resp = await this.fetchConTimeout(Component.POST_RESUMEN_FLOW_URL, payload);
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      this.logMsg('✅ Resumen del pallet ' + ruta + '/' + pnum + ' guardado', 'ok');
    } catch (e) {
      this.logMsg('❌ Error guardando resumen del pallet ' + ruta + '/' + pnum + ': ' + e.message, 'err');
    }
  }

  async authLogin() {
    const u = this.state.authUsername.trim();
    const p = this.state.authPassword.trim();
    if (!u || !p) { this.setState({ authError: 'Completa usuario y contraseña.' }); return; }
    if (this.state.authLoading) return;
    this.setState({ authLoading: true, authError: '' });
    let data;
    try {
      const resp = await this.fetchConTimeout(Component.LOGIN_FLOW_URL, { usuario: u, clave: p });
      if (resp.status === 401) {
        this.setState({ authLoading: false, authError: 'Usuario o clave incorrectos.' });
        return;
      }
      if (!resp.ok) {
        this.setState({ authLoading: false, authError: 'No se pudo conectar con el servidor. Intenta nuevamente.' });
        return;
      }
      data = await resp.json();
    } catch (e) {
      this.setState({ authLoading: false, authError: 'No se pudo conectar con el servidor. Intenta nuevamente.' });
      return;
    }
    const rol = String(data.rol || '').trim().toLowerCase();
    const authUser = { username: data.usuario || u, nombre: data.nombre, apellido: data.apellido, role: rol };
    if (rol === 'admin' || rol === 'verificador') {
      this.setState({ authLoading: false, screen: 'roleChoice', authUser, authError: '' });
      this.loadProductos();
      this.loadUsuarios();
      this.loadUltimoProceso();
    } else {
      if (!this.lastResultado) await this.loadUltimoProceso();
      if (!this.lastResultado || !this.lastResultado.resumen) {
        this.setState({ authLoading: false, authError: 'No hay asignaciones disponibles. Por favor comunicarse con el encargado.' });
        return;
      }
      const nombreCompleto = ((data.nombre || '') + ' ' + (data.apellido || '')).trim().toLowerCase();
      const miAsignacion = this.lastResultado.resumen.find(a => a.auxiliar.trim().toLowerCase() === nombreCompleto || a.auxiliar.trim().toLowerCase() === u.toLowerCase());
      if (!miAsignacion) {
        this.setState({ authLoading: false, authError: 'No tienes asignación. Por favor comunicarse con el encargado.' });
        return;
      }
      this.setState({ authLoading: false, screen: 'app', tab: 'arme', authUser, authError: '' });
      this.mostrarArme(miAsignacion);
    }
  }

  logoutFull() {
    this.setState({ screen: 'splash', authUser: null, authUsername: '', authPassword: '', authError: '' });
  }

  updateProductoCell(idx, field, value) {
    this.setState(s => {
      const productos = s.productos.slice();
      let row = { ...productos[idx], [field]: value };
      if (field === 'Categoria') {
        if (value === 'Alimentos') { row.Grupo = 'Alimentos'; row.Volumen = 'Alimentos'; }
        else if (value === 'N/A') { row.Grupo = 'N/A'; row.Volumen = 'N/A'; }
        else { row.Grupo = ''; row.Volumen = ''; }
      }
      if (field === 'Categoria' || field === 'Grupo' || field === 'Volumen') {
        row = { ...row, ...Component.computeZona(row.Categoria, row.Grupo, row.Volumen) };
      }
      productos[idx] = row;
      return { productos };
    });
  }

  static VOLUMEN_OPTIONS = ['355 ML','340 ML','1250 ML','500 ML','1500 ML','3000 ML','400 ML','19350 ML','6000 ML','473 ML','750 ML','2000 ML','250 ML','2500 ML','600 ML','300 ML','1000 ML','330 ML','700 ML','625 ML','3785 ML','20000 ML','4000 ML','1600 ML','1200 ML','19500 ML','269 ML'];

  static computeZona(categoria, grupo, volumen) {
    if (categoria === 'Alimentos' || categoria === 'N/A') return { NameZone: 'Zona 4', ZoneCode: 'Z4', DescriptionZone: 'Higiene-Alimentos' };
    if (grupo === 'PRB' || grupo === 'GRB') return { NameZone: 'Zona 1', ZoneCode: 'Z1', DescriptionZone: 'PRB-GRB' };
    if (grupo === 'Lata') return { NameZone: 'Zona 2', ZoneCode: 'Z2', DescriptionZone: 'PET >600ML - LATA' };
    if (grupo === 'Funda' || grupo === 'Tetrapack') return { NameZone: 'Zona 4', ZoneCode: 'Z4', DescriptionZone: 'Higiene-Alimentos' };
    if (grupo === 'OW') return { NameZone: 'Zona 3', ZoneCode: 'Z3', DescriptionZone: 'PET<600ML - OW' };
    if (grupo === 'PET') {
      const ml = parseInt(String(volumen), 10) || 0;
      return ml >= 600 ? { NameZone: 'Zona 2', ZoneCode: 'Z2', DescriptionZone: 'PET >600ML - LATA' } : { NameZone: 'Zona 3', ZoneCode: 'Z3', DescriptionZone: 'PET<600ML - OW' };
    }
    return { NameZone: '', ZoneCode: '', DescriptionZone: '' };
  }

  openAddProducto() {
    this.setState({ addProductoOpen: true, newProducto: { Codigo: '', Nombre_Material: '', IsPicked: '', Categoria: '', Grupo: '', Volumen: '', UnitBox: '', BoxPallet: '' } });
  }
  closeAddProducto() {
    this.setState({ addProductoOpen: false });
  }
  updateNewProductoField(field, value) {
    this.setState(s => {
      const np = { ...s.newProducto, [field]: value };
      if (field === 'Categoria') {
        if (value === 'Alimentos') { np.Grupo = 'Alimentos'; np.Volumen = 'Alimentos'; }
        else if (value === 'N/A') { np.Grupo = 'N/A'; np.Volumen = 'N/A'; }
        else { np.Grupo = ''; np.Volumen = ''; }
      }
      return { newProducto: np };
    });
  }
  async confirmAddProducto() {
    const np = this.state.newProducto;
    if (!np.Codigo.trim() || !np.Nombre_Material.trim() || !np.IsPicked || !np.Categoria || !np.Grupo || !np.Volumen) { return; }
    if (!(Number(np.UnitBox) > 0) || !(Number(np.BoxPallet) > 0)) {
      this.setState({ productosError: 'UnitBox y BoxPallet son obligatorios y deben ser mayores a 0.' });
      return;
    }
    const zona = Component.computeZona(np.Categoria, np.Grupo, np.Volumen);
    const producto = { Codigo: np.Codigo, Nombre_Material: np.Nombre_Material, IsPicked: np.IsPicked, Categoria: np.Categoria, Grupo: np.Grupo, Volumen: np.Volumen, ...zona, UnitBox: np.UnitBox, BoxPallet: np.BoxPallet };
    this.setState({ productoSaving: true, productosError: '' });
    try {
      const payload = {
        Codigo: producto.Codigo,
        Nombre_Material: producto.Nombre_Material,
        IsPicked: producto.IsPicked === 'Si',
        Categoria: producto.Categoria,
        Grupo: producto.Grupo,
        Volumen: producto.Volumen,
        NameZone: producto.NameZone,
        ZoneCode: producto.ZoneCode,
        DescriptionZone: producto.DescriptionZone,
        UnitBox: Number(producto.UnitBox) || 0,
        BoxPallet: Number(producto.BoxPallet) || 0
      };
      const resp = await fetch(Component.POST_PRODUCTO_FLOW_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      let created = null;
      try { created = await resp.json(); } catch (e) { /* sin cuerpo de respuesta */ }
      const savedProducto = created && created.ItemInternalId ? { ...producto, ItemInternalId: created.ItemInternalId } : producto;
      this.setState(s => ({ productos: [...s.productos, savedProducto], addProductoOpen: false, productoSaving: false }));
      if (!savedProducto.ItemInternalId) await this.loadProductos();
    } catch (e) {
      this.setState({ productoSaving: false, productosError: 'No se pudo guardar el producto en la base. Intenta nuevamente.' });
    }
  }

  startProductoEdit(idx) {
    this.setState(s => ({ productoEditIdx: idx, productoEditBackup: { ...s.productos[idx] } }));
  }
  async saveProductoEdit() {
    const s = this.state;
    const row = s.productos[s.productoEditIdx];
    if (!row.Codigo.trim() || !row.Nombre_Material.trim() || !row.IsPicked || !row.Categoria || !row.Grupo || !row.Volumen) { return; }
    if (!(Number(row.UnitBox) > 0) || !(Number(row.BoxPallet) > 0)) {
      this.setState({ productosError: 'UnitBox y BoxPallet son obligatorios y deben ser mayores a 0.' });
      return;
    }
    if (!row.ItemInternalId) { this.setState({ productosError: 'Este producto no tiene ID de base de datos (recarga la lista antes de editarlo).' }); return; }
    this.setState({ productoSaving: true, productosError: '' });
    try {
      const payload = {
        ItemInternalId: row.ItemInternalId,
        Codigo: row.Codigo,
        Nombre_Material: row.Nombre_Material,
        IsPicked: row.IsPicked === 'Si',
        Categoria: row.Categoria,
        Grupo: row.Grupo,
        Volumen: row.Volumen,
        NameZone: row.NameZone,
        ZoneCode: row.ZoneCode,
        DescriptionZone: row.DescriptionZone,
        UnitBox: Number(row.UnitBox) || 0,
        BoxPallet: Number(row.BoxPallet) || 0
      };
      const resp = await fetch(Component.PUT_PRODUCTO_FLOW_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      this.setState({ productoEditIdx: null, productoEditBackup: null, productoSaving: false });
    } catch (e) {
      this.setState({ productoSaving: false, productosError: 'No se pudo guardar el cambio en la base. Intenta nuevamente.' });
    }
  }
  toggleProductoExpanded(idx) {
    this.setState(s => ({ productosExpanded: { ...s.productosExpanded, [idx]: !s.productosExpanded[idx] } }));
  }
  cancelProductoEdit() {
    this.setState(s => {
      if (s.productoEditIdx == null || !s.productoEditBackup) return { productoEditIdx: null, productoEditBackup: null };
      const productos = s.productos.slice();
      productos[s.productoEditIdx] = s.productoEditBackup;
      return { productos, productoEditIdx: null, productoEditBackup: null };
    });
  }
  startUsuarioEdit(idx) {
    this.setState(s => ({ usuarioEditIdx: idx, usuarioEditBackup: { ...s.usuarios[idx] } }));
  }
  async saveUsuarioEdit() {
    const s = this.state;
    const idx = s.usuarioEditIdx;
    const row = s.usuarios[idx];
    if (!row.nombre.trim() || !row.apellido.trim() || !row.usuario.trim() || !row.role) { return; }
    if (!/^\d{4}$/.test(String(row.password).trim())) { this.setState({ usuariosError: 'La contraseña debe ser 4 números.' }); return; }
    const dup = s.usuarios.some((u, i) => i !== idx && String(u.usuario || '').toLowerCase() === String(row.usuario).trim().toLowerCase());
    if (dup) { this.setState({ usuariosError: 'Ese usuario ya existe, elige otro.' }); return; }
    if (!row.UserInternalId) { this.setState({ usuariosError: 'Este usuario no tiene ID de base de datos (recarga la lista antes de editarlo).' }); return; }
    this.setState({ usuariosSaving: true, usuariosError: '' });
    try {
      const payload = { UserInternalId: row.UserInternalId, Nombre: row.nombre, Apellido: row.apellido, Usuario: row.usuario, Contraseña: Number(row.password), Rol: row.role };
      const resp = await fetch(Component.PUT_USUARIO_FLOW_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      this.setState({ usuarioEditIdx: null, usuarioEditBackup: null, usuariosSaving: false });
    } catch (e) {
      this.setState({ usuariosSaving: false, usuariosError: 'No se pudo guardar el cambio en la base. Intenta nuevamente.' });
    }
  }
  cancelUsuarioEdit() {
    this.setState(s => {
      if (s.usuarioEditIdx == null || !s.usuarioEditBackup) return { usuarioEditIdx: null, usuarioEditBackup: null };
      const usuarios = s.usuarios.slice();
      usuarios[s.usuarioEditIdx] = s.usuarioEditBackup;
      return { usuarios, usuarioEditIdx: null, usuarioEditBackup: null };
    });
  }
  requestDeleteProducto(idx) {
    const p = this.state.productos[idx];
    this.setState({ deleteConfirm: { idx, codigo: p.Codigo, nombre: p.Nombre_Material } });
  }
  cancelDeleteProducto() {
    this.setState({ deleteConfirm: null });
  }
  async confirmDeleteProducto() {
    const idx = this.state.deleteConfirm.idx;
    const row = this.state.productos[idx];
    if (!row.Codigo) {
      this.setState(s => ({ productos: s.productos.filter((_, i) => i !== idx), deleteConfirm: null }));
      return;
    }
    this.setState({ productoSaving: true, productosError: '' });
    try {
      const resp = await fetch(Component.DELETE_PRODUCTO_FLOW_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ Codigo: row.Codigo }) });
      if (!resp.ok) {
        let bodyText = '';
        try { bodyText = await resp.text(); } catch (e2) { /* ignore */ }
        console.error('DELETE_Producto fall\u00f3', resp.status, bodyText);
        throw new Error('HTTP ' + resp.status);
      }
      this.setState(s => ({ productos: s.productos.filter((_, i) => i !== idx), deleteConfirm: null, productoSaving: false }));
    } catch (e) {
      console.error('confirmDeleteProducto error', e);
      this.setState({ productoSaving: false, deleteConfirm: null, productosError: 'No se pudo eliminar el producto en la base. Intenta nuevamente.' });
    }
  }

  updateUsuarioCell(idx, field, value) {
    this.setState(s => {
      const usuarios = s.usuarios.slice();
      usuarios[idx] = { ...usuarios[idx], [field]: value };
      return { usuarios };
    });
  }
  static generateUsuarioLogin(nombre, apellido, usuarios, excludeIdx) {
    const first = (nombre.trim()[0] || '').toLowerCase();
    const lastWord = (apellido.trim().split(/\s+/)[0] || '').toLowerCase();
    const base = (first + lastWord).replace(/[^a-z0-9]/g, '');
    const taken = new Set(usuarios.filter((_, i) => i !== excludeIdx).map(u => String(u.usuario || '').toLowerCase()));
    if (!taken.has(base)) return base;
    let candidate;
    do { candidate = base + String(Math.floor(10 + Math.random() * 90)); } while (taken.has(candidate));
    return candidate;
  }
  static generateUsuarioPassword() {
    return String(Math.floor(1000 + Math.random() * 9000));
  }
  openAddUsuario() {
    this.setState({ addUsuarioOpen: true, newUsuarioForm: { nombre: '', apellido: '', role: 'auxiliar' } });
  }
  closeAddUsuario() {
    this.setState({ addUsuarioOpen: false });
  }
  updateNewUsuarioField(field, value) {
    this.setState(s => ({ newUsuarioForm: { ...s.newUsuarioForm, [field]: value } }));
  }
  async confirmAddUsuario() {
    const nf = this.state.newUsuarioForm;
    if (!nf.nombre.trim() || !nf.apellido.trim() || !nf.role) return;
    const usuario = Component.generateUsuarioLogin(nf.nombre, nf.apellido, this.state.usuarios, -1);
    const password = Component.generateUsuarioPassword();
    this.setState({ usuariosSaving: true, usuariosError: '' });
    try {
      const payload = { Nombre: nf.nombre, Apellido: nf.apellido, Usuario: usuario, Contraseña: Number(password), Rol: nf.role };
      const resp = await fetch(Component.POST_USUARIO_FLOW_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
      if (!resp.ok) {
        let bodyText = '';
        try { bodyText = await resp.text(); } catch (e2) { /* ignore */ }
        console.error('POST_Usuario fall\u00f3', resp.status, bodyText);
        throw new Error('HTTP ' + resp.status);
      }
      let created = null;
      try { created = await resp.json(); } catch (e) { /* sin cuerpo de respuesta */ }
      const savedId = (created && (created.UserInternalId ?? created.ItemInternalId)) || null;
      const nuevo = { UserInternalId: savedId, nombre: nf.nombre, apellido: nf.apellido, usuario, password, role: nf.role };
      this.setState(s => ({ usuarios: [...s.usuarios, nuevo], addUsuarioOpen: false, usuariosSaving: false }));
      if (!savedId) await this.loadUsuarios();
    } catch (e) {
      console.error('confirmAddUsuario error', e);
      this.setState({ usuariosSaving: false, usuariosError: 'No se pudo guardar el usuario en la base. Intenta nuevamente.' });
    }
  }
  removeUsuario(idx) {
    this.setState(s => ({ usuarios: s.usuarios.filter((_, i) => i !== idx) }));
  }
  requestDeleteUsuario(idx) {
    const u = this.state.usuarios[idx];
    this.setState({ deleteUsuarioConfirm: { idx, nombre: u.nombre, apellido: u.apellido, usuario: u.usuario } });
  }
  cancelDeleteUsuario() {
    this.setState({ deleteUsuarioConfirm: null });
  }
  async confirmDeleteUsuario() {
    const idx = this.state.deleteUsuarioConfirm.idx;
    const row = this.state.usuarios[idx];
    if (!row.usuario) {
      this.setState(s => ({ usuarios: s.usuarios.filter((_, i) => i !== idx), deleteUsuarioConfirm: null }));
      return;
    }
    this.setState({ usuariosSaving: true, usuariosError: '' });
    try {
      const resp = await fetch(Component.DELETE_USUARIO_FLOW_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ Usuario: row.usuario }) });
      if (!resp.ok) {
        let bodyText = '';
        try { bodyText = await resp.text(); } catch (e2) { /* ignore */ }
        console.error('DELETE_Usuario falló', resp.status, bodyText);
        throw new Error('HTTP ' + resp.status);
      }
      this.setState(s => ({ usuarios: s.usuarios.filter((_, i) => i !== idx), deleteUsuarioConfirm: null, usuariosSaving: false }));
    } catch (e) {
      console.error('confirmDeleteUsuario error', e);
      this.setState({ usuariosSaving: false, deleteUsuarioConfirm: null, usuariosError: 'No se pudo eliminar el usuario en la base. Intenta nuevamente.' });
    }
  }

  renderVals() {
    const s = this.state;
    const showPasswordField = true;
    const propickFile = s.files.propick;
    const masterFile = s.files.master;
    const auxValidos = s.auxRows.filter(r => r.Nombre.trim()).length > 0;
    const productosListos = s.productos.length > 0;
    const listo = propickFile && productosListos && auxValidos && s.pyodideReady && !s.procesarBusy;

    const auxiliaresAll = s.usuarios.filter(u => u.role === 'auxiliar').map(u => (u.nombre + ' ' + u.apellido).trim());
    const auxRowsView = s.auxRows.map((row, idx) => {
      const selectedElsewhere = new Set(s.auxRows.filter((_, i) => i !== idx).map(r => r.Nombre).filter(Boolean));
      return {
        nombre: row.Nombre,
        dlId: 'dl-aux-' + idx,
        options: auxiliaresAll.filter(n => !selectedElsewhere.has(n)).map(n => ({ n })),
        onNombreChange: (e) => this.updateAuxCell(idx, 'Nombre', e.target.value),
        onRemove: () => this.removeAuxRow(idx)
      };
    });

    const detalleAll = this.lastResultado ? this.lastResultado.detalle : [];
    const filtAux = s.filterArmeAuxiliar || 'Todos';
    const filtRuta = s.filterArmeRuta || 'Todas';
    const armeAuxOptions = [...new Set(detalleAll.map(f => f.auxiliar))].sort();
    const armeRutaOptions = [...new Set(detalleAll.map(f => f.ruta))].sort();
    const detalleFiltrado = detalleAll.filter(f => (filtAux === 'Todos' || f.auxiliar === filtAux) && (filtRuta === 'Todas' || String(f.ruta) === String(filtRuta)));
    const gruposMap = {};
    detalleFiltrado.forEach(f => {
      gruposMap[f.auxiliar] = gruposMap[f.auxiliar] || {};
      gruposMap[f.auxiliar][f.ruta] = gruposMap[f.auxiliar][f.ruta] || {};
      gruposMap[f.auxiliar][f.ruta][f.pallet] = gruposMap[f.auxiliar][f.ruta][f.pallet] || { tipo: f.tipo, zona: f.notas || f.zona, items: [] };
      gruposMap[f.auxiliar][f.ruta][f.pallet].items.push(f);
    });
    const armeAdminGrupos = Object.keys(gruposMap).sort().map(aux => ({
      auxiliar: aux,
      rutas: Object.keys(gruposMap[aux]).sort().map(ruta => ({
        ruta,
        onEliminarRuta: () => this.onEliminarRuta(aux, ruta),
        pallets: Object.keys(gruposMap[aux][ruta]).sort((a, b) => a - b).map(pnum => {
          const p = gruposMap[aux][ruta][pnum];
          const estadoMap = s.armadoPorOperador[aux] || {};
          return {
            pnum,
            tipo: p.tipo,
            zona: p.zona,
            onEliminarPallet: () => this.onEliminarPallet(aux, ruta, pnum),
            items: p.items.map(it => {
              const key = 'r' + ruta + '_p' + pnum + '_m' + it.material;
              const st = estadoMap[key];
              const estadoLabel = st ? (st.estado === 'confirmado' ? 'Agregado' : st.estado === 'problema' ? 'Colocado ' + st.cajasReal + ' de ' + it.cajas : 'Pendiente') : 'Pendiente';
              const estadoTagClass = st && st.estado === 'confirmado' ? 'tag-accent' : (st && st.estado === 'problema' ? 'pp-tag-warn' : 'tag-neutral');
              return { materialLabel: it.material + (it.nombre ? ' — ' + it.nombre : ''), cajas: it.cajas, estadoLabel, estadoTagClass };
            })
          };
        })
      }))
    }));

    const armeRutas = Object.keys(s.armeRutasRaw).sort().map(ruta => ({
      ruta,
      pallets: Object.keys(s.armeRutasRaw[ruta]).sort((a, b) => a - b).map(pnum => {
        const p = s.armeRutasRaw[ruta][pnum];
        return {
          pnum,
          tipo: p.tipo,
          zona: p.zona,
          tagClass: p.tipo === 'Montacargas' ? 'tag-outline' : 'tag-accent',
          items: p.items.map((it, idx) => {
            const key = 'r' + ruta + '_p' + pnum + '_m' + it.material;
            const st = s.itemStatus[key] || { estado: 'pendiente', cajasReal: it.cajas };
            return {
              key,
              materialLabel: it.material + (it.nombre ? ' — ' + it.nombre : ''),
              cajas: it.cajas,
              cajasReal: st.cajasReal,
              rowClass: st.estado === 'confirmado' ? 'confirmed' : st.estado === 'problema' ? 'problem' : '',
              isPendiente: st.estado === 'pendiente',
              isConfirmado: st.estado === 'confirmado',
              isProblema: st.estado === 'problema',
              showQtyFix: st.estado === 'ajustando',
              qtyDraft: s.qtyDrafts[key] || '0',
              onConfirm: () => this.confirmarItem(key),
              onMarcarProblema: () => this.marcarProblema(key),
              onQtyChange: (e) => this.setQtyDraft(key, e.target.value),
              onGuardarCantidad: () => this.guardarCantidadReal(key)
            };
          })
        };
      })
    }));

    const palletsParaVerificar = [];
    const palletsEnProceso = [];
    if (this.lastResultado) {
      const operadoresTodos = [...new Set(this.lastResultado.detalle.map(d => d.auxiliar))];
      operadoresTodos.forEach(operador => {
        const itemStatus = s.armadoPorOperador[operador] || {};
        const misFilas = this.lastResultado.detalle.filter(d => d.auxiliar === operador);
        const porRuta = {};
        misFilas.forEach(f => {
          if (!porRuta[f.ruta]) porRuta[f.ruta] = {};
          if (!porRuta[f.ruta][f.pallet]) porRuta[f.ruta][f.pallet] = { tipo: f.tipo, zona: f.notas || f.zona, items: [] };
          porRuta[f.ruta][f.pallet].items.push(f);
        });
        const entradasOperador = [];
        Object.keys(porRuta).sort().forEach(ruta => {
          Object.keys(porRuta[ruta]).sort((a, b) => a - b).forEach(pnum => {
            const p = porRuta[ruta][pnum];
            const keys = p.items.map((it, idx) => 'r' + ruta + '_p' + pnum + '_m' + it.material);
            const terminado = keys.every(k => itemStatus[k] && itemStatus[k].estado !== 'pendiente' && itemStatus[k].estado !== 'ajustando');
            if (!terminado) {
              const hechos = keys.filter(k => itemStatus[k] && itemStatus[k].estado !== 'pendiente' && itemStatus[k].estado !== 'ajustando').length;
              const pct = keys.length > 0 ? Math.round((hechos / keys.length) * 100) : 0;
              palletsEnProceso.push({ operador, ruta, pnum, zona: p.zona, hechos, total: keys.length, pct, pctWidth: pct + '%' });
              return;
            }
            const verifOp = s.verificacionRegistro[operador] || {};
            const items = p.items.map((it, idx) => {
              const key = keys[idx];
              const st = itemStatus[key];
              const v = verifOp[key] || { estado: 'pendiente' };
              const diferente = st.cajasReal !== st.cajasPlan;
              return {
                key,
                materialLabel: it.material + (it.nombre ? ' — ' + it.nombre : ''),
                cajasPlan: st.cajasPlan,
                cajasReal: st.cajasReal,
                cajasLabel: diferente ? ('Pedido: ' + st.cajasPlan + ' — Real: ' + st.cajasReal) : (st.cajasReal + ' cajas'),
                cajasTagClass: diferente ? 'pp-item-cajas-warn' : '',
                pendienteVerif: v.estado === 'pendiente',
                verifCorrecto: v.estado === 'correcto',
                verifCorregido: v.estado === 'corregido',
                cantidadCorregida: v.cantidad,
                showVerifQtyFix: v.estado === 'ajustando',
                verifQtyDraft: s.verifQtyDrafts[key] || String(st.cajasReal),
                onVerifCorrecto: () => this.marcarVerifCorrecto(operador, key),
                onVerifIncorrecto: () => this.marcarVerifIncorrecto(operador, key, st.cajasReal),
                onVerifQtyChange: (e) => this.setVerifQtyDraft(key, e.target.value),
                onGuardarVerifCantidad: () => this.guardarVerifCantidad(operador, key)
              };
            });
            entradasOperador.push({
              operador,
              ruta,
              pnum,
              zona: p.zona,
              incompleto: items.some(it => it.cajasReal !== it.cajasPlan),
              items
            });
          });
        });
        entradasOperador.forEach(e => palletsParaVerificar.push(e));
      });
    }
    palletsEnProceso.sort((a, b) => b.pct - a.pct);
    const operadoresVerifOptions = [...new Set([...palletsParaVerificar, ...palletsEnProceso].map(p => p.operador))];
    const rutasVerifOptions = [...new Set([...palletsParaVerificar, ...palletsEnProceso].map(p => p.ruta))];
    const palletsParaVerificarFiltrados = palletsParaVerificar.filter(p =>
      (!s.verifFiltroOperador || p.operador === s.verifFiltroOperador) &&
      (!s.verifFiltroRuta || p.ruta === s.verifFiltroRuta)
    );

    const auxSummary = s.auxSummaryRaw.map(aux => ({
      auxiliar: aux.auxiliar,
      totalCajas: aux.total_cajas,
      zonas: aux.zonas.map(z => ({ zona: z.zona, rutas: z.rutas }))
    }));

    return {
      isSplash: s.screen === 'splash',
      isLogin: s.screen === 'login',
      isRoleChoice: s.screen === 'roleChoice',
      showPasswordField,
      isApp: s.screen === 'app',
      onIngresar: () => this.setState({ screen: 'login' }),
      onBackToSplash: () => this.setState({ screen: 'splash', authUsername: '', authPassword: '', authError: '' }),
      authUsername: s.authUsername,
      authPassword: s.authPassword,
      onAuthUsernameChange: (e) => this.setState({ authUsername: e.target.value }),
      onAuthPasswordChange: (e) => this.setState({ authPassword: e.target.value }),
      onAuthLogin: () => this.authLogin(),
      hasAuthError: !!s.authError,
      authError: s.authError,
      authLoading: !!s.authLoading,
      authLoginLabel: s.authLoading ? 'Ingresando…' : 'Ingresar',
      onChooseArmado: () => { this.setState({ screen: 'app', tab: 'armado' }); if (!this.state.productos.length && !this.state.productosLoading) this.loadProductos(); if (!this.state.usuarios.length && !this.state.usuariosLoading) this.loadUsuarios(); },
      onChooseArme: () => this.setState({ screen: 'app', tab: 'arme' }),
      onChooseVerificacion: () => this.setState({ screen: 'app', tab: 'verificacion' }),
      palletsParaVerificar: palletsParaVerificarFiltrados,
      noPalletsParaVerificar: palletsParaVerificarFiltrados.length === 0,
      operadoresVerifOptions,
      rutasVerifOptions,
      verifFiltroOperador: s.verifFiltroOperador || '',
      verifFiltroRuta: s.verifFiltroRuta || '',
      onVerifFiltroOperadorChange: (e) => this.setState({ verifFiltroOperador: e.target.value }),
      onVerifFiltroRutaChange: (e) => this.setState({ verifFiltroRuta: e.target.value }),
      isVerifSubtabListos: (s.verifSubtab || 'listos') === 'listos',
      isVerifSubtabProceso: (s.verifSubtab || 'listos') === 'proceso',
      verifSubtabListosBtnClass: (s.verifSubtab || 'listos') === 'listos' ? 'btn btn-primary' : 'btn btn-secondary',
      verifSubtabProcesoBtnClass: (s.verifSubtab || 'listos') === 'proceso' ? 'btn btn-primary' : 'btn btn-secondary',
      onVerifSubtabListos: () => this.setState({ verifSubtab: 'listos' }),
      onVerifSubtabProceso: () => this.setState({ verifSubtab: 'proceso' }),
      listosCount: palletsParaVerificarFiltrados.length,
      enProcesoCount: palletsEnProceso.length,
      palletsEnProceso: palletsEnProceso.filter(p =>
        (!s.verifFiltroOperador || p.operador === s.verifFiltroOperador) &&
        (!s.verifFiltroRuta || p.ruta === s.verifFiltroRuta)
      ),
      noPalletsEnProceso: palletsEnProceso.filter(p =>
        (!s.verifFiltroOperador || p.operador === s.verifFiltroOperador) &&
        (!s.verifFiltroRuta || p.ruta === s.verifFiltroRuta)
      ).length === 0,
      onVolverInicio: () => this.logoutFull(),
      onBackToRoleChoice: () => this.setState({ screen: 'roleChoice' }),
      onInicioClick: () => this.setState(st => ({ screen: (st.authUser && (st.authUser.role === 'admin' || st.authUser.role === 'verificador')) ? 'roleChoice' : 'splash', authUser: (st.authUser && (st.authUser.role === 'admin' || st.authUser.role === 'verificador')) ? st.authUser : null, authUsername: '', authPassword: '', authError: '' })),
      isAdminRole: !!s.authUser && s.authUser.role === 'admin',
      isMenuRole: !!s.authUser && (s.authUser.role === 'admin' || s.authUser.role === 'verificador'),
      inicioBtnStyle: (s.authUser && s.authUser.role === 'admin') ? '' : 'margin-left:auto',
      isTabArmado: s.tab === 'armado',
      isTabArme: s.tab === 'arme',
      isTabVerificacion: s.tab === 'verificacion',
      isTabProductos: s.tab === 'productos',
      productosLoading: !!s.productosLoading,
      reloadProductosLabel: s.productosLoading ? 'Cargando…' : 'Recargar',
      hasProductosError: !!s.productosError,
      productosError: s.productosError,
      onReloadProductos: () => this.loadProductos(),
      isTabUsuarios: s.tab === 'usuarios',
      usuariosLoading: !!s.usuariosLoading,
      reloadUsuariosLabel: s.usuariosLoading ? 'Cargando…' : 'Recargar',
      hasUsuariosError: !!s.usuariosError,
      usuariosError: s.usuariosError,
      usuariosSaving: !!s.usuariosSaving,
      onReloadUsuarios: () => this.loadUsuarios(),
      onChooseProductos: () => { this.setState({ screen: 'app', tab: 'productos' }); this.loadProductos(); },
      onChooseUsuarios: () => { this.setState({ screen: 'app', tab: 'usuarios' }); this.loadUsuarios(); },
      productosView: s.productos.map((p, idx) => ({
        idx, Codigo: p.Codigo, Nombre_Material: p.Nombre_Material, IsPicked: p.IsPicked, Categoria: p.Categoria, Grupo: p.Grupo, Volumen: p.Volumen, NameZone: p.NameZone, ZoneCode: p.ZoneCode, isExpanded: !!s.productosExpanded[idx], onToggleExpand: () => this.toggleProductoExpanded(idx), expandIconClass: s.productosExpanded[idx] ? 'ph-duotone ph-caret-up' : 'ph-duotone ph-caret-down', DescriptionZone: p.DescriptionZone, UnitBox: p.UnitBox, BoxPallet: p.BoxPallet,
        onCodigoChange: (e) => this.updateProductoCell(idx, 'Codigo', e.target.value),
        onNombreChange: (e) => this.updateProductoCell(idx, 'Nombre_Material', e.target.value),
        onIsPickedChange: (e) => this.updateProductoCell(idx, 'IsPicked', e.target.value),
        onCategoriaChange: (e) => this.updateProductoCell(idx, 'Categoria', e.target.value),
        onGrupoChange: (e) => this.updateProductoCell(idx, 'Grupo', e.target.value),
        onVolumenChange: (e) => this.updateProductoCell(idx, 'Volumen', e.target.value),
        onNameZoneChange: (e) => this.updateProductoCell(idx, 'NameZone', e.target.value),
        onZoneCodeChange: (e) => this.updateProductoCell(idx, 'ZoneCode', e.target.value),
        onDescriptionZoneChange: (e) => this.updateProductoCell(idx, 'DescriptionZone', e.target.value),
        onUnitBoxChange: (e) => this.updateProductoCell(idx, 'UnitBox', e.target.value),
        onBoxPalletChange: (e) => this.updateProductoCell(idx, 'BoxPallet', e.target.value),
        isEditing: idx === s.productoEditIdx,
        isNotEditing: idx !== s.productoEditIdx,
        isEditCatBebidas: String(p.Categoria).toLowerCase() === 'bebidas',
        isEditCatFixed: String(p.Categoria).toLowerCase() !== 'bebidas',
        onStartEdit: () => this.startProductoEdit(idx),
        onSaveEdit: () => this.saveProductoEdit(),
        onCancelEdit: () => this.cancelProductoEdit(),
        onRemove: () => this.requestDeleteProducto(idx)
      })),
      onAddProducto: () => this.openAddProducto(),
      addProductoOpen: s.addProductoOpen,
      newProducto: s.newProducto,
      onNewCodigoChange: (e) => this.updateNewProductoField('Codigo', e.target.value),
      onNewNombreChange: (e) => this.updateNewProductoField('Nombre_Material', e.target.value),
      onNewIsPickedChange: (e) => this.updateNewProductoField('IsPicked', e.target.value),
      onNewCategoriaChange: (e) => this.updateNewProductoField('Categoria', e.target.value),
      onNewGrupoChange: (e) => this.updateNewProductoField('Grupo', e.target.value),
      onNewVolumenChange: (e) => this.updateNewProductoField('Volumen', e.target.value),
      onNewUnitBoxChange: (e) => this.updateNewProductoField('UnitBox', e.target.value),
      onNewBoxPalletChange: (e) => this.updateNewProductoField('BoxPallet', e.target.value),
      isCatBebidas: s.newProducto.Categoria === 'Bebidas',
      volumenOptions: Component.VOLUMEN_OPTIONS,
      onCancelAddProducto: () => this.closeAddProducto(),
      onConfirmAddProducto: () => this.confirmAddProducto(),
      productoSaving: !!s.productoSaving,
      addProductoLabel: s.productoSaving ? 'Guardando…' : 'Agregar',
      deleteConfirmOpen: !!s.deleteConfirm,
      deleteConfirmCodigo: s.deleteConfirm ? s.deleteConfirm.codigo : '',
      deleteConfirmNombre: s.deleteConfirm ? s.deleteConfirm.nombre : '',
      onCancelDelete: () => this.cancelDeleteProducto(),
      onConfirmDelete: () => this.confirmDeleteProducto(),
      usuariosView: s.usuarios.map((u, idx) => ({
        idx, nombre: u.nombre, apellido: u.apellido, usuario: u.usuario, password: u.password, role: u.role,
        onNombreChange: (e) => this.updateUsuarioCell(idx, 'nombre', e.target.value),
        onApellidoChange: (e) => this.updateUsuarioCell(idx, 'apellido', e.target.value),
        onUsuarioChange: (e) => this.updateUsuarioCell(idx, 'usuario', e.target.value),
        onPasswordChange: (e) => this.updateUsuarioCell(idx, 'password', e.target.value),
        onRoleChange: (e) => this.updateUsuarioCell(idx, 'role', e.target.value),
        isEditing: idx === s.usuarioEditIdx,
        isNotEditing: idx !== s.usuarioEditIdx,
        usuarioSaving: !!s.usuariosSaving,
        onStartEdit: () => this.startUsuarioEdit(idx),
        onSaveEdit: () => this.saveUsuarioEdit(),
        onCancelEdit: () => this.cancelUsuarioEdit(),
        onRemove: () => this.requestDeleteUsuario(idx)
      })),
      deleteUsuarioConfirmOpen: !!s.deleteUsuarioConfirm,
      deleteUsuarioConfirmNombre: s.deleteUsuarioConfirm ? s.deleteUsuarioConfirm.nombre : '',
      deleteUsuarioConfirmApellido: s.deleteUsuarioConfirm ? s.deleteUsuarioConfirm.apellido : '',
      deleteUsuarioConfirmUsuario: s.deleteUsuarioConfirm ? s.deleteUsuarioConfirm.usuario : '',
      onCancelDeleteUsuario: () => this.cancelDeleteUsuario(),
      onConfirmDeleteUsuario: () => this.confirmDeleteUsuario(),
      onAddUsuario: () => this.openAddUsuario(),
      addUsuarioOpen: !!s.addUsuarioOpen,
      newUsuarioForm: s.newUsuarioForm,
      addUsuarioLabel: s.usuariosSaving ? 'Guardando…' : 'Agregar',
      onNewUsuarioNombreChange: (e) => this.updateNewUsuarioField('nombre', e.target.value),
      onNewUsuarioApellidoChange: (e) => this.updateNewUsuarioField('apellido', e.target.value),
      onNewUsuarioRoleChange: (e) => this.updateNewUsuarioField('role', e.target.value),
      onConfirmAddUsuario: () => this.confirmAddUsuario(),
      onCancelAddUsuario: () => this.closeAddUsuario(),

      armeNombre: s.armeNombre,
      armeResumen: s.armeResumen,
      armeRutas,
      isAuxiliarSelf: !!(s.authUser && s.authUser.role === 'auxiliar'),
      isConsulta: !(s.authUser && s.authUser.role === 'auxiliar'),
      filterArmeAuxiliar: s.filterArmeAuxiliar || 'Todos',
      filterArmeRuta: s.filterArmeRuta || 'Todas',
      onFilterArmeAuxiliarChange: (e) => this.setState({ filterArmeAuxiliar: e.target.value }),
      onFilterArmeRutaChange: (e) => this.setState({ filterArmeRuta: e.target.value }),
      armeAuxOptions,
      armeRutaOptions,
      armeAdminGrupos,
      noAsignaciones: armeAdminGrupos.length === 0,
      onEliminarTodas: () => this.onEliminarTodas(),
      armeDeleteConfirmOpen: !!s.armeDeleteConfirm,
      armeDeleteConfirmTexto: s.armeDeleteConfirm ? s.armeDeleteConfirm.texto : '',
      onCancelarEliminarAsignacion: () => this.cancelarEliminarAsignacion(),
      onConfirmarEliminarAsignacion: () => this.confirmarEliminarAsignacion(),

      propickFilledClass: propickFile ? 'filled' : '',
      propickFileLabel: propickFile ? '✓ ' + propickFile.name + ' (' + (propickFile.size / 1024).toFixed(1) + ' KB)' : '',
      masterFilledClass: masterFile ? 'filled' : '',
      masterFileLabel: masterFile ? '✓ ' + masterFile.name + ' (' + (masterFile.size / 1024).toFixed(1) + ' KB)' : '',
      onClickPropick: () => document.getElementById('pp-file-propick').click(),
      onClickMaster: () => document.getElementById('pp-file-master').click(),
      onPropickChange: (e) => this.onFileSelected('propick', e.target.files[0]),
      onMasterChange: (e) => this.onFileSelected('master', e.target.files[0]),

      auxRowsView,
      auxCountLabel: s.auxRows.length + ' operador' + (s.auxRows.length === 1 ? '' : 'es'),
      onAddAuxRow: () => this.addAuxRow(),

      procesarDisabled: !listo,
      onProcesar: () => this.procesar(),
      onCargarEjemplo: () => this.cargarDatosEjemplo(),
      hasProcesoResultado: !!s.procesoResultado,
      procesoResultadoTitulo: s.procesoResultado ? s.procesoResultado.titulo : '',
      procesoResultadoDetalle: s.procesoResultado ? s.procesoResultado.detalle : '',
      procesoResultadoColor: s.procesoResultado && s.procesoResultado.ok ? 'var(--color-accent-600)' : 'var(--color-accent-2-600)',
      procesoResultadoIcon: s.procesoResultado && s.procesoResultado.ok ? 'ph-duotone ph-check-circle' : 'ph-duotone ph-warning-circle',
      showStatusRow: !s.pyodideReady || s.procesarBusy,
      statusText: s.procesarBusy ? 'Procesando...' : s.statusText,
      consoleLines: s.consoleLines,

      resultsVisible: s.resultsVisible,
      stats: s.stats,
      sqlSyncOk: s.sqlSyncStatus === 'ok',
      sqlSyncPending: s.sqlSyncStatus === 'pending',
      sqlSyncFailed: s.sqlSyncStatus === 'error' || s.sqlSyncStatus === 'partial',
      sqlSyncDetalle: s.sqlSyncDetalle,

      auxSummary,

      missingVisible: s.missingVisible,
      missingList: s.missingList
    };
  }
}

