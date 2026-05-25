import pandas as pd, re

path_may = r'C:\proyectos\AVANCE_MOVISTAR\Archivos_Avance\AVANCE_2026-05-17.xlsx'
path_abr = r'C:\proyectos\AVANCE_MOVISTAR\Archivos_Avance\AVANCE_2026-04-30.xlsx'

# ── RT mayo: VDD2 vs hoja cruda ───────────────────────────────────────────────
vdd2 = pd.read_excel(path_may, sheet_name='VDD2')
rt_cols = [c for c in vdd2.columns if re.match(r'RT_D\d+', c)]
dias_rt = sorted([int(re.search(r'D(\d+)', c).group(1)) for c in rt_cols])
total_vdd2 = vdd2[rt_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum().sum()
print(f"VDD2 dias RT disponibles: {dias_rt}")
print(f"VDD2 RT total (suma): {int(total_vdd2)}")

rt_raw = pd.read_excel(path_may, sheet_name='RT', usecols=['Fecha_Registro', 'zonal'])
rt_raw['fecha'] = pd.to_datetime(rt_raw['Fecha_Registro'], errors='coerce')
rt_raw['mes'] = rt_raw['fecha'].dt.month
rt_raw['dia'] = rt_raw['fecha'].dt.day
may = rt_raw[rt_raw['mes'] == 5]
print(f"RT cruda hoja RT - total mayo: {len(may)}")
print(f"RT cruda dias 1-17: {len(may[may['dia']<=17])}")
print(f"RT cruda todos dias (distribucion): {may['dia'].value_counts().sort_index().to_dict()}")

# ── ALTAS mayo: VDD2 vs hoja cruda ───────────────────────────────────────────
alt_cols = [c for c in vdd2.columns if re.match(r'ALT_D\d+', c)]
dias_alt = sorted([int(re.search(r'D(\d+)', c).group(1)) for c in alt_cols])
total_alt_vdd2 = vdd2[alt_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum().sum()
print(f"\nVDD2 dias ALT disponibles: {dias_alt}")
print(f"VDD2 ALT total (suma): {int(total_alt_vdd2)}")

alt_raw = pd.read_excel(path_may, sheet_name='ALTAS', usecols=['Fecha_Alta', 'zonal'])
alt_raw['fecha'] = pd.to_datetime(alt_raw['Fecha_Alta'], errors='coerce')
alt_raw['mes'] = alt_raw['fecha'].dt.month
alt_raw['dia'] = alt_raw['fecha'].dt.day
may_alt = alt_raw[alt_raw['mes'] == 5]
print(f"ALTAS cruda - total mayo: {len(may_alt)}")
print(f"ALTAS cruda dias 1-17: {len(may_alt[may_alt['dia']<=17])}")
print(f"ALTAS cruda distribucion dias: {may_alt['dia'].value_counts().sort_index().to_dict()}")

# ── Verificar si AVANCE_17 tiene mas dias que el 14 ──────────────────────────
print(f"\nMax dia RT en hoja RT: {int(may['dia'].max())}")
print(f"Max dia ALT en hoja ALTAS: {int(may_alt['dia'].max())}")
