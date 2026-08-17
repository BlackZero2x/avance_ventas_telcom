"""
Módulo: Cuadro resumen diario por zonal para jefes zonales.
Complementa a cuadro_resumen_sup_process.py: en vez de una fila por vendedor,
genera una tabla con una fila por día (formato "ddd dd/mm") y columnas RT, RUS, ALTAS.
Se envía junto con las demás tablas (mismo formato visual, mismos destinos).
"""

import sys
import os
import time
import pandas as pd
from datetime import datetime
import tempfile

sys.path.insert(0, r"C:\proyectos\AVANCE_MOVISTAR\whatsapp_server")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wa_client import WhatsAppClient
from cuadro_resumen_sup_process import _calcular_zonal2, _obtener_fecha_maxima

_DIAS_ES = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom']


def _formatear_dia(fecha):
    """Formatea una fecha como 'ddd dd/mm' en español (ej. 'Lun 03/08')."""
    return f"{_DIAS_ES[fecha.weekday()]} {fecha.strftime('%d/%m')}"


def _construir_tabla_diaria_zonal2(df_rt, df_altas, zonal2_vals):
    """
    Construye tabla con una fila por día y columnas RT, RUS, ALTAS,
    agregando todos los vendedores de las zonales indicadas.

    Args:
        df_rt: DataFrame de hoja RT (para RT y RUS por día, vía Fecha_Registro)
        df_altas: DataFrame de hoja ALTAS (para ALTAS por día, vía Fecha_Alta)
        zonal2_vals: lista de valores zonal2 a incluir (ej. ['TACNA', 'ILO'])

    Returns:
        DataFrame con columnas: DIA (fecha), DIA_LABEL, RT, RUS, ALTAS
        más una fila final "TOTAL".
    """
    df_rt = df_rt.copy()
    df_altas = df_altas.copy()
    df_rt['zonal2'] = df_rt['zonal'].apply(_calcular_zonal2)
    df_altas['zonal2'] = df_altas['zonal'].apply(_calcular_zonal2)

    rt_filt = df_rt[df_rt['zonal2'].isin(zonal2_vals)].copy()
    altas_filt = df_altas[df_altas['zonal2'].isin(zonal2_vals)].copy()

    rt_filt['DIA'] = pd.to_datetime(rt_filt['Fecha_Registro'], errors='coerce').dt.date
    altas_filt['DIA'] = pd.to_datetime(altas_filt['Fecha_Alta'], errors='coerce').dt.date

    rt_filt = rt_filt.dropna(subset=['DIA'])
    altas_filt = altas_filt.dropna(subset=['DIA'])

    rt_por_dia = rt_filt.groupby('DIA').size().rename('RT')

    if 'Flag_Registro_Unico' in rt_filt.columns:
        rus_por_dia = rt_filt.groupby('DIA')['Flag_Registro_Unico'].apply(
            lambda s: pd.to_numeric(s, errors='coerce').fillna(0).sum()
        ).rename('RUS')
    else:
        rus_por_dia = pd.Series(dtype=int, name='RUS')

    altas_por_dia = altas_filt.groupby('DIA').size().rename('ALTAS')

    if 'Scoring' in altas_filt.columns:
        regular_por_dia = altas_filt[altas_filt['Scoring'] == 'REGULAR'].groupby('DIA').size().rename('REGULAR')
        flex_por_dia = altas_filt[altas_filt['Scoring'] == 'FLEX'].groupby('DIA').size().rename('FLEX')
    else:
        regular_por_dia = pd.Series(dtype=int, name='REGULAR')
        flex_por_dia = pd.Series(dtype=int, name='FLEX')

    dias = sorted(set(rt_por_dia.index) | set(rus_por_dia.index) | set(altas_por_dia.index))

    tabla = pd.DataFrame({'DIA': dias})
    tabla = tabla.merge(rt_por_dia, on='DIA', how='left')
    tabla = tabla.merge(rus_por_dia, on='DIA', how='left')
    tabla = tabla.merge(altas_por_dia, on='DIA', how='left')
    tabla = tabla.merge(regular_por_dia, on='DIA', how='left')
    tabla = tabla.merge(flex_por_dia, on='DIA', how='left')

    for col in ['RT', 'RUS', 'ALTAS', 'REGULAR', 'FLEX']:
        tabla[col] = tabla[col].fillna(0).astype(int)

    tabla = tabla.sort_values('DIA').reset_index(drop=True)
    tabla['DIA_LABEL'] = tabla['DIA'].apply(_formatear_dia)

    total_row = pd.DataFrame({
        'DIA': [None],
        'DIA_LABEL': ['TOTAL'],
        'RT': [tabla['RT'].sum()],
        'RUS': [tabla['RUS'].sum()],
        'ALTAS': [tabla['ALTAS'].sum()],
        'REGULAR': [tabla['REGULAR'].sum()],
        'FLEX': [tabla['FLEX'].sum()],
    })
    tabla_final = pd.concat(
        [tabla[['DIA', 'DIA_LABEL', 'RT', 'RUS', 'ALTAS', 'REGULAR', 'FLEX']], total_row],
        ignore_index=True
    )

    return tabla_final


def _capturar_imagen_tabla_diaria(tabla_df, titulo="Cuadro", subtitulo=None):
    """
    Renderiza la tabla diaria (DIA_LABEL, RT, RUS, ALTAS) con el mismo estilo visual
    que las tablas por supervisor (Aptos Narrow, encabezado verde agua, fila TOTAL resaltada).
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        print("[WARN] PIL no disponible")
        return None

    temp_png = os.path.join(
        tempfile.gettempdir(),
        f"tabla_diaria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    )

    try:
        user_fonts_dir = os.path.join(
            os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Windows', 'Fonts'
        )
        aptos_narrow_candidates = [
            os.path.join(user_fonts_dir, "Aptos-Narrow.ttf"),
            "C:\\Windows\\Fonts\\aptos-narrow.ttf",
            "C:\\Windows\\Fonts\\Aptos-Narrow.ttf",
            "C:\\Windows\\Fonts\\AptosNarrow.ttf",
            "C:\\Windows\\Fonts\\AptosNarrow-Regular.ttf",
            "C:\\Windows\\Fonts\\aptos-narrow-regular.ttf",
        ]
        aptos_narrow_bold_candidates = [
            os.path.join(user_fonts_dir, "Aptos-Narrow-Bold.ttf"),
            "C:\\Windows\\Fonts\\aptos-narrow-bold.ttf",
            "C:\\Windows\\Fonts\\Aptos-Narrow-Bold.ttf",
            "C:\\Windows\\Fonts\\AptosNarrow-Bold.ttf",
            "C:\\Windows\\Fonts\\aptos-narrowb.ttf",
        ]

        def _cargar_primera(candidatos, size):
            for ruta in candidatos:
                try:
                    return ImageFont.truetype(ruta, size)
                except:
                    continue
            return None

        font_normal = _cargar_primera(aptos_narrow_candidates, 13)
        font_header = _cargar_primera(aptos_narrow_bold_candidates, 13) or font_normal
        font_title = _cargar_primera(aptos_narrow_bold_candidates, 16) or _cargar_primera(aptos_narrow_candidates, 16)
        font_subtitulo = _cargar_primera(aptos_narrow_bold_candidates, 15) or font_header

        if font_normal is None or font_header is None or font_title is None:
            try:
                font_header = ImageFont.truetype("C:\\Windows\\Fonts\\arialbd.ttf", 13)
                font_normal = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 13)
                font_title = ImageFont.truetype("C:\\Windows\\Fonts\\arialbd.ttf", 16)
                font_subtitulo = ImageFont.truetype("C:\\Windows\\Fonts\\arialbd.ttf", 15)
            except:
                font_header = font_normal = font_title = font_subtitulo = ImageFont.load_default()

        margin = 10
        padding_x = 4
        cell_height = int(18 * 0.9 * 0.9 * 0.9 * 0.85)
        header_height = int(20 * 0.9 * 0.9 * 0.9 * 0.85)
        title_height = int(22 * 0.9 * 0.9 * 0.9) + 4
        subtitulo_height = (int(18 * 0.9 * 0.9 * 0.9) + 4) if subtitulo else 0
        col_spacing = 0
        scale_factor = 2
        gap_entre_tablas = 14  # separación vertical entre la tabla diaria y la de REGULAR/FLEX
        subtitulo2_height = int(18 * 0.9 * 0.9 * 0.9) + 4

        headers = ['DÍA', 'RT', 'RUS', 'ALTAS']
        col_mapping = {'DÍA': 'DIA_LABEL', 'RT': 'RT', 'RUS': 'RUS', 'ALTAS': 'ALTAS'}

        # Tabla 2: resumen ALTAS por tipo (REGULAR / FLEX) — solo fila de totales
        total_regular = int(tabla_df.loc[tabla_df['DIA_LABEL'] == 'TOTAL', 'REGULAR'].iloc[0]) \
            if 'REGULAR' in tabla_df.columns else 0
        total_flex = int(tabla_df.loc[tabla_df['DIA_LABEL'] == 'TOTAL', 'FLEX'].iloc[0]) \
            if 'FLEX' in tabla_df.columns else 0
        headers2 = ['REGULAR', 'FLEX']
        tabla2_valores = {'REGULAR': total_regular, 'FLEX': total_flex}

        col_widths = {}
        temp_img = Image.new('RGB', (1, 1))
        temp_draw = ImageDraw.Draw(temp_img)

        for display_header in headers:
            bbox = temp_draw.textbbox((0, 0), display_header, font=font_header)
            header_width = bbox[2] - bbox[0] + padding_x * 2

            if display_header in ['RT', 'RUS', 'ALTAS']:
                # Ancho forzado (20% más angosto que el original de 40px): ignora el header aunque se recorte
                col_widths[display_header] = 32
                continue
            else:
                max_width = header_width
                internal_col = col_mapping[display_header]
                for row_val in tabla_df[internal_col]:
                    val_str = str(row_val)
                    if val_str.strip():
                        bbox = temp_draw.textbbox((0, 0), val_str, font=font_normal)
                        val_width = bbox[2] - bbox[0] + padding_x * 2
                        max_width = max(max_width, val_width)
                data_width = min(max_width, 120)

            col_widths[display_header] = max(header_width, data_width)
            if display_header == 'DÍA':
                # Reducida 20% igual que RT/RUS/ALTAS
                col_widths[display_header] = int(col_widths[display_header] * 0.8)

        col_widths2 = {}
        for display_header in headers2:
            bbox = temp_draw.textbbox((0, 0), display_header, font=font_header)
            header_width = bbox[2] - bbox[0] + padding_x * 2
            col_widths2[display_header] = max(header_width, 50)

        total_width = sum(col_widths.values()) + len(col_widths) * col_spacing + margin * 2
        total_width2 = sum(col_widths2.values()) + len(col_widths2) * col_spacing + margin * 2
        total_width = max(total_width, total_width2)

        num_rows = len(tabla_df)
        img_height = (
            title_height + subtitulo_height + header_height + (num_rows * cell_height)
            + gap_entre_tablas + subtitulo2_height + header_height + cell_height
            + margin * 2
        )

        img_width_scaled = int(total_width * scale_factor)
        img_height_scaled = int(img_height * scale_factor)
        img = Image.new('RGB', (img_width_scaled, img_height_scaled), color='white')
        draw = ImageDraw.Draw(img)

        def scale(val):
            return int(val * scale_factor)

        draw.text((scale(margin), scale(margin)), titulo, fill='black', font=font_title)

        if subtitulo:
            draw.text((scale(margin), scale(margin + title_height)), subtitulo, fill='#0000FF', font=font_subtitulo)

        y = scale(margin + title_height + subtitulo_height)
        x = scale(margin)

        for display_header in headers:
            col_width = scale(col_widths[display_header])
            draw.rectangle([x, y, x + col_width, y + scale(header_height)],
                          fill='#B2F7EF', outline='black', width=2)
            bbox = draw.textbbox((0, 0), display_header, font=font_header)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            text_x = x + (col_width - text_width) // 2
            text_y = y + (scale(header_height) - text_height) // 2 - bbox[1]
            draw.text((text_x, text_y), display_header, fill='black', font=font_header)
            x += col_width + scale(col_spacing)

        y = scale(margin + title_height + subtitulo_height + header_height)
        for row_idx, row in tabla_df.iterrows():
            is_total = str(row.get('DIA_LABEL', '')) == 'TOTAL'

            if is_total:
                bg_color = '#E8E8E8'
                cell_font = font_header
            else:
                bg_color = 'white'
                cell_font = font_normal

            x = scale(margin)
            for display_header in headers:
                col_width = scale(col_widths[display_header])
                internal_col = col_mapping[display_header]

                draw.rectangle([x, y, x + col_width, y + scale(cell_height)],
                              fill=bg_color, outline='#666666', width=1)

                val = row.get(internal_col, '')
                val_str = str(val)

                bbox = draw.textbbox((0, 0), val_str, font=cell_font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x + (col_width - text_width) // 2
                text_y = y + (scale(cell_height) - text_height) // 2 - bbox[1]

                draw.text((text_x, text_y), val_str, fill='black', font=cell_font)
                x += col_width + scale(col_spacing)

            y += scale(cell_height)

        table_top = scale(margin + title_height + subtitulo_height)
        table_left = scale(margin)
        table_right = scale(margin) + sum(scale(w) for w in col_widths.values())
        table_bottom = y
        thick_border_width = scale(1)
        draw.rectangle([table_left, table_top, table_right, table_bottom],
                      outline='black', width=thick_border_width)

        # --- Tabla 2: resumen ALTAS REGULAR / FLEX ---
        y += scale(gap_entre_tablas)

        draw.text((scale(margin), y), "Resumen ALTAS por tipo", fill='#0000FF', font=font_subtitulo)
        y += scale(subtitulo2_height)

        tabla2_top = y
        x = scale(margin)
        for display_header in headers2:
            col_width = scale(col_widths2[display_header])
            draw.rectangle([x, y, x + col_width, y + scale(header_height)],
                          fill='#B2F7EF', outline='black', width=2)
            bbox = draw.textbbox((0, 0), display_header, font=font_header)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            text_x = x + (col_width - text_width) // 2
            text_y = y + (scale(header_height) - text_height) // 2 - bbox[1]
            draw.text((text_x, text_y), display_header, fill='black', font=font_header)
            x += col_width + scale(col_spacing)

        y += scale(header_height)
        x = scale(margin)
        for display_header in headers2:
            col_width = scale(col_widths2[display_header])
            draw.rectangle([x, y, x + col_width, y + scale(cell_height)],
                          fill='#E8E8E8', outline='#666666', width=1)
            val_str = str(tabla2_valores[display_header])
            bbox = draw.textbbox((0, 0), val_str, font=font_header)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            text_x = x + (col_width - text_width) // 2
            text_y = y + (scale(cell_height) - text_height) // 2 - bbox[1]
            draw.text((text_x, text_y), val_str, fill='black', font=font_header)
            x += col_width + scale(col_spacing)
        y += scale(cell_height)

        tabla2_left = scale(margin)
        tabla2_right = scale(margin) + sum(scale(w) for w in col_widths2.values())
        draw.rectangle([tabla2_left, tabla2_top, tabla2_right, y],
                      outline='black', width=thick_border_width)

        img.save(temp_png, format='PNG')
        print(f"[INFO] Imagen diaria creada: {temp_png}")
        return temp_png

    except Exception as e:
        print(f"[ERROR] PIL rendering (diario) falló: {e}")
        import traceback
        traceback.print_exc()
        return None


def procesar_cuadro_resumen_sup_diario(avance_path, grupos=None):
    """
    Procesa cuadro resumen diario por zonal y envía imágenes por WhatsApp.
    Usa los mismos grupos/destinos que cuadro_resumen_sup_process.py por defecto.

    Args:
        avance_path: ruta al archivo AVANCE_{fecha}.xlsx
        grupos: dict opcional de grupos (mismo formato que cuadro_resumen_sup_process);
                si no se pasa, usa el set por defecto (FRANZ/LETICIA).

    Returns:
        bool: True si al menos una imagen se envió exitosamente, False si todas fallaron.
    """
    if not avance_path or not os.path.exists(avance_path):
        print(f"[CUADRO RESUMEN SUP DIARIO] Archivo no encontrado: {avance_path}")
        return False

    print("[CUADRO RESUMEN SUP DIARIO] Iniciando...")

    try:
        df_rt = pd.read_excel(avance_path, sheet_name='RT')
        df_altas = pd.read_excel(avance_path, sheet_name='ALTAS')

        max_fecha = _obtener_fecha_maxima(df_rt)
        fecha_str = max_fecha.strftime("%d/%m/%Y") if max_fecha else datetime.now().strftime("%d/%m/%Y")

        if grupos is None:
            grupos = {
                'FRANZ_TACNA': {'zonales': ['TACNA'], 'numero': '51933540190@c.us', 'titulo': 'TACNA'},
                'FRANZ_ILO': {'zonales': ['ILO'], 'numero': '51933540190@c.us', 'titulo': 'ILO'},
                'LETICIA_TRUJILLO': {'zonales': ['TRUJILLO'], 'numero': '51952099313@c.us', 'titulo': 'TRUJILLO'},
                'LETICIA_CHIMBOTE': {'zonales': ['CHIMBOTE'], 'numero': '51952099313@c.us', 'titulo': 'CHIMBOTE'},
                'LETICIA_HUARAZ': {'zonales': ['HUARAZ'], 'numero': '51952099313@c.us', 'titulo': 'HUARAZ'},
            }

        wa = WhatsAppClient()
        envios_ok = 0
        total_intentos = 0

        for jefe, config in grupos.items():
            zonales = config['zonales']
            destino = config['numero']
            titulo = config['titulo']

            print(f"  Procesando (diario) {jefe} ({titulo})...")
            total_intentos += 1

            try:
                tabla = _construir_tabla_diaria_zonal2(df_rt, df_altas, zonales)

                if len(tabla) == 0:
                    print(f"    → Sin datos diarios para {titulo}")
                    continue

                png_path = _capturar_imagen_tabla_diaria(
                    tabla,
                    titulo=f"{titulo} - Detalle diario al {fecha_str}",
                    subtitulo="RT / RUS / ALTAS por día"
                )

                if png_path and os.path.exists(png_path):
                    mensaje = f"Detalle diario de {titulo} al {fecha_str}"
                    wa.send_image(destino, png_path, caption=mensaje)
                    print(f"    [OK] Enviado (diario) a {jefe}")
                    envios_ok += 1
                    try:
                        os.remove(png_path)
                    except:
                        pass
                else:
                    print(f"    [ERROR] Falló captura de imagen diaria para {jefe}")

                time.sleep(2)

            except Exception as e:
                print(f"    [ERROR] Procesando diario {jefe}: {e}")

        print(f"[CUADRO RESUMEN SUP DIARIO] Completado: {envios_ok}/{total_intentos} envíos exitosos")
        return envios_ok > 0

    except Exception as e:
        print(f"[ERROR] Cuadro resumen sup diario: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    avance_file = r"C:\proyectos\AVANCE_MOVISTAR\Archivos_Avance\AVANCE_2026-08-16.xlsx"
    procesar_cuadro_resumen_sup_diario(avance_file)
