"""
Módulo: Cuadro resumen por supervisor para jefes zonales.
Genera tablas agrupadas por zonal2:
  - FRANZ: TACNA + ILO (1 cuadro)
  - LETICIA: TRUJILLO, CHIMBOTE, HUARAZ (3 cuadros separados)
Captura imagen de cada tabla y envía por WhatsApp.
"""

import sys
import os
import time
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
import tempfile

sys.path.insert(0, r"C:\proyectos\AVANCE_MOVISTAR\whatsapp_server")
from wa_client import WhatsAppClient

def _calcular_zonal2(zonal):
    """Calcula zonal2 según regla: LIMA si empieza con 'LIMA', sino igual a zonal."""
    if pd.isna(zonal):
        return zonal
    zonal_str = str(zonal).strip()
    return "LIMA" if zonal_str.startswith("LIMA") else zonal_str

def _obtener_fecha_maxima(df):
    """Extrae fecha máxima del campo Fecha_de_alta o Fecha_Alta."""
    col_fecha = None
    for c in df.columns:
        if c.lower() in ['fecha_de_alta', 'fecha_alta']:
            col_fecha = c
            break

    if col_fecha is None:
        return None

    max_fecha = pd.to_datetime(df[col_fecha], errors='coerce').max()
    return max_fecha if pd.notna(max_fecha) else None

def _construir_tabla_zonal2(df_rh, df_rt, df_altas, zonal2_vals):
    """
    Construye tabla con supervisores y vendedores de zonales específicas.
    Usa RRHH (MiniMatriz) como base principal y suma RT/ALTAS sobre vendedores de esas zonales.

    Args:
        df_rh: DataFrame de hoja RH (MiniMatriz con OPERADOR, ESTADO, feedback_rh, etc.)
        df_rt: DataFrame de hoja RT (para contar RT y RUS por vendedor)
        df_altas: DataFrame de hoja ALTAS (para contar ALTAS por vendedor)
        zonal2_vals: lista de valores zonal2 a incluir (ej. ['TACNA', 'ILO'])

    Returns:
        DataFrame con columnas: SUPERVISOR, VENDEDOR, RT, RUS, ALTAS, ALTAS_REGULAR, ALTAS_FLEX
    """
    # Calcular zonal2 en RRHH
    df_rh['zonal2'] = df_rh['ZONA'].apply(_calcular_zonal2)

    # Filtrar RRHH por zonales (base principal)
    rh_filt = df_rh[df_rh['zonal2'].isin(zonal2_vals)].copy()

    # Aplicar filtros: OPERADOR='MOVISTAR', ESTADO='ACTIVO' y feedback_rh='EN CAMPO'
    rh_filt = rh_filt[
        (rh_filt['OPERADOR'].fillna('').str.strip().str.upper() == 'MOVISTAR') &
        (rh_filt['ESTADO'].fillna('').str.strip().str.upper() == 'ACTIVO') &
        (rh_filt['feedback_rh'].fillna('').str.strip().str.upper() == 'EN CAMPO')
    ].copy()

    # Reemplazar SUPERVISOR vacío o "0" con "SIN ASIGNAR"
    rh_filt['SUPERVISOR'] = rh_filt['SUPERVISOR'].fillna('SIN ASIGNAR')
    rh_filt.loc[rh_filt['SUPERVISOR'] == '0', 'SUPERVISOR'] = 'SIN ASIGNAR'

    # Tabla base: SUPERVISOR + VENDEDOR desde RH
    tabla = rh_filt[['SUPERVISOR', 'VENDEDOR']].drop_duplicates().copy()
    tabla = tabla.sort_values(['SUPERVISOR', 'VENDEDOR']).reset_index(drop=True)

    # Calcular zonal2 en RT y ALTAS para los joins
    df_rt = df_rt.copy()
    df_altas = df_altas.copy()
    df_rt['zonal2'] = df_rt['zonal'].apply(_calcular_zonal2)
    df_altas['zonal2'] = df_altas['zonal'].apply(_calcular_zonal2)

    # Filtrar RT y ALTAS por zonales
    rt_filt = df_rt[df_rt['zonal2'].isin(zonal2_vals)].copy()
    altas_filt = df_altas[df_altas['zonal2'].isin(zonal2_vals)].copy()

    # Contar RT por VENDEDOR
    rt_counts = rt_filt.groupby('VENDEDOR').size().reset_index(name='RT')

    # Sumar Flag_Registro_Unico (RUS) por VENDEDOR (valores numéricos del Excel)
    # Flag_Registro_Unico puede venir como float/int, llenarlo con 0 si falta
    if 'Flag_Registro_Unico' in rt_filt.columns:
        rus_counts = rt_filt.groupby('VENDEDOR')['Flag_Registro_Unico'].sum().reset_index(name='RUS')
        rus_counts['RUS'] = pd.to_numeric(rus_counts['RUS'], errors='coerce').fillna(0).astype(int)
    else:
        # Si no existe la columna, dejar RUS en 0
        rus_counts = rt_filt.groupby('VENDEDOR').size().reset_index(name='RUS')
        rus_counts['RUS'] = 0

    # Contar ALTAS totales por VENDEDOR
    altas_counts = altas_filt.groupby('VENDEDOR').size().reset_index(name='ALTAS')

    # Contar ALTAS por SCORING
    altas_regular = altas_filt[altas_filt['Scoring'] == 'REGULAR'].groupby('VENDEDOR').size().reset_index(name='ALTAS_REGULAR')
    altas_flex = altas_filt[altas_filt['Scoring'] == 'FLEX'].groupby('VENDEDOR').size().reset_index(name='ALTAS_FLEX')

    # Contar ALTAS con RIESG >= 1 (riesgo Integratel) por VENDEDOR → columna AXB
    if 'RIESG' in altas_filt.columns:
        altas_riesg = pd.to_numeric(altas_filt['RIESG'], errors='coerce').fillna(0)
        axb_counts = altas_filt[altas_riesg >= 1].groupby('VENDEDOR').size().reset_index(name='AXB')
    else:
        axb_counts = pd.DataFrame(columns=['VENDEDOR', 'AXB'])

    # Merge con tabla base (RRHH)
    tabla = tabla.merge(rt_counts, on='VENDEDOR', how='left')
    tabla['RT'] = tabla['RT'].fillna(0).astype(int)

    tabla = tabla.merge(rus_counts, on='VENDEDOR', how='left')
    tabla['RUS'] = tabla['RUS'].fillna(0).astype(int)

    tabla = tabla.merge(altas_counts, on='VENDEDOR', how='left')
    tabla['ALTAS'] = tabla['ALTAS'].fillna(0).astype(int)

    tabla = tabla.merge(altas_regular, on='VENDEDOR', how='left')
    tabla['ALTAS_REGULAR'] = tabla['ALTAS_REGULAR'].fillna(0).astype(int)

    tabla = tabla.merge(altas_flex, on='VENDEDOR', how='left')
    tabla['ALTAS_FLEX'] = tabla['ALTAS_FLEX'].fillna(0).astype(int)

    tabla = tabla.merge(axb_counts, on='VENDEDOR', how='left')
    tabla['AXB'] = tabla['AXB'].fillna(0).astype(int)

    # Reordenar columnas
    tabla = tabla[['SUPERVISOR', 'VENDEDOR', 'RT', 'RUS', 'ALTAS', 'ALTAS_REGULAR', 'ALTAS_FLEX', 'AXB']]
    tabla = tabla.sort_values(['SUPERVISOR', 'VENDEDOR']).reset_index(drop=True)

    # Agregar subtotales por supervisor y total general
    tabla_con_subtotales = []
    supervisores = tabla['SUPERVISOR'].unique()

    for supervisor in supervisores:
        datos_sup = tabla[tabla['SUPERVISOR'] == supervisor]
        tabla_con_subtotales.append(datos_sup)

        # Agregar fila de subtotal para este supervisor
        subtotal_row = pd.DataFrame({
            'SUPERVISOR': [f'SUB: {supervisor}'],
            'VENDEDOR': [''],
            'RT': [datos_sup['RT'].sum()],
            'RUS': [datos_sup['RUS'].sum()],
            'ALTAS': [datos_sup['ALTAS'].sum()],
            'ALTAS_REGULAR': [datos_sup['ALTAS_REGULAR'].sum()],
            'ALTAS_FLEX': [datos_sup['ALTAS_FLEX'].sum()],
            'AXB': [datos_sup['AXB'].sum()]
        })
        tabla_con_subtotales.append(subtotal_row)

    # Concatenar todas las filas
    tabla_final = pd.concat(tabla_con_subtotales, ignore_index=True)

    # Agregar total general al final
    total_row = pd.DataFrame({
        'SUPERVISOR': ['TOTAL GENERAL'],
        'VENDEDOR': [''],
        'RT': [tabla['RT'].sum()],
        'RUS': [tabla['RUS'].sum()],
        'ALTAS': [tabla['ALTAS'].sum()],
        'ALTAS_REGULAR': [tabla['ALTAS_REGULAR'].sum()],
        'ALTAS_FLEX': [tabla['ALTAS_FLEX'].sum()],
        'AXB': [tabla['AXB'].sum()]
    })
    tabla_final = pd.concat([tabla_final, total_row], ignore_index=True)

    return tabla_final

def _crear_xlsx_tabla(tabla, titulo="Cuadro Resumen"):
    """
    Crea un archivo Excel con la tabla formateada.
    Retorna la ruta del archivo temporal.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen"

    # Estilos
    font_titulo = Font(name='Aptos Narrow', size=12, bold=True)
    font_header = Font(name='Aptos Narrow', size=11, bold=True)
    font_normal = Font(name='Aptos Narrow', size=11)
    font_subtotal = Font(name='Aptos Narrow', size=11, bold=True)

    fill_header = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    fill_subtotal = PatternFill(start_color="E8E8E8", end_color="E8E8E8", fill_type="solid")

    border_thin = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    alignment_center = Alignment(horizontal='center', vertical='center')
    alignment_left = Alignment(horizontal='left', vertical='center')
    alignment_right = Alignment(horizontal='right', vertical='center')

    # Título
    ws['A1'] = titulo
    ws['A1'].font = font_titulo
    ws.merge_cells('A1:E1')
    ws['A1'].alignment = alignment_center
    ws.row_dimensions[1].height = 20

    # Headers
    headers = ['SUPERVISOR', 'VENDEDOR', 'RT', 'RUS', 'ALTAS']
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.value = header
        cell.font = font_header
        cell.fill = fill_header
        cell.border = border_thin
        cell.alignment = alignment_center

    ws.row_dimensions[3].height = 15

    # Filas de datos
    row = 4
    supervisor_actual = None
    supervisor_row_inicio = None
    supervisor_subtotales = {}

    for idx, r in tabla.iterrows():
        sup = r['SUPERVISOR']
        vend = r['VENDEDOR']
        rt = int(r['RT'])
        rus = int(r['RUS'])
        altas = int(r['ALTAS'])

        # Detectar cambio de supervisor
        if sup != supervisor_actual:
            # Si hay supervisor anterior, insertar subtotal
            if supervisor_actual is not None and supervisor_row_inicio is not None:
                subtotal_row = row
                row += 1

                ws.cell(row=subtotal_row, column=1).value = f"SUP {supervisor_actual}"
                ws.cell(row=subtotal_row, column=2).value = ""
                ws.cell(row=subtotal_row, column=3).value = supervisor_subtotales['rt']
                ws.cell(row=subtotal_row, column=4).value = supervisor_subtotales['rus']
                ws.cell(row=subtotal_row, column=5).value = supervisor_subtotales['altas']

                for col_idx in range(1, 6):
                    cell = ws.cell(row=subtotal_row, column=col_idx)
                    cell.font = font_subtotal
                    cell.fill = fill_subtotal
                    cell.border = border_thin
                    if col_idx >= 3:
                        cell.alignment = alignment_right
                    else:
                        cell.alignment = alignment_left

            # Nuevo supervisor
            supervisor_actual = sup
            supervisor_row_inicio = row
            supervisor_subtotales = {'rt': 0, 'rus': 0, 'altas': 0}

        # Fila de vendedor
        ws.cell(row=row, column=1).value = sup
        ws.cell(row=row, column=2).value = vend
        ws.cell(row=row, column=3).value = rt
        ws.cell(row=row, column=4).value = rus
        ws.cell(row=row, column=5).value = altas

        for col_idx in range(1, 6):
            cell = ws.cell(row=row, column=col_idx)
            cell.font = font_normal
            cell.border = border_thin
            if col_idx >= 3:
                cell.alignment = alignment_right
            else:
                cell.alignment = alignment_left

        # Acumular para subtotal
        supervisor_subtotales['rt'] += rt
        supervisor_subtotales['rus'] += rus
        supervisor_subtotales['altas'] += altas

        row += 1

    # Subtotal final
    if supervisor_actual is not None:
        subtotal_row = row
        row += 1

        ws.cell(row=subtotal_row, column=1).value = f"SUP {supervisor_actual}"
        ws.cell(row=subtotal_row, column=2).value = ""
        ws.cell(row=subtotal_row, column=3).value = supervisor_subtotales['rt']
        ws.cell(row=subtotal_row, column=4).value = supervisor_subtotales['rus']
        ws.cell(row=subtotal_row, column=5).value = supervisor_subtotales['altas']

        for col_idx in range(1, 6):
            cell = ws.cell(row=subtotal_row, column=col_idx)
            cell.font = font_subtotal
            cell.fill = fill_subtotal
            cell.border = border_thin
            if col_idx >= 3:
                cell.alignment = alignment_right
            else:
                cell.alignment = alignment_left

    # Total general
    total_row = row + 1
    ws.cell(row=total_row, column=1).value = "Total general"
    ws.cell(row=total_row, column=3).value = tabla['RT'].sum()
    ws.cell(row=total_row, column=4).value = tabla['RUS'].sum()
    ws.cell(row=total_row, column=5).value = tabla['ALTAS'].sum()

    for col_idx in range(1, 6):
        cell = ws.cell(row=total_row, column=col_idx)
        cell.font = font_subtotal
        cell.border = border_thin
        if col_idx >= 3:
            cell.alignment = alignment_right
        else:
            cell.alignment = alignment_left

    # Ajustar anchos
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 10

    # Guardar
    temp_file = os.path.join(tempfile.gettempdir(), f"cuadro_resumen_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb.save(temp_file)

    return temp_file

def _capturar_imagen_tabla_pil(tabla_df, titulo="Cuadro", subtitulo=None):
    """
    Renderiza tabla compacta con Aptos Narrow tamaño 13 y alta resolución.
    Título: "[REGIÓN] - Avance al [FECHA]" en negrita.
    Subtítulo opcional debajo del título: "Reporte de Ventas del [SUPERVISOR]".
    Columnas: VENDEDOR, RT, RUS, ALT_REG, ALT_FLEX (sin SUPERVISOR, sin TOTAL GENERAL)
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        print("[WARN] PIL no disponible")
        return None

    temp_png = os.path.join(
        tempfile.gettempdir(),
        f"tabla_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    )

    try:
        # Fuentes Aptos Narrow tamaño 12 (más legible)
        # Office instala Aptos en fuentes por-usuario, no en C:\Windows\Fonts
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

        # Parámetros de diseño tipo MODELO1 (compacto pero legible, sin truncar nombres)
        margin = 10
        padding_x = 4
        cell_height = int(18 * 0.9 * 0.9 * 0.9)  # Reducido 10% + 10% + 10% adicional
        header_height = int(20 * 0.9 * 0.9 * 0.9)  # Reducido 10% + 10% + 10% adicional
        title_height = int(22 * 0.9 * 0.9 * 0.9) + 4  # Reducido 10%+10%+10%, +4 por tamaño de fuente 16
        subtitulo_height = (int(18 * 0.9 * 0.9 * 0.9) + 4) if subtitulo else 0  # +4 por tamaño de fuente 15
        col_spacing = 0
        scale_factor = 2  # Escala 2x para mejor resolución
        line_width = 1  # Grosor de líneas

        # Calcular columna ALTAS = ALT_REG + ALT_FLEX por vendedor
        tabla_df = tabla_df.copy()
        tabla_df['ALTAS_TOTAL'] = (
            pd.to_numeric(tabla_df['ALTAS_REGULAR'], errors='coerce').fillna(0) +
            pd.to_numeric(tabla_df['ALTAS_FLEX'], errors='coerce').fillna(0)
        ).astype(int)

        # Columnas con etiquetas (sin SUPERVISOR: ahora va como subtítulo)
        headers = ['VENDEDOR', 'RT', 'RUS', 'ALTAS', 'REGULAR', 'FLEX', 'AXB']
        # Mapeo de columnas internas del DataFrame
        col_mapping = {
            'VENDEDOR': 'VENDEDOR',
            'RT': 'RT',
            'RUS': 'RUS',
            'ALTAS': 'ALTAS_TOTAL',
            'REGULAR': 'ALTAS_REGULAR',
            'FLEX': 'ALTAS_FLEX',
            'AXB': 'AXB'
        }

        # Calcular ancho necesario para cada columna (SIN TRUNCAR nombres)
        col_widths = {}
        temp_img = Image.new('RGB', (1, 1))
        temp_draw = ImageDraw.Draw(temp_img)

        for display_header in headers:
            # Ancho del header
            bbox = temp_draw.textbbox((0, 0), display_header, font=font_header)
            header_width = bbox[2] - bbox[0] + padding_x * 2

            # Ancho de los datos
            if display_header in ['REGULAR', 'FLEX']:
                # Columnas numéricas ALT_REG/ALT_FLEX: ancho fijo
                data_width = 40
            elif display_header in ['RT', 'RUS', 'ALTAS', 'AXB']:
                # Columnas numéricas: estrechas pero proporcionadas
                data_width = 24
            else:
                # Columnas de texto: ancho EXACTO del contenido más largo con límite máximo
                max_width = header_width
                internal_col = col_mapping[display_header]
                for row_val in tabla_df[internal_col]:
                    val_str = str(row_val)  # SIN TRUNCAR
                    if val_str.strip():  # Solo si no está vacío
                        bbox = temp_draw.textbbox((0, 0), val_str, font=font_normal)
                        val_width = bbox[2] - bbox[0] + padding_x * 2
                        max_width = max(max_width, val_width)
                # Límite máximo de 120 píxeles para columnas de texto
                data_width = min(max_width, 120)

            if display_header in ['REGULAR', 'FLEX']:
                # Ancho forzado: ignora el ancho del header aunque el texto se recorte
                col_widths[display_header] = data_width
            else:
                col_widths[display_header] = max(header_width, data_width)

        # Calcular dimensiones de la imagen
        total_width = sum(col_widths.values()) + len(col_widths) * col_spacing + margin * 2
        num_rows = len(tabla_df)
        img_height = title_height + subtitulo_height + header_height + (num_rows * cell_height) + margin * 2

        # Crear imagen escalada (2x para mejor resolución)
        img_width_scaled = int(total_width * scale_factor)
        img_height_scaled = int(img_height * scale_factor)
        img = Image.new('RGB', (img_width_scaled, img_height_scaled), color='white')
        draw = ImageDraw.Draw(img)

        def scale(val):
            return int(val * scale_factor)

        # Dibujar título en negrita
        draw.text((scale(margin), scale(margin)), titulo, fill='black', font=font_title)

        # Dibujar subtítulo (Reporte de [SUPERVISOR]) debajo del título, en azul y negrita
        if subtitulo:
            draw.text((scale(margin), scale(margin + title_height)), subtitulo, fill='#0000FF', font=font_subtitulo)

        # Dibujar encabezados (una sola fila)
        y = scale(margin + title_height + subtitulo_height)
        x = scale(margin)

        for display_header in headers:
            col_width = scale(col_widths[display_header])

            # Fondo verde agua para encabezado con línea más gruesa
            draw.rectangle([x, y, x + col_width, y + scale(header_height)],
                          fill='#B2F7EF', outline='black', width=2)

            # Texto centrado en negrita
            bbox = draw.textbbox((0, 0), display_header, font=font_header)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            text_x = x + (col_width - text_width) // 2
            text_y = y + (scale(header_height) - text_height) // 2
            draw.text((text_x, text_y), display_header, fill='black', font=font_header)

            x += col_width + scale(col_spacing)

        # Dibujar filas de datos (tabla_df viene filtrada por supervisor, con su fila SUB: al final)
        y = scale(margin + title_height + subtitulo_height + header_height)
        for row_idx, row in tabla_df.iterrows():
            supervisor_val = str(row.get('SUPERVISOR', ''))
            is_subtotal = supervisor_val.startswith('SUB:')

            if is_subtotal:
                bg_color = '#E8E8E8'  # Gris claro para subtotal
                cell_font = font_header
            else:
                bg_color = 'white'    # Blanco para vendedores
                cell_font = font_normal

            x = scale(margin)
            for display_header in headers:
                col_width = scale(col_widths[display_header])
                internal_col = col_mapping[display_header]

                # Fondo de celda con gridlines definidas
                draw.rectangle([x, y, x + col_width, y + scale(cell_height)],
                              fill=bg_color, outline='#666666', width=1)

                # Obtener valor (SIN TRUNCAR)
                if is_subtotal and display_header == 'VENDEDOR':
                    val_str = 'SUBTOTAL'
                else:
                    val = row.get(internal_col, '')
                    val_str = str(val)

                # Alineación: centrada para números, izquierda para texto
                if display_header in ['RT', 'RUS', 'ALTAS', 'REGULAR', 'FLEX', 'AXB']:
                    # Alineación centrada para números
                    bbox = draw.textbbox((0, 0), val_str, font=cell_font)
                    text_width = bbox[2] - bbox[0]
                    text_x = x + (col_width - text_width) // 2
                else:
                    # Alineación izquierda para texto
                    text_x = x + scale(padding_x)

                # Centrar verticalmente
                bbox = draw.textbbox((0, 0), val_str, font=cell_font)
                text_height = bbox[3] - bbox[1]
                text_y = y + (scale(cell_height) - text_height) // 2

                draw.text((text_x, text_y), val_str, fill='black', font=cell_font)

                x += col_width + scale(col_spacing)

            y += scale(cell_height)

        # Borde grueso perimetral alrededor de toda la tabla (encabezado + filas)
        table_top = scale(margin + title_height + subtitulo_height)
        table_left = scale(margin)
        table_right = scale(margin) + sum(scale(w) for w in col_widths.values())
        table_bottom = y
        thick_border_width = scale(1)
        draw.rectangle([table_left, table_top, table_right, table_bottom],
                      outline='black', width=thick_border_width)

        print(f"[DEBUG] Guardando imagen a: {temp_png} ({img_width_scaled}x{img_height_scaled})")
        img.save(temp_png, format='PNG')
        file_size = os.path.getsize(temp_png)
        print(f"[INFO] Imagen creada: {temp_png} ({file_size} bytes)")
        return temp_png

    except Exception as e:
        print(f"[ERROR] PIL rendering falló: {e}")
        import traceback
        traceback.print_exc()
        return None

def _capturar_imagen_tabla(tabla_df, titulo="Cuadro", subtitulo=None):
    """
    Captura la tabla como imagen PNG usando PIL.
    Retorna la ruta del archivo PNG.
    """
    return _capturar_imagen_tabla_pil(tabla_df, titulo, subtitulo)

def procesar_cuadro_resumen_sup(avance_path):
    """
    Procesa cuadro resumen de supervisores y envía imágenes por WhatsApp.

    Args:
        avance_path: ruta al archivo AVANCE_{fecha}.xlsx

    Returns:
        bool: True si al menos una imagen se envió exitosamente, False si todas fallaron.
    """
    if not avance_path or not os.path.exists(avance_path):
        print(f"[CUADRO RESUMEN SUP] Archivo no encontrado: {avance_path}")
        return False

    print("[CUADRO RESUMEN SUP] Iniciando...")

    try:
        # Cargar datos
        df_rh = pd.read_excel(avance_path, sheet_name='RH')
        df_rt = pd.read_excel(avance_path, sheet_name='RT')
        df_altas = pd.read_excel(avance_path, sheet_name='ALTAS')

        # Obtener fecha máxima
        max_fecha = _obtener_fecha_maxima(df_rt)
        if max_fecha:
            fecha_str = max_fecha.strftime("%d/%m/%Y")
        else:
            fecha_str = datetime.now().strftime("%d/%m/%Y")

        # Definir grupos de zonales y sus destinos
        grupos = {
            'FRANZ_TACNA': {
                'zonales': ['TACNA'],
                'numero': '51933540190@c.us',
                'titulo': 'TACNA'
            },
            'FRANZ_ILO': {
                'zonales': ['ILO'],
                'numero': '51933540190@c.us',
                'titulo': 'ILO'
            },
            'LETICIA_TRUJILLO': {
                'zonales': ['TRUJILLO'],
                'numero': '51952099313@c.us',
                'titulo': 'TRUJILLO'
            },
            'LETICIA_CHIMBOTE': {
                'zonales': ['CHIMBOTE'],
                'numero': '51952099313@c.us',
                'titulo': 'CHIMBOTE'
            },
            'LETICIA_HUARAZ': {
                'zonales': ['HUARAZ'],
                'numero': '51952099313@c.us',
                'titulo': 'HUARAZ'
            },
            # Deshabilitado: No enviar a Carlos Parra por ahora
            # 'CARLOS PARRA': {
            #     'zonales': ['LIMA'],
            #     'numero': 'Carlos Parra',  # Usa alias de config.json
            #     'titulo': 'LIMA'
            # },
            # AREQUIPA está en stand by
        }

        wa = WhatsAppClient()
        envios_ok = 0
        total_intentos = 0

        # Procesar cada grupo
        for jefe, config in grupos.items():
            zonales = config['zonales']
            destino = config['numero']
            titulo = config['titulo']

            print(f"  Procesando {jefe} ({titulo})...")
            total_intentos += 1

            try:
                # Construir tabla (RRHH es base principal)
                tabla = _construir_tabla_zonal2(df_rh, df_rt, df_altas, zonales)

                if len(tabla) == 0:
                    print(f"    → Sin datos para {titulo}")
                    continue

                # Quitar filas de subtotal (SUB:) y total general; agrupar por supervisor real
                tabla_vendedores = tabla[
                    ~tabla['SUPERVISOR'].str.startswith('SUB:') &
                    (tabla['SUPERVISOR'] != 'TOTAL GENERAL')
                ].copy()

                supervisores = tabla_vendedores['SUPERVISOR'].unique()

                # Una imagen por supervisor
                for supervisor in supervisores:
                    tabla_sup = tabla_vendedores[tabla_vendedores['SUPERVISOR'] == supervisor].copy()

                    if len(tabla_sup) == 0:
                        continue

                    # Agregar fila de subtotal del supervisor al final de su propia tabla
                    subtotal_row = pd.DataFrame({
                        'SUPERVISOR': [f'SUB: {supervisor}'],
                        'VENDEDOR': [''],
                        'RT': [tabla_sup['RT'].sum()],
                        'RUS': [tabla_sup['RUS'].sum()],
                        'ALTAS': [tabla_sup['ALTAS'].sum()] if 'ALTAS' in tabla_sup.columns else [0],
                        'ALTAS_REGULAR': [tabla_sup['ALTAS_REGULAR'].sum()],
                        'ALTAS_FLEX': [tabla_sup['ALTAS_FLEX'].sum()],
                        'AXB': [tabla_sup['AXB'].sum()] if 'AXB' in tabla_sup.columns else [0]
                    })
                    tabla_sup = pd.concat([tabla_sup, subtotal_row], ignore_index=True)

                    subtitulo = f"Reporte de {supervisor}"

                    # Crear Excel (opcional, para backup)
                    xlsx_path = _crear_xlsx_tabla(tabla_sup, titulo=f"{titulo} - Avance al {fecha_str}")

                    # Capturar imagen directamente del DataFrame con nuevo formato
                    png_path = _capturar_imagen_tabla(
                        tabla_sup,
                        titulo=f"{titulo} - Avance al {fecha_str}",
                        subtitulo=subtitulo
                    )

                    if png_path and os.path.exists(png_path):
                        # Enviar por WhatsApp
                        mensaje = f"Avance de {supervisor} al {fecha_str}"

                        wa.send_image(destino, png_path, caption=mensaje)
                        print(f"    [OK] Enviado a {jefe} ({supervisor})")
                        envios_ok += 1

                        # Limpiar
                        try:
                            os.remove(png_path)
                        except:
                            pass
                    else:
                        print(f"    [ERROR] Falló captura de imagen para {jefe} ({supervisor})")

                    # Limpiar Excel temporal
                    try:
                        os.remove(xlsx_path)
                    except:
                        pass

                    time.sleep(2)  # Pequeña pausa entre envíos

            except Exception as e:
                print(f"    [ERROR] Procesando {jefe}: {e}")

        print(f"[CUADRO RESUMEN SUP] Completado: {envios_ok}/{total_intentos} envíos exitosos")
        return envios_ok > 0

    except Exception as e:
        print(f"[ERROR] Cuadro resumen sup: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    # Para pruebas
    avance_file = r"C:\proyectos\AVANCE_MOVISTAR\AVANCE_2026-05-06.xlsx"
    procesar_cuadro_resumen_sup(avance_file)
