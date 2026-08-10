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
    Usa RRHH como base principal y suma RT/ALTAS sobre vendedores de esas zonales.

    Args:
        df_rh: DataFrame de hoja RH (estructura base: SUPERVISOR, VENDEDOR, ZONA, etc.)
        df_rt: DataFrame de hoja RT (para contar RT por vendedor)
        df_altas: DataFrame de hoja ALTAS (para contar ALTAS por vendedor)
        zonal2_vals: lista de valores zonal2 a incluir (ej. ['TACNA', 'ILO'])

    Returns:
        DataFrame con columnas: SUPERVISOR, VENDEDOR, RT, RUS, ALTAS, ALTAS_REGULAR, ALTAS_FLEX
    """
    # Calcular zonal2 en RRHH
    df_rh['zonal2'] = df_rh['ZONA'].apply(_calcular_zonal2)

    # Filtrar RRHH por zonales (base principal)
    rh_filt = df_rh[df_rh['zonal2'].isin(zonal2_vals)].copy()

    # Aplicar filtros: ESTADO='ACTIVO' y feedback_rh='EN CAMPO'
    rh_filt = rh_filt[
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
    df_rt['zonal2'] = df_rt['zonal'].apply(_calcular_zonal2)
    df_altas['zonal2'] = df_altas['zonal'].apply(_calcular_zonal2)

    # Filtrar RT y ALTAS por zonales
    rt_filt = df_rt[df_rt['zonal2'].isin(zonal2_vals)].copy()
    altas_filt = df_altas[df_altas['zonal2'].isin(zonal2_vals)].copy()

    # Contar RT por VENDEDOR
    rt_counts = rt_filt.groupby('VENDEDOR').size().reset_index(name='RT')

    # Contar ALTAS totales por VENDEDOR
    altas_counts = altas_filt.groupby('VENDEDOR').size().reset_index(name='ALTAS')

    # Contar ALTAS por SCORING
    altas_regular = altas_filt[altas_filt['Scoring'] == 'REGULAR'].groupby('VENDEDOR').size().reset_index(name='ALTAS_REGULAR')
    altas_flex = altas_filt[altas_filt['Scoring'] == 'FLEX'].groupby('VENDEDOR').size().reset_index(name='ALTAS_FLEX')

    # Merge con tabla base (RRHH)
    tabla = tabla.merge(rt_counts, on='VENDEDOR', how='left')
    tabla['RT'] = tabla['RT'].fillna(0).astype(int)

    tabla = tabla.merge(altas_counts, on='VENDEDOR', how='left')
    tabla['ALTAS'] = tabla['ALTAS'].fillna(0).astype(int)

    tabla = tabla.merge(altas_regular, on='VENDEDOR', how='left')
    tabla['ALTAS_REGULAR'] = tabla['ALTAS_REGULAR'].fillna(0).astype(int)

    tabla = tabla.merge(altas_flex, on='VENDEDOR', how='left')
    tabla['ALTAS_FLEX'] = tabla['ALTAS_FLEX'].fillna(0).astype(int)

    # RUS: por ahora siempre 0
    tabla['RUS'] = 0

    # Reordenar columnas
    tabla = tabla[['SUPERVISOR', 'VENDEDOR', 'RT', 'RUS', 'ALTAS', 'ALTAS_REGULAR', 'ALTAS_FLEX']]
    tabla = tabla.sort_values(['SUPERVISOR', 'VENDEDOR']).reset_index(drop=True)

    return tabla

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

def _capturar_imagen_tabla_pil(tabla_df, titulo="Cuadro"):
    """
    Renderiza la tabla como imagen PNG usando PIL con ajuste automático de columnas.
    Retorna la ruta del archivo PNG.
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
        # Fuente (intentar Aptos Narrow o fallback)
        try:
            font_normal = ImageFont.truetype("C:\\Windows\\Fonts\\aptos.ttf", 9)
            font_bold = ImageFont.truetype("C:\\Windows\\Fonts\\aptos.ttf", 10)
            font_title = ImageFont.truetype("C:\\Windows\\Fonts\\aptos.ttf", 12)
        except:
            try:
                font_normal = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 9)
                font_bold = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 10)
                font_title = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 12)
            except:
                font_normal = ImageFont.load_default()
                font_bold = font_normal
                font_title = font_normal

        # Parámetros de diseño
        margin = 15
        padding_x = 8
        padding_y = 6
        cell_height = 24
        header_height = 30
        subheader_height = 18  # Altura para subheaders
        title_height = 28
        col_spacing = 2

        # Definir anchos de columna con autoajuste
        headers = ['SUPERVISOR', 'VENDEDOR', 'RT', 'RUS', 'ALTAS', 'ALTAS_REGULAR', 'ALTAS_FLEX']

        # Calcular ancho necesario para cada columna
        col_widths = {}
        temp_img = Image.new('RGB', (1, 1))
        temp_draw = ImageDraw.Draw(temp_img)

        for col_idx, header in enumerate(headers):
            # Ancho del header
            bbox = temp_draw.textbbox((0, 0), header, font=font_bold)
            header_width = bbox[2] - bbox[0] + padding_x * 2

            # Ancho de los datos
            if header in ['RT', 'RUS', 'ALTAS', 'ALTAS_REGULAR', 'ALTAS_FLEX']:
                # Columnas numéricas: ancho fijo
                data_width = 40
            else:
                # Columnas de texto: buscar el más largo
                max_width = header_width
                for row_val in tabla_df[header]:
                    val_str = str(row_val)[:40]  # Limitar a 40 caracteres
                    bbox = temp_draw.textbbox((0, 0), val_str, font=font_normal)
                    val_width = bbox[2] - bbox[0] + padding_x * 2
                    max_width = max(max_width, val_width)
                data_width = min(max_width, 250)  # Máximo 250px por columna

            col_widths[header] = max(header_width, data_width)

        # Calcular dimensiones de la imagen
        total_width = sum(col_widths.values()) + len(col_widths) * col_spacing + margin * 2
        num_rows = len(tabla_df)
        # Una sola fila de headers con dos niveles (ALTAS se combina sobre REGULAR y FLEX)
        img_height = title_height + header_height + (num_rows * cell_height) + margin * 2

        # Crear imagen
        img = Image.new('RGB', (total_width, img_height), color='white')
        draw = ImageDraw.Draw(img)

        # Dibujar título
        draw.text((margin, margin), titulo, fill='black', font=font_title)

        # Dibujar headers
        y = margin + title_height
        x = margin
        col_positions = {}

        # Calcular ancho combinado para ALTAS (abarca solo ALTAS_REGULAR + ALTAS_FLEX)
        altas_combined_width = col_widths['ALTAS_REGULAR'] + col_widths['ALTAS_FLEX'] + col_spacing

        for col_idx, header in enumerate(headers):
            col_width = col_widths[header]
            col_positions[header] = (x, col_width)

            if header == 'ALTAS':
                # Dibujar ALTAS como encabezado principal combinado (gris oscuro)
                # que abarca exactamente ALTAS_REGULAR + ALTAS_FLEX
                draw.rectangle([x, y, x + altas_combined_width, y + header_height],
                              fill='#D3D3D3', outline='black', width=1)
                # Texto centrado
                bbox = draw.textbbox((0, 0), 'ALTAS', font=font_bold)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x + (altas_combined_width - text_width) // 2
                text_y = y + (header_height - text_height) // 2
                draw.text((text_x, text_y), 'ALTAS', fill='black', font=font_bold)

            elif header in ['ALTAS_REGULAR', 'ALTAS_FLEX']:
                # Subheaders: REGULAR y FLEX bajo ALTAS (gris más claro)
                display_name = 'REGULAR' if header == 'ALTAS_REGULAR' else 'FLEX'
                draw.rectangle([x, y, x + col_width, y + header_height],
                              fill='#E8E8E8', outline='black', width=1)
                bbox = draw.textbbox((0, 0), display_name, font=font_normal)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x + (col_width - text_width) // 2
                text_y = y + (header_height - text_height) // 2
                draw.text((text_x, text_y), display_name, fill='black', font=font_normal)

            else:
                # Headers normales (SUPERVISOR, VENDEDOR, RT, RUS)
                draw.rectangle([x, y, x + col_width, y + header_height],
                              fill='#D3D3D3', outline='black', width=1)
                # Texto centrado
                bbox = draw.textbbox((0, 0), header, font=font_bold)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x + (col_width - text_width) // 2
                text_y = y + (header_height - text_height) // 2
                draw.text((text_x, text_y), header, fill='black', font=font_bold)

            x += col_width + col_spacing

        # Dibujar filas de datos
        y = margin + title_height + header_height
        for row_idx, row in tabla_df.iterrows():
            x = margin
            for col_idx, header in enumerate(headers):
                col_width = col_widths[header]

                # Fondo blanco para celda
                draw.rectangle([x, y, x + col_width, y + cell_height],
                              fill='white', outline='black', width=1)

                # Obtener valor
                val = row.get(header, '')
                val_str = str(val)[:40]  # Limitar a 40 caracteres

                # Alineación: derecha para números, izquierda para texto
                if header in ['RT', 'RUS', 'ALTAS', 'ALTAS_REGULAR', 'ALTAS_FLEX']:
                    # Alineación derecha para números
                    bbox = draw.textbbox((0, 0), val_str, font=font_normal)
                    text_width = bbox[2] - bbox[0]
                    text_x = x + col_width - text_width - padding_x
                else:
                    # Alineación izquierda para texto
                    text_x = x + padding_x

                # Centrar verticalmente
                bbox = draw.textbbox((0, 0), val_str, font=font_normal)
                text_height = bbox[3] - bbox[1]
                text_y = y + (cell_height - text_height) // 2

                draw.text((text_x, text_y), val_str, fill='black', font=font_normal)

                x += col_width + col_spacing

            y += cell_height

        print(f"[DEBUG] Guardando imagen a: {temp_png} ({total_width}x{img_height})")
        img.save(temp_png, format='PNG')
        file_size = os.path.getsize(temp_png)
        print(f"[INFO] Imagen creada: {temp_png} ({file_size} bytes)")
        return temp_png

    except Exception as e:
        print(f"[ERROR] PIL rendering falló: {e}")
        import traceback
        traceback.print_exc()
        return None

def _capturar_imagen_tabla(tabla_df, titulo="Cuadro"):
    """
    Captura la tabla como imagen PNG usando PIL.
    Retorna la ruta del archivo PNG.
    """
    return _capturar_imagen_tabla_pil(tabla_df, titulo)

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
            'FRANZ': {
                'zonales': ['TACNA', 'ILO'],
                'numero': '51933540190@c.us',
                'titulo': 'TACNA - ILO'
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

                # Crear Excel (opcional, para backup)
                xlsx_path = _crear_xlsx_tabla(tabla, titulo=f"Cuadro {titulo} - Avance al {fecha_str}")

                # Capturar imagen directamente del DataFrame
                png_path = _capturar_imagen_tabla(tabla, titulo=f"Cuadro {titulo} - Avance al {fecha_str}")

                if png_path and os.path.exists(png_path):
                    # Enviar por WhatsApp
                    # Mensaje personalizado por jefe
                    if jefe == 'FRANZ':
                        mensaje = f"Hola Franz, te comparto el cuadro avance por supervisor/vendedor según el último avance"
                    elif jefe.startswith('LETICIA'):
                        mensaje = f"Hola Leticia, te comparto el cuadro avance {titulo} por supervisor/vendedor según el último avance"
                    else:
                        mensaje = f"Cuadro resumen {titulo} - Avance al {fecha_str}"

                    wa.send_image(destino, png_path, caption=mensaje)
                    print(f"    [OK] Enviado a {jefe}")
                    envios_ok += 1

                    # Limpiar
                    try:
                        os.remove(png_path)
                    except:
                        pass
                else:
                    print(f"    [ERROR] Falló captura de imagen para {jefe}")

                # Limpiar Excel temporal
                try:
                    os.remove(xlsx_path)
                except:
                    pass

            except Exception as e:
                print(f"    [ERROR] Procesando {jefe}: {e}")

            time.sleep(2)  # Pequeña pausa entre envíos

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
