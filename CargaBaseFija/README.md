# Cargador BBDD Fija

Herramienta standalone para Katerine Palacios: un `.exe` que ella ejecuta con
doble clic, elige su archivo `BDD....xlsx` (parte por defecto en la carpeta de
red del proyecto FIJA) y sube la hoja `BBDD Fija` como valores al Google Sheet
de destino, reemplazando el contenido de la pestaña con GID `407197609`.

No requiere Python instalado en su equipo. No requiere que ella inicie sesión
en ninguna cuenta de Google.

## Componentes

- `cargar_base_fija.py` — script fuente.
- `service_account.json` — credencial de la cuenta de servicio de Google Cloud
  (secreto, **nunca** se sube a git; ver `.gitignore`). Debe copiarse junto al
  `.exe` compilado, en la misma carpeta.
- `build_exe.ps1` — compila el `.exe` con PyInstaller.
- `requirements_build.txt` — solo `pyinstaller` (el resto de dependencias ya
  están en `C:\proyectos\requirements.txt`, el venv compartido).

## Preparar la cuenta de servicio (una sola vez)

1. Google Cloud Console → crear/seleccionar proyecto → habilitar **Google
   Sheets API**.
2. IAM y administración → Cuentas de servicio → crear una (sin rol de GCP
   necesario).
3. Esa cuenta → pestaña Claves → Agregar clave → JSON → se descarga el
   archivo. Renombrarlo a `service_account.json` y colocarlo en esta carpeta.
4. Compartir el Google Sheet destino con el correo de la cuenta de servicio
   (termina en `.iam.gserviceaccount.com`) como **Editor**.

## Compilar el .exe

```powershell
C:\proyectos\.venv\Scripts\Activate.ps1
cd C:\proyectos\AVANCE_MOVISTAR\CargaBaseFija
pip install -r requirements_build.txt
.\build_exe.ps1
```

Resultado: `dist\Cargador_BBDD_Fija.exe`.

## Entregar a Katerine

Copiar a su PC, en la **misma carpeta**:
- `Cargador_BBDD_Fija.exe`
- `service_account.json`

Sugerido: crear un acceso directo al `.exe` en su Escritorio. Ella solo hace
doble clic, elige el archivo `BDD...` correspondiente al día, y espera el
pop-up de confirmación.

## Configuración (editar en `cargar_base_fija.py` si cambia algo)

| Variable | Valor actual |
|---|---|
| `CARPETA_PREDETERMINADA` | `\\SERVER\compartido\Reclutamiento\Reclutamiento General\PROYECTOS\PROYECTO FIJA\Base Fija\FIJA` |
| `PATRON_ARCHIVO` | `BDD*.xlsx` |
| `HOJA_ORIGEN` | `BBDD Fija` (⚠ nombre a confirmar — si Google Sheets no encuentra la hoja, el pop-up de error lista las hojas disponibles del archivo) |
| `SHEET_ID` | `1UwIkcOq_pqDRHfMrsH-RnGWRQqdgUpNLytsaX_RHmyc` |
| `GID_DESTINO` | `407197609` |

Si `HOJA_ORIGEN` no coincide, Katerine verá en el mensaje de error la lista
exacta de pestañas de su archivo — con eso se corrige el valor en el script y
se recompila.

## Modo de carga

Reemplaza **todo** el contenido de la pestaña destino (`clear()` + `update()`)
con los valores de la hoja origen, sin fórmulas ni formato (equivalente a
"Pegado especial → solo valores").

## Si algo falla

El `.exe` muestra un pop-up de error con el motivo, y además guarda el detalle
técnico completo en `error_ultima_carga.log` (misma carpeta del .exe) para
diagnóstico.
