# Compila cargar_base_fija.py como un .exe de un solo archivo con PyInstaller.
# Ejecutar desde esta carpeta (CargaBaseFija) con el venv compartido activado:
#   C:\proyectos\.venv\Scripts\Activate.ps1
#   pip install -r requirements_build.txt
#   .\build_exe.ps1

pyinstaller --onefile --noconsole `
    --name "Cargador_BBDD_Fija" `
    --distpath ".\dist" `
    --workpath ".\build" `
    --specpath "." `
    cargar_base_fija.py

Write-Host ""
Write-Host "Listo. El ejecutable esta en .\dist\Cargador_BBDD_Fija.exe"
Write-Host "Antes de entregarselo a Katerine, copia junto al .exe:"
Write-Host "  - service_account.json (la credencial de la cuenta de servicio)"
Write-Host "Ambos archivos deben quedar en la MISMA carpeta en su PC."
