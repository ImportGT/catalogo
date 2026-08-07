@echo off
echo ========================================
echo Subiendo cambios a GitHub...
echo ========================================

git status
git add .
set /p mensaje="Actualizacion de precios e imagenes de charms beads: "
git commit -m "%mensaje%"
git push origin main

echo ========================================
echo ¡Proceso completado con exito!
echo ========================================
pause