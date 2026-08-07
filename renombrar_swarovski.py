import os
import shutil
import re

# Rutas de origen y destino para las pulseras de Swarovski
origen = "PULSERAS SWA"
destino = os.path.join("imagenes", "SWA", "pulseras_swa")

# Asegurar que la carpeta de destino exista
os.makedirs(destino, exist_ok=True)

if not os.path.exists(origen):
    print(f"La carpeta de origen '{origen}' no existe.")
else:
    archivos = os.listdir(origen)
    
    # Función para ordenar numéricamente de forma estricta (ej. 1, 2, 3 ... 10 en lugar de 1, 10, 2)
    def extraer_numero(nombre):
        nums = re.findall(r'\d+', nombre)
        return int(nums[0]) if nums else 0

    archivos_ordenados = sorted([f for f in archivos if os.path.isfile(os.path.join(origen, f))], key=extraer_numero)
    print(f"Procesando {len(archivos_ordenados)} archivos ordenados numéricamente desde '{origen}'...")
    
    count = 1
    for archivo in archivos_ordenados:
        ruta_origen = os.path.join(origen, archivo)
        _, ext = os.path.splitext(archivo)
        
        # Nuevo nombre estandarizado alineado perfectamente con el ID del Excel: pulseras_swa_1.jpg, etc.
        nuevo_nombre = f"pulseras_swa_{count}{ext.lower()}"
        ruta_destino = os.path.join(destino, nuevo_nombre)
        
        shutil.copy2(ruta_origen, ruta_destino)
        print(f"Alineado y copiado: {archivo} -> {nuevo_nombre}")
        count += 1

    print("¡Proceso de renombrado y sincronización de pulseras finalizado con éxito!")