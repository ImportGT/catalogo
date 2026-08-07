import os
import shutil

# Rutas de origen y destino relativas al proyecto
origen = "ANILLOS SWA"
destino = os.path.join("imagenes", "SWA", "anillos_swa")

# Asegurar que la carpeta de destino exista
os.makedirs(destino, exist_ok=True)

if not os.path.exists(origen):
    print(f"La carpeta de origen '{origen}' no existe.")
else:
    archivos = os.listdir(origen)
    print(f"Encontrados {len(archivos)} archivos en '{origen}'. Procesando...")
    
    count = 1
    for archivo in archivos:
        ruta_origen = os.path.join(origen, archivo)
        if os.path.isfile(ruta_origen):
            # Obtenemos la extensión original (ej. .jpg, .webp, .mp4)
            _, ext = os.path.splitext(archivo)
            
            # Nuevo nombre estandarizado
            nuevo_nombre = f"anillos_swa_{count}{ext}"
            ruta_destino = os.path.join(destino, nuevo_nombre)
            
            shutil.copy2(ruta_origen, ruta_destino)
            print(f"Copiado y renombrado: {archivo} -> {nuevo_nombre}")
            count += 1

    print("¡Proceso de renombrado y copia finalizado con éxito!")