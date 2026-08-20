import os
from PIL import Image

# Carpeta donde están tus imágenes .avif (puedes cambiar la ruta si están en una subcarpeta)
carpeta_imagenes = "." 

for archivo in os.listdir(carpeta_imagenes):
    if archivo.lower().endswith(".avif"):
        ruta_avif = os.path.join(carpeta_imagenes, archivo)
        nombre_base = os.path.splitext(archivo)[0]
        ruta_jpg = os.path.join(carpeta_imagenes, nombre_base + ".jpg")
        
        try:
            with Image.open(ruta_avif) as img:
                # Convertir a RGB por si el avif tiene transparencia o formato especial
                img_rgb = img.convert("RGB")
                img_rgb.save(ruta_jpg, "JPEG", quality=90)
            print(f"Convertido: {archivo} -> {nombre_base}.jpg")
        except Exception as e:
            print(f"Error al convertir {archivo}: {e}")

print("¡Conversión masiva finalizada!")