import os
import re
import shutil

CARPETA_ORIGEN = "ANILLOS BP"
CARPETA_DESTINO = os.path.join("imagenes", "BP", "anillosbp")
PREFIJO = "anillosbp"

def renombrar_anillos_bp():
    if not os.path.exists(CARPETA_ORIGEN):
        print(f"⚠️ No se encontró la carpeta de origen: {CARPETA_ORIGEN}")
        return

    if not os.path.exists(CARPETA_DESTINO):
        os.makedirs(CARPETA_DESTINO, exist_ok=True)
        print(f"📁 Carpeta creada: {CARPETA_DESTINO}")

    archivos = os.listdir(CARPETA_ORIGEN)
    print(f"🔍 Encontrados {len(archivos)} archivos en '{CARPETA_ORIGEN}'. Procesando...")

    count = 0
    for archivo in archivos:
        # Extraer extensión
        ext = archivo.split('.')[-1].lower()
        if ext not in ['jpg', 'jpeg', 'png', 'webp', 'avif', 'mp4', 'mov', 'webm']:
            continue

        # Buscar cualquier número presente en el nombre original para mantener su ID original
        numeros = re.findall(r'\d+', archivo)
        if not numeros:
            print(f"⚠️ Omitido (sin número identificador): {archivo}")
            continue

        # Usamos el primer número encontrado como ID principal
        prod_id = numeros[0]
        
        # Si hay más de un número (ej. sub-índices como anillos 1.2 o 1_2)
        if len(numeros) > 1:
            sub_id = numeros[1]
            nuevo_nombre = f"{PREFIJO}_{prod_id}.{sub_id}.{ext}"
        else:
            nuevo_nombre = f"{PREFIJO}_{prod_id}.{ext}"

        origen_path = os.path.join(CARPETA_ORIGEN, archivo)
        destino_path = os.path.join(CARPETA_DESTINO, nuevo_nombre)

        shutil.copy2(origen_path, destino_path)
        print(f"✅ {archivo} ---> {nuevo_nombre}")
        count += 1

    print(f"\n🎉 ¡Proceso finalizado! {count} archivos copiados y renombrados correctamente en '{CARPETA_DESTINO}'.")

if __name__ == "__main__":
    renombrar_anillos_bp()