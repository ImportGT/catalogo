import os
import shutil
import re
import pandas as pd

# Rutas configuradas según tu estructura
ORIGEN = "ARETES BP"
DESTINO = os.path.join("imagenes", "BP", "aretessbp")
EXCEL_PATH = "ARETES BP.xlsx"
PREFIJO_DESTINO = "aretessbp"

# Asegurar que la carpeta de destino exista
os.makedirs(DESTINO, exist_ok=True)

if not os.path.exists(EXCEL_PATH):
    print(f"⚠️ Error: No se encontró el archivo Excel '{EXCEL_PATH}'.")
else:
    # Leer los IDs válidos desde el Excel (encabezado en la fila 3)
    try:
        df = pd.read_excel(EXCEL_PATH, header=3)
        ids_validos = set()
        for idx, row in df.iterrows():
            col_id = None
            for col in ['NUMERO', 'NUM.', 'ID', 'Id', 'numero', 'Número', 'ARETES']:
                if col in df.columns:
                    col_id = col
                    break
            if col_id and not pd.isna(row[col_id]):
                try:
                    ids_validos.add(int(row[col_id]))
                except:
                    pass
        print(f"📋 Se cargaron {len(ids_validos)} IDs válidos desde {EXCEL_PATH}.")
    except Exception as e:
        print(f"⚠️ Error al leer el Excel: {e}")
        ids_validos = set()

if not os.path.exists(ORIGEN):
    print(f"⚠️ La carpeta de origen '{ORIGEN}' no existe.")
else:
    archivos = [f for f in os.listdir(ORIGEN) if os.path.isfile(os.path.join(ORIGEN, f))]
    print(f"🔍 Encontrados {len(archivos)} archivos en la carpeta '{ORIGEN}'...")

    copiados = 0
    for archivo in archivos:
        ruta_origen = os.path.join(ORIGEN, archivo)
        nombre_base, ext = os.path.splitext(archivo)
        ext = ext.lower()

        # Filtrar solo formatos de imagen o video válidos
        if ext not in ('.jpg', '.jpeg', '.png', '.webp', '.avif', '.mp4', '.mov', '.webm'):
            continue

        # Extraer números del nombre original del archivo (ej. "aretes_1_2.jpg" -> [1, 2])
        numeros = re.findall(r'\d+', nombre_base)
        if not numeros:
            print(f"⏭️ Omitido (sin número identificador): {archivo}")
            continue

        # El primer número encontrado se asume como el ID principal del producto
        prod_id = int(numeros[0])

        if ids_validos and prod_id not in ids_validos:
            print(f"⚠️ Advertencia: El ID {prod_id} del archivo '{archivo}' no se encuentra en el Excel.")

        # Obtener el resto de la cadena numérica o sufijos (ej. .1, _2, etc.) si existen
        # Si hay más números o sufijos, los preservamos limpiamente
        sufijo = ""
        if len(numeros) > 1:
            # Reconstruir sufijo basado en cómo venía nombrado originalmente (ej. _2 o .2)
            resto_nombre = nombre_base.replace(str(prod_id), "", 1)
            sufijo_limpio = re.sub(r'[^a-zA-Z0-9_\.]', '', resto_nombre)
            if sufijo_limpio:
                if not sufijo_limpio.startswith(('_', '.')):
                    sufijo = "_" + sufijo_limpio
                else:
                    sufijo = sufijo_limpio

        # Nuevo nombre estandarizado manteniendo correspondencia exacta de ID
        nuevo_nombre = f"{PREFIJO_DESTINO}_{prod_id}{sufijo}{ext}"
        ruta_destino = os.path.join(DESTINO, nuevo_nombre)

        shutil.copy2(ruta_origen, ruta_destino)
        print(f"✅ Copiado y renombrado: {archivo} -> {nuevo_nombre}")
        copiados += 1

    print(f"\n🎉 ¡Proceso finalizado! Se procesaron y copiaron {copiados} archivos a '{DESTINO}'.")