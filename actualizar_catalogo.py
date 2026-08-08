import os
import re
import pandas as pd
import json

CONFIGURACIONES = [
    {
        "categoria": "Anillos",
        "excel": "ANILLOS.xlsx",
        "js": "anillos.js",
        "variable_js": "productosAnillos",
        "prefijo": "anillos",
        "carpeta": "imagenes/anillos"
    },
    {
        "categoria": "Aretes",
        "excel": "ARETES.xlsx",
        "js": "aretes.js",
        "variable_js": "productosAretes",
        "prefijo": "aretes",
        "carpeta": "imagenes/aretes"
    },
    {
        "categoria": "Collares",
        "excel": "COLLARES.xlsx",
        "js": "collares.js",
        "variable_js": "productosCollares",
        "prefijo": "collares",
        "carpeta": "imagenes/collares"
    },
    {
        "categoria": "Pulseras",
        "excel": "PULSERAS.xlsx",
        "js": "pulseras.js",
        "variable_js": "productosPulseras",
        "prefijo": "pulseras",
        "carpeta": "imagenes/pulseras"
    },
    {
        "categoria": "Charms Beads",
        "excel": "CHARMS BEADS.xlsx",
        "js": "charmsbeads.js",
        "variable_js": "productosCharmsBeads",
        "prefijo": "chb",
        "carpeta": "imagenes/charms_beads"
    },
    {
        "categoria": "Charms Beads Disney",
        "excel": "CHARMS BEADS DISNEY.xlsx",
        "js": "charmsbeadsdisney.js",
        "variable_js": "productosCharmsBeadsDisney",
        "prefijo": "chbd",
        "carpeta": "imagenes/charms_beads_disney"
    },
    {
        "categoria": "Charms Colgantes",
        "excel": "CHARMS COLGANTES.xlsx",
        "js": "charmscolgantes.js",
        "variable_js": "productosCharmsColgantes",
        "prefijo": "chc",
        "carpeta": "imagenes/charms_colgantes"
    },
    {
        "categoria": "Charms Colgantes Disney",
        "excel": "CHARMS COLGANTES DISNEY.xlsx",
        "js": "charmscolgantesdisney.js",
        "variable_js": "productosCharmsColgantesDisney",
        "prefijo": "chcd",
        "carpeta": "imagenes/charms_colgantes_disney"
    },
    {
        "categoria": "Charms Reflexions",
        "excel": "CHARMS REFLEXION.xlsx",
        "js": "charmsreflexion.js",
        "variable_js": "productosCharmsReflexion",
        "prefijo": "chr",
        "carpeta": "imagenes/charms_reflexion"
    },
    {
        "categoria": "Charms Muranos",
        "excel": "CHARMS MURANOS.xlsx",
        "js": "charmsmuranos.js",
        "variable_js": "productosCharmsMuranos",
        "prefijo": "chm",
        "carpeta": "imagenes/charms_muranos"
    },
    {
        "categoria": "Charms Cadenas de Seguridad",
        "excel": "CHARMS CADENAS DE SEGURIDAD.xlsx",
        "js": "charmscadenasseguridad.js",
        "variable_js": "productosCharmsCadenasSeguridad",
        "prefijo": "chcsd",
        "carpeta": "imagenes/charms_cadenasseguridad"
    },
    {
        "categoria": "Charms Clips y Topes",
        "excel": "CHARMS CLIPS Y TOPES.xlsx",
        "js": "charmsclipsytopes.js",
        "variable_js": "productosCharmsClipsYTopes",
        "prefijo": "chct",
        "carpeta": "imagenes/charms_clipsytopes"
    },
    {
        "categoria": "Charms Locket",
        "excel": "CHARMS LOCKET.xlsx",
        "js": "charmslocke.js",
        "variable_js": "productosCharmsLocket",
        "prefijo": "chl",
        "carpeta": "imagenes/charms_lockets"
    },
    {
        "categoria": "Charms ME",
        "excel": "CHARMS ME.xlsx",
        "js": "charmsme.js",
        "variable_js": "productosCharmsME",
        "prefijo": "chme",
        "carpeta": "imagenes/charms_me"
    },
    {
        "categoria": "Charms Accesorios ME",
        "excel": "CHARMS ACCESORIOS ME.xlsx",
        "js": "charmsaccesoriosme.js",
        "variable_js": "productosCharmsAccesoriosME",
        "prefijo": "chaME",
        "carpeta": "imagenes/charms_accesoriosme"
    },
    {
        "categoria": "Anillos Swarovski",
        "excel": "ANILLOS SWA.xlsx",
        "js": "anillosswa.js",
        "variable_js": "productosAnillosSwa",
        "prefijo": "anillos_swa",
        "carpeta": "imagenes/SWA/anillos_swa",
        "header_excel": 2
    },
    {
        "categoria": "Aretes Swarovski",
        "excel": "ARETES SWA.xlsx",
        "js": "aretesswa.js",
        "variable_js": "productosAretesSwa",
        "prefijo": "aretes_swa",
        "carpeta": "imagenes/SWA/aretes_swa",
        "header_excel": 3
    },
    {
        "categoria": "Pulseras Swarovski",
        "excel": "PULSERAS SWA.xlsx",
        "js": "pulserasswa.js",
        "variable_js": "productosPulserasSwa",
        "prefijo": "pulseras_swa",
        "carpeta": "imagenes/SWA/pulseras_swa",
        "header_excel": 2
    },
    {
        "categoria": "Collares Swarovski",
        "excel": "COLLARES SWA.xlsx",
        "js": "collarswa.js",
        "variable_js": "productosCollaresSwa",
        "prefijo": "collares_swa",
        "carpeta": "imagenes/SWA/collares_swa",
        "header_excel": 2
    },
    {
        "categoria": "Anillos Baño de Plata",
        "excel": "ANILLOS BP.xlsx",
        "js": "anillosbp.js",
        "variable_js": "productosAnillosBp",
        "prefijo": "anillosbp",
        "carpeta": "imagenes/BP/anillosbp",
        "header_excel": 2
    },
    {
        "categoria": "Aretes Baño de Plata",
        "excel": "ARETES BP.xlsx",
        "js": "aretesbp.js",
        "variable_js": "productosAretesBp",
        "prefijo": "aretessbp",
        "carpeta": "imagenes/BP/aretessbp",
        "header_excel": 3
    }
]

def buscar_archivo_flexible(nombre_buscado):
    if os.path.exists(nombre_buscado):
        return nombre_buscado
    for f in os.listdir("."):
        if f.lower() == nombre_buscado.lower():
            return f
    return None

def actualizar_todo():
    for conf in CONFIGURACIONES:
        cat = conf["categoria"]
        excel_esperado = conf["excel"]
        js_path = conf["js"]
        var_name = conf["variable_js"]
        prefijo = conf["prefijo"]
        carpeta = conf["carpeta"]
        header_fila = conf.get("header_excel", 0)

        print(f"\n--- Procesando categoría: {cat} ---")

        excel_path = buscar_archivo_flexible(excel_esperado)
        if not excel_path:
            print(f"⚠️ Aviso: No se encontró el archivo {excel_esperado}. Se omite esta categoría.")
            continue

        if not os.path.exists(carpeta):
            os.makedirs(carpeta, exist_ok=True)
            print(f"📁 Se creó la carpeta faltante: {carpeta}")

        try:
            df = pd.read_excel(excel_path, header=header_fila)
        except Exception as e:
            print(f"⚠️ Error al leer {excel_path}: {e}")
            continue

        productos_lista = []

        for index, row in df.iterrows():
            col_id = None
            for posible_col in ['NUMERO', 'NUM.', 'ID', 'Id', 'numero', 'Número', 'ARETES', 'ANILLOS', 'PULSERAS', 'COLLARES', 'CODIGO']:
                if posible_col in df.columns:
                    col_id = posible_col
                    break
            
            if not col_id or pd.isna(row[col_id]):
                continue
                
            try:
                prod_id = int(row[col_id])
            except:
                continue
            
            col_precio = None
            for posible_col in ['PRECIO', 'Precio', 'precio']:
                if posible_col in df.columns:
                    col_precio = posible_col
                    break

            precio = float(row[col_precio]) if col_precio and not pd.isna(row[col_precio]) else 0.0
            
            stock_tallas = {}
            for col in df.columns:
                col_upper = str(col).strip().upper()
                if col_upper.startswith('STOCK') or col_upper == 'TALLAS DISPONIBLES' or col_upper == 'TALLAS':
                    tallas_str = str(row[col])
                    if tallas_str and tallas_str != 'nan':
                        for t in tallas_str.replace(" ", "").split("-"):
                            if t:
                                stock_tallas[t] = 1

            archivos_encontrados = []
            if os.path.exists(carpeta):
                patron_estricto = re.compile(rf"^{re.escape(prefijo)}_{prod_id}(?:[\._\(]|$)", re.IGNORECASE)

                for archivo in os.listdir(carpeta):
                    if patron_estricto.match(archivo):
                        ext = archivo.split('.')[-1].lower()
                        if ext in ('jpg', 'jpeg', 'png', 'webp', 'avif', 'mp4', 'mov', 'webm'):
                            sub_part = archivo.replace(f"{prefijo}_{prod_id}", "").split('.')[0]
                            tupla_orden = (0,) if not sub_part else tuple([int(n) for n in re.findall(r'\d+', sub_part)] or [0])
                            
                            tipo_media = "video" if ext in ('mp4', 'mov', 'webm') else "imagen"
                            archivos_encontrados.append({
                                "archivo": archivo,
                                "tupla_orden": tupla_orden,
                                "tipo": tipo_media,
                                "ruta": f"{carpeta}/{archivo}"
                            })

            archivos_encontrados = sorted(archivos_encontrados, key=lambda x: x['tupla_orden'])

            galeria_items = []
            for item in archivos_encontrados:
                galeria_items.append({"tipo": item["tipo"], "url": item["ruta"]})

            if not galeria_items:
                imagen_principal = f"{carpeta}/{prefijo}_{prod_id}.jpg"
                galeria_items = [{"tipo": "imagen", "url": imagen_principal}]
            else:
                imagen_principal = galeria_items[0]["url"]

            producto_obj = {
                "id": prod_id,
                "categoria": cat,
                "precio": precio,
                "imagen": imagen_principal,
                "stockTallas": stock_tallas,
                "galeria": galeria_items
            }
            productos_lista.append(producto_obj)

        contenido_js = f"const {var_name} = {json.dumps(productos_lista, indent=4, ensure_ascii=False)};"
        with open(js_path, "w", encoding="utf-8") as f:
            f.write(contenido_js)

        print(f"✅ Archivo {js_path} actualizado correctamente con {len(productos_lista)} productos.")

    print("\n🎉 ¡Proceso completo de actualización finalizado!")

if __name__ == "__main__":
    actualizar_todo()