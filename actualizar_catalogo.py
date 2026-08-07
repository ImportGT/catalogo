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
    }
]

def actualizar_todo():
    for conf in CONFIGURACIONES:
        cat = conf["categoria"]
        excel_path = conf["excel"]
        js_path = conf["js"]
        var_name = conf["variable_js"]
        prefijo = conf["prefijo"]
        carpeta = conf["carpeta"]
        header_fila = conf.get("header_excel", 0)

        print(f"\n--- Procesando categoría: {cat} ---")

        if not os.path.exists(excel_path):
            print(f"⚠️ Aviso: No se encontró el archivo {excel_path}. Se omite esta categoría.")
            continue

        if not os.path.exists(carpeta):
            print(f"⚠️ Aviso: No se encontró la carpeta {carpeta}. Se omite esta categoría.")
            continue

        df = pd.read_excel(excel_path, header=header_fila)
        productos_lista = []

        for index, row in df.iterrows():
            col_id = None
            for posible_col in ['NUMERO', 'NUM.', 'ID', 'Id', 'numero', 'Número', 'ARETES', 'ANILLOS', 'PULSERAS', 'COLLARES']:
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
                if col_upper.startswith('STOCK') or col_upper == 'TALLAS':
                    if col_upper == 'TALLAS':
                        tallas_str = str(row[col])
                        if tallas_str and tallas_str != 'nan':
                            for t in tallas_str.replace(" ", "").split("--"):
                                if t:
                                    stock_tallas[t] = 1
                    else:
                        talla_nombre = col_upper.replace('STOCK', '').strip()
                        if talla_nombre:
                            val_stock = row[col]
                            if not pd.isna(val_stock):
                                try:
                                    stock_val = int(val_stock)
                                    if stock_val > 0:
                                        stock_tallas[talla_nombre] = stock_val
                                except:
                                    pass

            patron_archivo = re.compile(rf"^{re.escape(prefijo)}[\._]{prod_id}(?:[\._](\d+))?\.([a-zA-Z0-9]+)$", re.IGNORECASE)

            archivos_encontrados = []
            for archivo in os.listdir(carpeta):
                match = patron_archivo.match(archivo)
                if match:
                    sub_val = match.group(1)
                    ext = match.group(2).lower()
                    orden = 0 if sub_val is None else int(sub_val)
                    archivos_encontrados.append({
                        "archivo": archivo,
                        "orden": orden,
                        "ext": ext,
                        "ruta": f"{carpeta}/{archivo}"
                    })

            archivos_encontrados = sorted(archivos_encontrados, key=lambda x: x['orden'])

            galeria_items = []
            for item in archivos_encontrados:
                if item["ext"] in ('mp4', 'mov', 'webm'):
                    galeria_items.append({"tipo": "video", "url": item["ruta"]})
                elif item["ext"] in ('jpg', 'jpeg', 'png', 'webp'):
                    galeria_items.append({"tipo": "imagen", "url": item["ruta"]})

            imagenes_solas = [item for item in galeria_items if item['tipo'] == 'imagen']
            if imagenes_solas:
                imagen_principal = imagenes_solas[0]['url']
            elif galeria_items:
                imagen_principal = galeria_items[0]['url']
            else:
                imagen_principal = f"{carpeta}/{prefijo}_{prod_id}.jpg"

            producto_obj = {
                "id": prod_id,
                "categoria": cat,
                "precio": precio,
                "imagen": imagen_principal,
                "stockTallas": stock_tallas
            }
            
            if len(galeria_items) > 1:
                producto_obj["galeria"] = galeria_items
            else:
                producto_obj["galeria"] = [{"tipo": "imagen", "url": imagen_principal}]
                
            productos_lista.append(producto_obj)

        contenido_js = f"const {var_name} = {json.dumps(productos_lista, indent=4, ensure_ascii=False)};"

        with open(js_path, "w", encoding="utf-8") as f:
            f.write(contenido_js)

        print(f"✅ Archivo {js_path} actualizado correctamente con {len(productos_lista)} productos.")

    print("\n🎉 ¡Proceso completo de actualización finalizado!")

if __name__ == "__main__":
    actualizar_todo()