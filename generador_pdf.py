import os
import pandas as pd
from PIL import Image
from fpdf import FPDF

class PDFCatalogoUnificado(FPDF):
    def __init__(self, titulo_general, lista_categorias_cortas):
        super().__init__(orientation='P', unit='mm', format='A4')
        self.titulo_general = titulo_general
        self.lista_categorias_cortas = lista_categorias_cortas
        self.link_indice = None

    def header(self):
        pass

    def footer(self):
        if self.page_no() > 1:
            self.set_y(-18)
            self.set_font('helvetica', '', 7)
            self.set_draw_color(210, 210, 210)
            self.line(12, self.get_y(), 198, self.get_y())
            self.ln(2)
            
            mitad = (len(self.lista_categorias_cortas) + 1) // 2
            fila_1 = self.lista_categorias_cortas[:mitad]
            fila_2 = self.lista_categorias_cortas[mitad:]

            self.set_text_color(100, 100, 100)
            self.cell(10, 4, "Ir:", align='L')
            
            if self.link_indice:
                self.set_text_color(31, 78, 121)
                self.set_font('helvetica', 'B', 7)
                self.cell(14, 4, "[Índice]", link=self.link_indice, align='L')
            
            self.set_font('helvetica', '', 7)
            for nombre_corto, l_obj in fila_1:
                self.set_text_color(0, 102, 204)
                self.cell(len(nombre_corto) * 2.1 + 3, 4, f"[{nombre_corto}]", link=l_obj, align='L')

            self.set_text_color(140, 140, 140)
            self.set_xy(180, self.get_y())
            self.cell(18, 4, f"Pág. {self.page_no()}", align='R')

            if fila_2:
                self.ln(4)
                self.set_x(12)
                self.cell(24, 4, "", align='L')
                for nombre_corto, l_obj in fila_2:
                    self.set_text_color(0, 102, 204)
                    self.cell(len(nombre_corto) * 2.1 + 3, 4, f"[{nombre_corto}]", link=l_obj, align='L')

def generar_pdf_desde_excel(ruta_excel, nombre_linea, coleccion_carpeta, prefijo_archivo, carpeta_imagenes, mostrar_precios=True, margen_porcentaje=0):
    subcarpeta_tipo = "Catálogo PDF precios" if mostrar_precios else "Catálogo PDF"
    carpeta_destino = os.path.join("catalogos", coleccion_carpeta, subcarpeta_tipo)
    
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)

    sufijo_precio = "p" if mostrar_precios else ""
    nombre_archivo = os.path.join(carpeta_destino, f"{prefijo_archivo}{sufijo_precio}.pdf")

    if not os.path.exists(ruta_excel):
        print(f"\n[Aviso] El archivo '{ruta_excel}' no se encuentra en el directorio.")
        return

    try:
        if coleccion_carpeta.lower() == "swarovski":
            df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
        elif coleccion_carpeta.lower() == "bano_de_plata":
            df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
        else:
            df = pd.read_excel(ruta_excel, header=0 if "Aretes" in nombre_linea else 0)
        df.columns = [str(c).strip().upper() for c in df.columns]
    except Exception as e:
        print(f"\n[Error] No se pudo leer el archivo Excel '{ruta_excel}': {e}")
        return

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()
    
    startX = 20
    startY = 20
    cols = 2
    cardWidth = 82
    cardHeight = 95 if mostrar_precios else 85
    gapX = 6
    gapY = 8
    
    currentX = startX
    currentY = startY
    countInRow = 0

    if df.empty:
        return

    for index, row in df.iterrows():
        raw_id = None
        for col_id in ['NUMERO', 'NUM.', 'CODIGO', 'CÓDIGO']:
            if col_id in row and not pd.isna(row[col_id]):
                raw_id = row[col_id]
                break
        if raw_id is None:
            raw_id = row.iloc[0]

        try:
            prod_id = int(float(raw_id))
        except (ValueError, TypeError):
            continue

        precio_base = 0.0
        for col_p in ['PRECIO', 'VENTA MAYORISTA', 'PRECIO MAYORISTA', 'PRECIO DOBLE']:
            if col_p in row and not pd.isna(row[col_p]):
                try:
                    precio_base = float(row[col_p])
                    break
                except (ValueError, TypeError):
                    pass

        precio_final = precio_base * (1 + (margen_porcentaje / 100))

        imagen_path = ""
        separador = "_"
        
        if coleccion_carpeta.lower() == "swarovski":
            palabra_clave = "anillos_swa" if "Anillos" in nombre_linea else ("aretes_swa" if "Aretes" in nombre_linea else ("pulseras_swa" if "Pulseras" in nombre_linea else "collares_swa"))
        elif coleccion_carpeta.lower() == "bano_de_plata":
            palabra_clave = "anillosbp" if "Anillos" in nombre_linea else ("aretessbp" if "Aretes" in nombre_linea else ("cadenasbp" if "Cadenas" in nombre_linea else "pulserasbp"))
        elif coleccion_carpeta.lower() == "pandora":
            if "Aretes" in nombre_linea: palabra_clave = "aretes"
            elif "Anillos" in nombre_linea: palabra_clave = "anillos"
            elif "Collares" in nombre_linea: palabra_clave = "collares"
            elif "Pulseras" in nombre_linea: palabra_clave = "pulseras"
            elif "Charms Accesorios ME" in nombre_linea: palabra_clave = "chamE"
            elif "Charms ME" in nombre_linea: palabra_clave = "chme"
            elif "Charms Reflexion" in nombre_linea: palabra_clave = "chr"
            elif "Charms Locket" in nombre_linea: palabra_clave = "chl"
            elif "Charms Clips y Topes" in nombre_linea: palabra_clave = "chct"
            elif "Charms Muranos" in nombre_linea: palabra_clave = "chm"
            elif "Charms Cadenas de Seguridad" in nombre_linea: palabra_clave = "chcsd"
            elif "Charms Beads Disney" in nombre_linea: palabra_clave = "chbd"
            elif "Charms Colgantes Disney" in nombre_linea: palabra_clave = "chcd"
            elif "Charms Beads" in nombre_linea: palabra_clave = "chb"
            elif "Charms Colgantes" in nombre_linea: 
                palabra_clave = "chc"
                separador = "."
            else:
                palabra_clave = nombre_linea.split()[0].lower()
        else:
            palabra_clave = nombre_linea.split()[0].lower()
            
        extensiones = ['.jpg', '.jpeg', '.avif', '.webp', '.png']
        patrones_sufijos = ['', '.1', '.0', '_1', '_0', '.2', '.3', '.4', '.5', '_2', '_3', '_4', '_5']
        
        posibles_rutas = []
        for suf in patrones_sufijos:
            for ext in extensiones:
                posibles_rutas.append(f"{carpeta_imagenes}/{palabra_clave}{separador}{prod_id}{suf}{ext}")

        for ruta in posibles_rutas:
            if os.path.exists(ruta):
                imagen_path = ruta
                break

        if currentY + cardHeight > 270:
            pdf.add_page()
            currentY = 20
            currentX = startX
            countInRow = 0

        pdf.set_draw_color(220, 220, 220)
        pdf.set_fill_color(255, 255, 255)
        pdf.rect(currentX, currentY, cardWidth, cardHeight, 'DF')

        img_w_std = 60
        img_h_std = 60
        img_x = currentX + (cardWidth - img_w_std) / 2
        img_y = currentY + 5

        # Enlace directo optimizado incluyendo el parámetro de visibilidad de precios (p=1 o p=0)
        nombre_categoria_limpio = nombre_linea.lower().replace(' ', '_')
        estado_precio = 1 if mostrar_precios else 0
        enlace_producto_web = f"https://importgt.github.io/catalogo/?prod={prod_id}&cat={nombre_categoria_limpio}&p={estado_precio}"
        pdf.link(img_x, img_y, img_w_std, img_h_std, enlace_producto_web)

        if imagen_path and os.path.exists(str(imagen_path)):
            try:
                pdf.image(str(imagen_path), x=img_x, y=img_y, w=img_w_std, h=img_h_std)
            except Exception:
                pass

        pdf.set_font('helvetica', 'B', 11)
        pdf.set_text_color(44, 62, 80)
        pdf.set_xy(currentX, currentY + 68)
        pdf.cell(cardWidth, 6, f"Cod. {prod_id}", new_x="RIGHT", new_y="TOP", align='C')

        if mostrar_precios:
            pdf.set_font('helvetica', 'B', 12)
            pdf.set_text_color(0, 128, 0)
            pdf.set_xy(currentX, currentY + 76)
            pdf.cell(cardWidth, 6, f"Q{precio_final:.2f}", new_x="RIGHT", new_y="TOP", align='C')

        countInRow += 1
        if countInRow >= cols:
            countInRow = 0
            currentX = startX
            currentY += cardHeight + gapY
        else:
            currentX += cardWidth + gapX

    pdf.output(nombre_archivo)
    print(f"\n-> ¡Éxito! Catálogo guardado en: {nombre_archivo}\n")

def generar_pdf_unificado(tareas, coleccion_carpeta, mostrar_precios=True, margen_porcentaje=0):
    subcarpeta_tipo = "Catálogo PDF precios" if mostrar_precios else "Catálogo PDF"
    carpeta_destino = os.path.join("catalogos", coleccion_carpeta, subcarpeta_tipo)
    
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)

    sufijo_precio = "p" if mostrar_precios else ""
    nombre_archivo = os.path.join(carpeta_destino, f"catalogo_general_{coleccion_carpeta.lower()}{sufijo_precio}.pdf")

    def obtener_nombre_corto(nombre_largo):
        l = nombre_largo.lower()
        if "disney" in l: return "Disney Beads" if "beads" in l else "Disney Colg."
        if "seguridad" in l: return "Seguridad"
        if "accesorios" in l: return "Acc. ME"
        if "reflexion" in l: return "Reflexion"
        if "locket" in l: return "Lockets"
        if "topes" in l: return "Clips/Topes"
        if "muranos" in l: return "Muranos"
        partes = nombre_largo.split()
        if len(partes) > 1:
            if partes[0].lower() == "charms": return partes[1]
            elif partes[1].lower() in ["pandora", "swarovski", "baño", "de", "plata"]: return partes[0]
            else: return f"{partes[0]} {partes[1]}"
        return partes[0]

    link_objs = {}
    lista_links_tuplas = []
    
    for _, nombre_linea, _ in tareas:
        link_objs[nombre_linea] = None

    pdf = PDFCatalogoUnificado(coleccion_carpeta, lista_links_tuplas)
    pdf.set_auto_page_break(auto=True, margin=18)
    
    pdf.add_page()
    link_indice_obj = pdf.add_link()
    pdf.link_indice = link_indice_obj
    pdf.set_link(link_indice_obj, y=0)

    pdf.set_font('helvetica', 'B', 20)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 10, f"CATÁLOGO EXCLUSIVO", new_x="LMARGIN", new_y="NEXT", align='C')
    
    pdf.set_font('helvetica', 'B', 14)
    pdf.set_text_color(90, 100, 110)
    pdf.cell(0, 8, f"Colección: {coleccion_carpeta.upper().replace('_', ' ')}", new_x="LMARGIN", new_y="NEXT", align='C')
    pdf.ln(6)

    pdf.set_draw_color(31, 78, 121)
    pdf.set_line_width(0.8)
    pdf.line(55, pdf.get_y(), 155, pdf.get_y())
    pdf.ln(10)

    pdf.set_font('helvetica', '', 10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 6, "Seleccione una categoría o navegue mediante el menú inferior en cada página:", new_x="LMARGIN", new_y="NEXT", align='C')
    pdf.ln(8)

    for _, nombre_linea, _ in tareas:
        link_objs[nombre_linea] = pdf.add_link()
        nombre_corto = obtener_nombre_corto(nombre_linea)
        lista_links_tuplas.append((nombre_corto, link_objs[nombre_linea]))

    idx_card_w = 85
    idx_card_h = 16
    idx_start_x = 20
    idx_gap_x = 10
    idx_gap_y = 6
    
    current_x = idx_start_x
    current_y = pdf.get_y()
    i_count = 0

    for _, nombre_linea, _ in tareas:
        pdf.set_draw_color(200, 210, 220)
        pdf.set_fill_color(248, 250, 252)
        pdf.rect(current_x, current_y, idx_card_w, idx_card_h, 'DF')

        pdf.set_font('helvetica', 'B', 11)
        pdf.set_text_color(31, 78, 121)
        pdf.set_xy(current_x + 5, current_y + 4)
        pdf.cell(idx_card_w - 10, 8, f">>  {nombre_linea}", link=link_objs[nombre_linea], align='L')

        i_count += 1
        if i_count % 2 == 0:
            current_x = idx_start_x
            current_y += idx_card_h + idx_gap_y
        else:
            current_x = idx_start_x + idx_card_w + idx_gap_x

    datos_lineas = []
    for ruta_excel, nombre_linea, carpeta_imagenes in tareas:
        if not os.path.exists(ruta_excel):
            datos_lineas.append((pd.DataFrame(), nombre_linea, carpeta_imagenes))
            continue
        try:
            if coleccion_carpeta.lower() == "swarovski":
                df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
            elif coleccion_carpeta.lower() == "bano_de_plata":
                df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
            else:
                df = pd.read_excel(ruta_excel, header=0 if "Aretes" in nombre_linea else 0)
            df.columns = [str(c).strip().upper() for c in df.columns]
        except Exception:
            df = pd.DataFrame()
        datos_lineas.append((df, nombre_linea, carpeta_imagenes))

    startX = 20
    startY = 20
    cols = 2
    cardWidth = 82
    cardHeight = 95 if mostrar_precios else 85
    gapX = 6
    gapY = 8

    for df, nombre_linea, carpeta_imagenes in datos_lineas:
        if df.empty:
            continue

        pdf.add_page()
        pdf.set_link(link_objs[nombre_linea], y=0)

        currentX = startX
        currentY = startY
        countInRow = 0

        pdf.set_font('helvetica', 'B', 14)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(0, 8, f"Categoría: {nombre_linea}", new_x="LMARGIN", new_y="NEXT", align='L')
        pdf.ln(2)
        currentY += 12

        for index, row in df.iterrows():
            raw_id = None
            for col_id in ['NUMERO', 'NUM.', 'CODIGO', 'CÓDIGO']:
                if col_id in row and not pd.isna(row[col_id]):
                    raw_id = row[col_id]
                    break
            if raw_id is None:
                raw_id = row.iloc[0]

            try:
                prod_id = int(float(raw_id))
            except (ValueError, TypeError):
                continue

            precio_base = 0.0
            for col_p in ['PRECIO', 'VENTA MAYORISTA', 'PRECIO MAYORISTA', 'PRECIO DOBLE']:
                if col_p in row and not pd.isna(row[col_p]):
                    try:
                        precio_base = float(row[col_p])
                        break
                    except (ValueError, TypeError):
                        pass

            precio_final = precio_base * (1 + (margen_porcentaje / 100))

            imagen_path = ""
            separador = "_"
            if coleccion_carpeta.lower() == "swarovski":
                palabra_clave = "anillos_swa" if "Anillos" in nombre_linea else ("aretes_swa" if "Aretes" in nombre_linea else ("pulseras_swa" if "Pulseras" in nombre_linea else "collares_swa"))
            elif coleccion_carpeta.lower() == "bano_de_plata":
                palabra_clave = "anillosbp" if "Anillos" in nombre_linea else ("aretessbp" if "Aretes" in nombre_linea else ("cadenasbp" if "Cadenas" in nombre_linea else "pulserasbp"))
            elif coleccion_carpeta.lower() == "pandora":
                if "Aretes" in nombre_linea: palabra_clave = "aretes"
                elif "Anillos" in nombre_linea: palabra_clave = "anillos"
                elif "Collares" in nombre_linea: palabra_clave = "collares"
                elif "Pulseras" in nombre_linea: palabra_clave = "pulseras"
                elif "Charms Accesorios ME" in nombre_linea: palabra_clave = "chamE"
                elif "Charms ME" in nombre_linea: palabra_clave = "chme"
                elif "Charms Reflexion" in nombre_linea: palabra_clave = "chr"
                elif "Charms Locket" in nombre_linea: palabra_clave = "chl"
                elif "Charms Clips y Topes" in nombre_linea: palabra_clave = "chct"
                elif "Charms Muranos" in nombre_linea: palabra_clave = "chm"
                elif "Charms Cadenas de Seguridad" in nombre_linea: palabra_clave = "chcsd"
                elif "Charms Beads Disney" in nombre_linea: palabra_clave = "chbd"
                elif "Charms Colgantes Disney" in nombre_linea: palabra_clave = "chcd"
                elif "Charms Beads" in nombre_linea: palabra_clave = "chb"
                elif "Charms Colgantes" in nombre_linea: 
                    palabra_clave = "chc"
                    separador = "."
                else:
                    palabra_clave = nombre_linea.split()[0].lower()
            else:
                palabra_clave = nombre_linea.split()[0].lower()
                
            extensiones = ['.jpg', '.jpeg', '.avif', '.webp', '.png']
            patrones_sufijos = ['', '.1', '.0', '_1', '_0', '.2', '.3', '.4', '.5', '_2', '_3', '_4', '_5']
            
            posibles_rutas = []
            for suf in patrones_sufijos:
                for ext in extensiones:
                    posibles_rutas.append(f"{carpeta_imagenes}/{palabra_clave}{separador}{prod_id}{suf}{ext}")

            for ruta in posibles_rutas:
                if os.path.exists(ruta):
                    imagen_path = ruta
                    break

            if currentY + cardHeight > 270:
                pdf.add_page()
                currentY = 20
                currentX = startX
                countInRow = 0

            pdf.set_draw_color(220, 220, 220)
            pdf.set_fill_color(255, 255, 255)
            pdf.rect(currentX, currentY, cardWidth, cardHeight, 'DF')

            img_w_std = 60
            img_h_std = 60
            img_x = currentX + (cardWidth - img_w_std) / 2
            img_y = currentY + 5

            nombre_categoria_limpio = nombre_linea.lower().replace(' ', '_')
            estado_precio = 1 if mostrar_precios else 0
            enlace_producto_web = f"https://importgt.github.io/catalogo/?prod={prod_id}&cat={nombre_categoria_limpio}&p={estado_precio}"
            pdf.link(img_x, img_y, img_w_std, img_h_std, enlace_producto_web)

            if imagen_path and os.path.exists(str(imagen_path)):
                try:
                    pdf.image(str(imagen_path), x=img_x, y=img_y, w=img_w_std, h=img_h_std)
                except Exception:
                    pass

            pdf.set_font('helvetica', 'B', 11)
            pdf.set_text_color(44, 62, 80)
            pdf.set_xy(currentX, currentY + 68)
            pdf.cell(cardWidth, 6, f"Cod. {prod_id}", new_x="RIGHT", new_y="TOP", align='C')

            if mostrar_precios:
                pdf.set_font('helvetica', 'B', 12)
                pdf.set_text_color(0, 128, 0)
                pdf.set_xy(currentX, currentY + 76)
                pdf.cell(cardWidth, 6, f"Q{precio_final:.2f}", new_x="RIGHT", new_y="TOP", align='C')

            countInRow += 1
            if countInRow >= cols:
                countInRow = 0
                currentX = startX
                currentY += cardHeight + gapY
            else:
                currentX += cardWidth + gapX

    pdf.output(nombre_archivo)
    print(f"\n-> ¡Éxito! Catálogo guardado en: {nombre_archivo}\n")

if __name__ == "__main__":
    while True:
        print("=" * 55)
        print("     ASISTENTE DE GENERACIÓN DE CATÁLOGOS PDF")
        print("=" * 55)
        print("Selecciona la Colección:")
        print("1. Pandora")
        print("2. Swarovski")
        print("3. Baño de Plata")
        print("4. Salir")
        
        coleccion_op = input("\nElige una opción (1-4): ").strip()
        
        if coleccion_op == "1":
            coleccion_nombre = "Pandora"
            tareas = [
                ("ANILLOS.xlsx", "Anillos Pandora", "imagenes/anillos"),
                ("ARETES.xlsx", "Aretes Pandora", "imagenes/aretes"),
                ("COLLARES.xlsx", "Collares Pandora", "imagenes/collares"),
                ("PULSERAS.xlsx", "Pulseras Pandora", "imagenes/pulseras"),
                ("CHARMS BEADS.xlsx", "Charms Beads Pandora", "imagenes/charms_beads"),
                ("CHARMS COLGANTES.xlsx", "Charms Colgantes Pandora", "imagenes/charms_colgantes"),
                ("CHARMS BEADS DISNEY.xlsx", "Charms Beads Disney Pandora", "imagenes/charms_beads_disney"),
                ("CHARMS COLGANTES DISNEY.xlsx", "Charms Colgantes Disney Pandora", "imagenes/charms_colgantes_disney"),
                ("CHARMS CADENAS DE SEGURIDAD.xlsx", "Charms Cadenas de Seguridad Pandora", "imagenes/charms_cadenasseguridad"),
                ("CHARMS MURANOS.xlsx", "Charms Muranos Pandora", "imagenes/charms_muranos"),
                ("CHARMS CLIPS Y TOPES.xlsx", "Charms Clips y Topes Pandora", "imagenes/charms_clipsytopes"),
                ("CHARMS LOCKET.xlsx", "Charms Locket Pandora", "imagenes/charms_lockets"),
                ("CHARMS REFLEXION.xlsx", "Charms Reflexion Pandora", "imagenes/charms_reflexion"),
                ("CHARMS ME.xlsx", "Charms ME Pandora", "imagenes/charms_me"),
                ("CHARMS ACCESORIOS ME.xlsx", "Charms Accesorios ME Pandora", "imagenes/charms_accesoriosme")
            ]
        elif coleccion_op == "2":
            coleccion_nombre = "Swarovski"
            tareas = [
                ("ANILLOS SWA.xlsx", "Anillos Swarovski", "imagenes/SWA/anillos_swa"),
                ("ARETES SWA.xlsx", "Aretes Swarovski", "imagenes/SWA/aretes_swa"),
                ("PULSERAS SWA.xlsx", "Pulseras Swarovski", "imagenes/SWA/pulseras_swa"),
                ("COLLARES SWA.xlsx", "Collares Swarovski", "imagenes/SWA/collares_swa")
            ]
        elif coleccion_op == "3":
            coleccion_nombre = "Bano_de_Plata"
            tareas = [
                ("ANILLOS BP.xlsx", "Anillos Baño de Plata", "imagenes/BP/anillosbp"),
                ("ARETES BP.xlsx", "Aretes Baño de Plata", "imagenes/BP/aretessbp"),
                ("CADENAS BP.xlsx", "Cadenas Baño de Plata", "imagenes/BP/cadenasbp"),
                ("PULSERAS BP.xlsx", "Pulseras Baño de Plata", "imagenes/BP/pulserasbp")
            ]
        elif coleccion_op == "4" or coleccion_op.lower() == "salir":
            print("Saliendo del asistente...")
            break
        else:
            print("\nOpción no válida. Inténtalo de nuevo.\n")
            continue

        print(f"\n--- Opciones para {coleccion_nombre} ---")
        print("1. Generar Catálogo UNIFICADO")
        print("2. Generar una línea específica directamente")
        
        tipo_gen = input("\nSelecciona una opción: ").strip()

        resp_precios = input("¿Deseas incluir precios en el catálogo? (s/n): ").strip().lower()
        mostrar_precios = resp_precios in ["s", "si", "sí", "y", "yes"]

        margen_porcentaje = 0.0
        if mostrar_precios:
            try:
                resp_margen = input("Ingresa el porcentaje de margen extra (ej. 0 para base, 15 para +15%): ").strip()
                margen_porcentaje = float(resp_margen) if resp_margen else 0.0
            except ValueError:
                margen_porcentaje = 0.0

        if tipo_gen == "1":
            print(f"\nProcesando catálogo unificado para {coleccion_nombre}...")
            generar_pdf_unificado(tareas, coleccion_nombre, mostrar_precios=mostrar_precios, margen_porcentaje=margen_porcentaje)
        elif tipo_gen == "2":
            print("\n--- Selecciona la línea que deseas generar ---")
            for idx, (_, linea, _) in enumerate(tareas, 1):
                print(f"{idx}. {linea}")
            
            sub_op = input(f"\nElige una opción (1-{len(tareas)}): ").strip()
            
            try:
                idx_sel = int(sub_op) - 1
                if 0 <= idx_sel < len(tareas):
                    excel, linea, imagenes = tareas[idx_sel]
                    prefijo = linea.lower().replace(" ", "_")
                    print(f"\nGenerando PDF para {linea}...")
                    generar_pdf_desde_excel(
                        ruta_excel=excel,
                        nombre_linea=linea,
                        coleccion_carpeta=coleccion_nombre,
                        prefijo_archivo=prefijo,
                        carpeta_imagenes=imagenes,
                        mostrar_precios=mostrar_precios,
                        margen_porcentaje=margen_porcentaje
                    )
                else:
                    print("Opción no válida.")
            except ValueError:
                print("Entrada no válida.")
        else:
            print("Opción no válida.")

        otro = input("¿Deseas realizar otra operación? (s/n): ").strip().lower()
        if otro not in ["s", "si", "sí", "y", "yes"]:
            print("¡Hasta luego!")
            break