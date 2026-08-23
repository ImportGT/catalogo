import os
import pandas as pd
from PIL import Image
from fpdf import FPDF

class PDFCatalogoJerarquico(FPDF):
    def __init__(self, lista_colecciones_links, lista_lineas_links, coleccion_actual):
        super().__init__(orientation='P', unit='mm', format='A4')
        self.lista_colecciones_links = lista_colecciones_links
        self.lista_lineas_links = lista_lineas_links
        self.coleccion_actual = coleccion_actual
        self.link_indice_general = None
        self.link_indice_coleccion = None

    def header(self):
        pass

    def footer(self):
        if self.page_no() > 1:
            self.set_y(-15)
            self.set_font('helvetica', '', 6.5)
            self.set_draw_color(210, 210, 210)
            self.line(12, self.get_y(), 198, self.get_y())
            self.ln(1.5)

            self.set_text_color(100, 100, 100)
            self.cell(6, 3.5, "Ir:", align='L')
            
            if self.link_indice_general:
                self.set_text_color(120, 40, 140)
                self.set_font('helvetica', 'B', 6.5)
                self.cell(16, 3.5, "[Ind. Gral.]", link=self.link_indice_general, align='L')

            if self.link_indice_coleccion:
                self.set_text_color(31, 78, 121)
                self.set_font('helvetica', 'B', 6.5)
                self.cell(16, 3.5, "[Subíndice]", link=self.link_indice_coleccion, align='L')

            self.set_font('helvetica', '', 6.5)
            for nombre_corto, l_obj in self.lista_lineas_links:
                self.set_text_color(0, 102, 204)
                self.cell(len(nombre_corto) * 1.8 + 2, 3.5, f"[{nombre_corto}]", link=l_obj, align='L')

            for col_nombre, l_col in self.lista_colecciones_links:
                if col_nombre != self.coleccion_actual:
                    self.set_text_color(180, 80, 0)
                    self.cell(len(col_nombre) * 1.8 + 2, 3.5, f"[{col_nombre}]", link=l_col, align='L')

            self.set_text_color(140, 140, 140)
            self.set_xy(182, self.get_y())
            self.cell(16, 3.5, f"Pág. {self.page_no()}", align='R')

def generar_pdf_estructurado(diccionario_colecciones, mostrar_precios=True, margen_porcentaje=0, coleccion_seleccionada=None):
    subcarpeta_tipo = "Catálogo PDF precios" if mostrar_precios else "Catálogo PDF"
    nombre_archivo_destino = coleccion_seleccionada if coleccion_seleccionada else "Maestro_General"
    carpeta_destino = os.path.join("catalogos", nombre_archivo_destino, subcarpeta_tipo)
    
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)

    sufijo_precio = "p" if mostrar_precios else ""
    nombre_archivo = os.path.join(carpeta_destino, f"catalogo_{nombre_archivo_destino.lower()}{sufijo_precio}.pdf")

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

    pdf = PDFCatalogoJerarquico([], [], "General")
    pdf.set_auto_page_break(auto=True, margin=18)

    # PÁGINA 1: ÍNDICE GENERAL DE COLECCIONES
    pdf.add_page()
    link_gen_obj = pdf.add_link()
    pdf.link_indice_general = link_gen_obj
    pdf.set_link(link_gen_obj, y=0)

    pdf.set_font('helvetica', 'B', 22)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 12, "CATÁLOGO DE JOYERÍA", new_x="LMARGIN", new_y="NEXT", align='C')
    
    pdf.set_font('helvetica', 'B', 14)
    pdf.set_text_color(90, 100, 110)
    pdf.cell(0, 8, "Índice General de Colecciones", new_x="LMARGIN", new_y="NEXT", align='C')
    pdf.ln(4)

    pdf.set_draw_color(31, 78, 121)
    pdf.set_line_width(0.8)
    pdf.line(55, pdf.get_y(), 155, pdf.get_y())
    pdf.ln(8)

    links_colecciones_maestro = {}
    for nombre_col in diccionario_colecciones.keys():
        links_colecciones_maestro[nombre_col] = pdf.add_link()

    pdf.lista_colecciones_links = [(k, v) for k, v in links_colecciones_maestro.items()]

    for nombre_col in diccionario_colecciones.keys():
        pdf.set_draw_color(200, 210, 220)
        pdf.set_fill_color(240, 244, 248)
        pdf.rect(35, pdf.get_y(), 140, 16, 'DF')
        
        pdf.set_font('helvetica', 'B', 12)
        pdf.set_text_color(31, 78, 121)
        pdf.set_xy(40, pdf.get_y() + 4.5)
        pdf.cell(130, 7, f">> Colección: {nombre_col}", link=links_colecciones_maestro[nombre_col], align='L')
        pdf.ln(20)

    # PROCESAMIENTO DE CADA COLECCIÓN Y SUS SUBÍNDICES
    for col_nombre, tareas_col in diccionario_colecciones.items():
        pdf.add_page()
        link_sub_obj = pdf.add_link()
        pdf.link_indice_coleccion = link_sub_obj
        pdf.set_link(link_sub_obj, y=0)
        pdf.set_link(links_colecciones_maestro[col_nombre], y=0)

        pdf.coleccion_actual = col_nombre

        pdf.set_font('helvetica', 'B', 20)
        pdf.set_text_color(31, 78, 121)
        pdf.cell(0, 10, f"Colección: {col_nombre}", new_x="LMARGIN", new_y="NEXT", align='C')
        
        pdf.set_font('helvetica', 'B', 13)
        pdf.set_text_color(90, 100, 110)
        pdf.cell(0, 7, "Subíndice de Categorías", new_x="LMARGIN", new_y="NEXT", align='C')
        pdf.ln(4)

        pdf.set_draw_color(31, 78, 121)
        pdf.set_line_width(0.6)
        pdf.line(55, pdf.get_y(), 155, pdf.get_y())
        pdf.ln(6)

        links_lineas_col = {}
        for _, nombre_linea, _ in tareas_col:
            links_lineas_col[nombre_linea] = pdf.add_link()

        pdf.lista_lineas_links = [(obtener_nombre_corto(n), l) for n, l in links_lineas_col.items()]

        idx_card_w = 85
        idx_card_h = 12
        idx_start_x = 20
        idx_gap_x = 10
        idx_gap_y = 4
        
        current_x = idx_start_x
        current_y = pdf.get_y()
        i_count = 0

        for _, nombre_linea, _ in tareas_col:
            pdf.set_draw_color(200, 210, 220)
            pdf.set_fill_color(248, 250, 252)
            pdf.rect(current_x, current_y, idx_card_w, idx_card_h, 'DF')

            pdf.set_font('helvetica', 'B', 9.5)
            pdf.set_text_color(31, 78, 121)
            pdf.set_xy(current_x + 3, current_y + 3)
            pdf.cell(idx_card_w - 6, 6, f">> {nombre_linea}", link=links_lineas_col[nombre_linea], align='L')

            i_count += 1
            if i_count % 2 == 0:
                current_x = idx_start_x
                current_y += idx_card_h + idx_gap_y
            else:
                current_x = idx_start_x + idx_card_w + idx_gap_x

        # Renderizar productos de las líneas de esta colección
        for ruta_excel, nombre_linea, carpeta_imagenes in tareas_col:
            if not os.path.exists(ruta_excel):
                continue
            try:
                if "swarovski" in col_nombre.lower() or "swa" in carpeta_imagenes.lower():
                    df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
                elif "baño" in col_nombre.lower() or "bano" in col_nombre.lower() or "bp" in carpeta_imagenes.lower():
                    df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
                else:
                    df = pd.read_excel(ruta_excel, header=0 if "Aretes" in nombre_linea else 0)
                df.columns = [str(c).strip().upper() for c in df.columns]
            except Exception:
                continue

            if df.empty:
                continue

            pdf.add_page()
            pdf.set_link(links_lineas_col[nombre_linea], y=0)

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
                img_folder = carpeta_imagenes.lower()
                separador_char = "_"
                
                if "swa" in img_folder:
                    if "Anillos" in nombre_linea: prefijo_base = "anillos_swa"
                    elif "Aretes" in nombre_linea: prefijo_base = "aretes_swa"
                    elif "Pulseras" in nombre_linea: prefijo_base = "pulseras_swa"
                    else: prefijo_base = "collares_swa"
                elif "bp" in img_folder:
                    if "Anillos" in nombre_linea: prefijo_base = "anillosbp"
                    elif "Aretes" in nombre_linea: prefijo_base = "aretessbp"
                    elif "Cadenas" in nombre_linea: prefijo_base = "cadenasbp"
                    else: prefijo_base = "pulserasbp"
                else:
                    if "Aretes" in nombre_linea: prefijo_base = "aretes"
                    elif "Anillos" in nombre_linea: prefijo_base = "anillos"
                    elif "Collares" in nombre_linea: prefijo_base = "collares"
                    elif "Pulseras" in nombre_linea: prefijo_base = "pulseras"
                    elif "Charms Accesorios ME" in nombre_linea: prefijo_base = "chamE"
                    elif "Charms ME" in nombre_linea: prefijo_base = "chme"
                    elif "Charms Reflexion" in nombre_linea: prefijo_base = "chr"
                    elif "Charms Locket" in nombre_linea: prefijo_base = "chl"
                    elif "Charms Clips y Topes" in nombre_linea: prefijo_base = "chct"
                    elif "Charms Muranos" in nombre_linea: prefijo_base = "chm"
                    elif "Charms Cadenas de Seguridad" in nombre_linea: prefijo_base = "chcsd"
                    elif "Charms Beads Disney" in nombre_linea: prefijo_base = "chbd"
                    elif "Charms Colgantes Disney" in nombre_linea: 
                        prefijo_base = "chcd"
                    elif "Charms Colgantes" in nombre_linea:
                        prefijo_base = "chc"
                        separador_char = "."
                    elif "Charms Beads" in nombre_linea: prefijo_base = "chb"
                    else: prefijo_base = nombre_linea.split()[0].lower()

                extensiones = ['.jpg', '.jpeg', '.avif', '.webp', '.png']
                patrones_sufijos = ['', '.1', '.0', '_1', '_0'] # Estricto: sin .2
                
                posibles_rutas = []
                for suf in patrones_sufijos:
                    for ext in extensiones:
                        posibles_rutas.append(f"{carpeta_imagenes}/{prefijo_base}{separador_char}{prod_id}{suf}{ext}")

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

                # ENLACE CORREGIDO Y ALINEADO CON EL INDEX.HTML:
                # Transforma el nombre de la línea exactamente al formato plano que el index.html procesa para abrir la galería
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
    tareas_pandora = [
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

    tareas_swarovski = [
        ("ANILLOS SWA.xlsx", "Anillos Swarovski", "imagenes/SWA/anillos_swa"),
        ("ARETES SWA.xlsx", "Aretes Swarovski", "imagenes/SWA/aretes_swa"),
        ("PULSERAS SWA.xlsx", "Pulseras Swarovski", "imagenes/SWA/pulseras_swa"),
        ("COLLARES SWA.xlsx", "Collares Swarovski", "imagenes/SWA/collares_swa")
    ]

    tareas_bano_plata = [
        ("ANILLOS BP.xlsx", "Anillos Baño de Plata", "imagenes/BP/anillosbp"),
        ("ARETES BP.xlsx", "Aretes Baño de Plata", "imagenes/BP/aretessbp"),
        ("CADENAS BP.xlsx", "Cadenas Baño de Plata", "imagenes/BP/cadenasbp"),
        ("PULSERAS BP.xlsx", "Pulseras Baño de Plata", "imagenes/BP/pulserasbp")
    ]

    while True:
        print("=" * 55)
        print("     ASISTENTE DE GENERACIÓN DE CATÁLOGOS PDF")
        print("=" * 55)
        print("Menú Principal - Selecciona una Opción:")
        print("1. Catálogo GENERAL Completo (Índice de Colecciones + Subíndices)")
        print("2. Catálogo completo solo de Pandora")
        print("3. Catálogo completo solo de Swarovski")
        print("4. Catálogo completo solo de Baño de Plata")
        print("5. Salir")
        
        coleccion_op = input("\nElige una opción (1-5): ").strip()
        
        if coleccion_op == "1":
            diccionario_a_pasar = {
                "Pandora": tareas_pandora,
                "Swarovski": tareas_swarovski,
                "Baño de Plata": tareas_bano_plata
            }
            nombre_col_str = "Maestro_General"
        elif coleccion_op == "2":
            diccionario_a_pasar = {"Pandora": tareas_pandora}
            nombre_col_str = "Pandora"
        elif coleccion_op == "3":
            diccionario_a_pasar = {"Swarovski": tareas_swarovski}
            nombre_col_str = "Swarovski"
        elif coleccion_op == "4":
            diccionario_a_pasar = {"Baño de Plata": tareas_bano_plata}
            nombre_col_str = "Bano_de_Plata"
        elif coleccion_op == "5" or coleccion_op.lower() == "salir":
            print("Saliendo del asistente...")
            break
        else:
            print("\nOpción no válida. Inténtalo de nuevo.\n")
            continue

        resp_precios = input("¿Deseas incluir precios en el catálogo? (s/n): ").strip().lower()
        mostrar_precios = resp_precios in ["s", "si", "sí", "y", "yes"]

        margen_porcentaje = 0.0
        if mostrar_precios:
            try:
                resp_margen = input("Ingresa el porcentaje de margen extra (ej. 0 para base, 15 para +15%): ").strip()
                margen_porcentaje = float(resp_margen) if resp_margen else 0.0
            except ValueError:
                margen_porcentaje = 0.0

        print(f"\nGenerando estructura de catálogo con índice y subíndices...")
        generar_pdf_estructurado(
            diccionario_colecciones=diccionario_a_pasar, 
            mostrar_precios=mostrar_precios, 
            margen_porcentaje=margen_porcentaje, 
            coleccion_seleccionada=nombre_col_str
        )

        otro = input("¿Deseas realizar otra operación? (s/n): ").strip().lower()
        if otro not in ["s", "si", "sí", "y", "yes"]:
            print("¡Hasta luego!")
            break