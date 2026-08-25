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
            self.set_y(-18)
            self.set_font('helvetica', '', 6.5)
            self.set_draw_color(210, 210, 210)
            self.line(12, self.get_y(), 198, self.get_y())
            self.ln(2)
            
            # Menú inferior ordenado en dos filas para evitar amontonamiento
            mitad = (len(self.lista_lineas_links) + 1) // 2
            fila_1 = self.lista_lineas_links[:mitad]
            fila_2 = self.lista_lineas_links[mitad:]

            self.set_text_color(100, 100, 100)
            self.cell(6, 3.5, "Ir:", align='L')
            
            if self.link_indice_coleccion:
                self.set_text_color(31, 78, 121)
                self.set_font('helvetica', 'B', 6.5)
                self.cell(16, 3.5, "[Subíndice]", link=self.link_indice_coleccion, align='L')

            self.set_font('helvetica', '', 6.5)
            for nombre_corto, l_obj in fila_1:
                self.set_text_color(0, 102, 204)
                self.cell(len(nombre_corto) * 1.8 + 2, 3.5, f"[{nombre_corto}]", link=l_obj, align='L')

            self.set_text_color(140, 140, 140)
            self.set_xy(182, self.get_y())
            self.cell(16, 3.5, f"Pág. {self.page_no()}", align='R')

            if fila_2:
                self.ln(3.5)
                self.set_x(12)
                self.cell(22, 3.5, "", align='L')
                for nombre_corto, l_obj in fila_2:
                    self.set_text_color(0, 102, 204)
                    self.cell(len(nombre_corto) * 1.8 + 2, 3.5, f"[{nombre_corto}]", link=l_obj, align='L')

def generar_pdf_estructurado(tareas_coleccion, coleccion_nombre, mostrar_precios=True, margen_porcentaje=0):
    subcarpeta_tipo = "Catálogo PDF precios" if mostrar_precios else "Catálogo PDF"
    carpeta_destino = os.path.join("catalogos", coleccion_nombre, subcarpeta_tipo)
    
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)

    sufijo_precio = "p" if mostrar_precios else ""
    nombre_archivo = os.path.join(carpeta_destino, f"catalogo_{coleccion_nombre.lower()}{sufijo_precio}.pdf")

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

    pdf = PDFCatalogoJerarquico([], [], coleccion_nombre)
    pdf.set_auto_page_break(auto=True, margin=18)

    # PÁGINA 1: SUBÍNDICE DE CATEGORÍAS DE LA COLECCIÓN
    pdf.add_page()
    link_sub_obj = pdf.add_link()
    pdf.link_indice_coleccion = link_sub_obj
    pdf.set_link(link_sub_obj, y=0)

    pdf.set_font('helvetica', 'B', 20)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 10, f"Colección: {coleccion_nombre.replace('_', ' ')}", new_x="LMARGIN", new_y="NEXT", align='C')
    
    pdf.set_font('helvetica', 'B', 13)
    pdf.set_text_color(90, 100, 110)
    pdf.cell(0, 7, "Subíndice de Categorías", new_x="LMARGIN", new_y="NEXT", align='C')
    pdf.ln(4)

    pdf.set_draw_color(31, 78, 121)
    pdf.set_line_width(0.6)
    pdf.line(55, pdf.get_y(), 155, pdf.get_y())
    pdf.ln(6)

    links_lineas_col = {}
    for _, nombre_linea, _ in tareas_coleccion:
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

    for _, nombre_linea, _ in tareas_coleccion:
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

    # Renderizar productos de cada línea de la colección unificada
    for ruta_excel, nombre_linea, carpeta_imagenes in tareas_coleccion:
        if not os.path.exists(ruta_excel):
            continue
        try:
            if "swarovski" in coleccion_nombre.lower() or "swa" in carpeta_imagenes.lower():
                df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
            elif "baño" in coleccion_nombre.lower() or "bano" in coleccion_nombre.lower() or "bp" in carpeta_imagenes.lower():
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
                elif "Charms Colgantes Disney" in nombre_linea: prefijo_base = "chcd"
                elif "Charms Colgantes" in nombre_linea:
                    prefijo_base = "chc"
                    separador_char = "."
                elif "Charms Beads" in nombre_linea: prefijo_base = "chb"
                else: prefijo_base = nombre_linea.split()[0].lower()

            extensiones = ['.jpg', '.jpeg', '.avif', '.webp', '.png']
            patrones_sufijos = ['', '.1', '.0', '_1', '_0']
            
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
    print(f"\n-> ¡Éxito! Catálogo de {coleccion_nombre} guardado en: {nombre_archivo}\n")

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
        print("Selecciona la Colección Completa:")
        print("1. Pandora (Todas las líneas unificadas)")
        print("2. Swarovski (Todas las líneas unificadas)")
        print("3. Baño de Plata (Todas las líneas unificadas)")
        print("4. Salir")
        
        coleccion_op = input("\nElige una opción (1-4): ").strip()
        
        if coleccion_op == "1":
            tareas_seleccionadas = tareas_pandora
            nombre_col = "Pandora"
        elif coleccion_op == "2":
            tareas_seleccionadas = tareas_swarovski
            nombre_col = "Swarovski"
        elif coleccion_op == "3":
            tareas_seleccionadas = tareas_bano_plata
            nombre_col = "Bano_de_Plata"
        elif coleccion_op == "4" or coleccion_op.lower() == "salir":
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

        print(f"\nGenerando catálogo completo de {nombre_col} con su subíndice y menú de navegación...")
        generar_pdf_estructurado(
            tareas_coleccion=tareas_seleccionadas,
            coleccion_nombre=nombre_col,
            mostrar_precios=mostrar_precios,
            margen_porcentaje=margen_porcentaje
        )

        otro = input("¿Deseas realizar otra operación? (s/n): ").strip().lower()
        if otro not in ["s", "si", "sí", "y", "yes"]:
            print("¡Hasta luego!")
            break