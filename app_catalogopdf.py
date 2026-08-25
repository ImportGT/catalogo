import os
import io
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
from fpdf import FPDF
from PIL import Image

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

class AppGeneradorCatalogoPDF:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Catálogos PDF - Joyería")
        self.root.geometry("520x420")
        self.root.config(bg="#f4f6f9")

        self.tareas_pandora = [
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

        self.tareas_swarovski = [
            ("ANILLOS SWA.xlsx", "Anillos Swarovski", "imagenes/SWA/anillos_swa"),
            ("ARETES SWA.xlsx", "Aretes Swarovski", "imagenes/SWA/aretes_swa"),
            ("PULSERAS SWA.xlsx", "Pulseras Swarovski", "imagenes/SWA/pulseras_swa"),
            ("COLLARES SWA.xlsx", "Collares Swarovski", "imagenes/SWA/collares_swa")
        ]

        self.tareas_bano_plata = [
            ("ANILLOS BP.xlsx", "Anillos Baño de Plata", "imagenes/BP/anillosbp"),
            ("ARETES BP.xlsx", "Aretes Baño de Plata", "imagenes/BP/aretessbp"),
            ("CADENAS BP.xlsx", "Cadenas Baño de Plata", "imagenes/BP/cadenasbp"),
            ("PULSERAS BP.xlsx", "Pulseras Baño de Plata", "imagenes/BP/pulserasbp")
        ]

        # --- INTERFAZ GRÁFICA ---
        frame_main = tk.LabelFrame(root, text=" Configuración de Catálogo PDF ", font=("Helvetica", 10, "bold"), bg="#f4f6f9", fg="#1F4E79", padx=15, pady=15)
        frame_main.pack(fill="both", expand=True, padx=20, pady=20)

        tk.Label(frame_main, text="Selecciona la Colección:", bg="#f4f6f9", font=("Helvetica", 9, "bold")).pack(anchor="w", pady=(5, 2))
        self.coleccion_var = tk.StringVar(value="Pandora")
        
        tk.Radiobutton(frame_main, text="Pandora (Completo)", variable=self.coleccion_var, value="Pandora", bg="#f4f6f9", font=("Helvetica", 9)).pack(anchor="w")
        tk.Radiobutton(frame_main, text="Swarovski", variable=self.coleccion_var, value="Swarovski", bg="#f4f6f9", font=("Helvetica", 9)).pack(anchor="w")
        tk.Radiobutton(frame_main, text="Baño de Plata", variable=self.coleccion_var, value="Bano_de_Plata", bg="#f4f6f9", font=("Helvetica", 9)).pack(anchor="w")

        self.precios_var = tk.BooleanVar(value=True)
        chk_precios = tk.Checkbutton(frame_main, text="Incluir Precios en el Catálogo", variable=self.precios_var, bg="#f4f6f9", font=("Helvetica", 9, "bold"), command=self.toggle_margen)
        chk_precios.pack(anchor="w", pady=(15, 5))

        frame_margen = tk.Frame(frame_main, bg="#f4f6f9")
        frame_margen.pack(anchor="w", pady=5)
        
        tk.Label(frame_margen, text="Porcentaje de Margen Extra (%):", bg="#f4f6f9", font=("Helvetica", 9)).pack(side="left", padx=(0, 10))
        self.entry_margen = tk.Entry(frame_margen, width=10, font=("Helvetica", 9))
        self.entry_margen.insert(0, "0")
        self.entry_margen.pack(side="left")

        btn_generar = tk.Button(root, text="Generar Catálogo en PDF", bg="#008000", fg="white", font=("Helvetica", 10, "bold"), padx=15, pady=8, command=self.ejecutar_generacion)
        btn_generar.pack(pady=10)

    def toggle_margen(self):
        if self.precios_var.get():
            self.entry_margen.config(state="normal")
        else:
            self.entry_margen.config(state="disabled")

    def ejecutar_generacion(self):
        col_op = self.coleccion_var.get()
        mostrar_precios = self.precios_var.get()
        
        try:
            margen_porcentaje = float(self.entry_margen.get().strip() or 0) if mostrar_precios else 0.0
        except ValueError:
            margen_porcentaje = 0.0

        if col_op == "Pandora":
            tareas = self.tareas_pandora
            nombre_col = "Pandora"
        elif col_op == "Swarovski":
            tareas = self.tareas_swarovski
            nombre_col = "Swarovski"
        else:
            tareas = self.tareas_bano_plata
            nombre_col = "Bano_de_Plata"

        try:
            self.generar_pdf_estructurado(tareas, nombre_col, mostrar_precios, margen_porcentaje)
            messagebox.showinfo("¡Éxito!", f"El catálogo optimizado se ha generado correctamente en 'CATALOGOS PDF'.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al generar el PDF: {e}")

    def comprimir_imagen_para_pdf(self, ruta_imagen, max_dim=800, calidad=75):
        try:
            img = Image.open(ruta_imagen)
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            img_io = io.BytesIO()
            img.save(img_io, format="JPEG", quality=calidad, optimize=True)
            img_io.seek(0)
            return img_io
        except Exception:
            return ruta_imagen

    def generar_pdf_estructurado(self, tareas_coleccion, coleccion_nombre, mostrar_precios=True, margen_porcentaje=0):
        carpeta_destino = "CATALOGOS PDF"
        if not os.path.exists(carpeta_destino):
            os.makedirs(carpeta_destino)

        if coleccion_nombre == "Pandora":
            nombre_archivo_final = "PRECIOS PANDORA.pdf" if mostrar_precios else "CATALOGO PANDORA.pdf"
        elif coleccion_nombre == "Swarovski":
            nombre_archivo_final = "PRECIOS SWAROVSKI.pdf" if mostrar_precios else "CATALOGO SWAROVSKI.pdf"
        else:
            nombre_archivo_final = "BAÑO DE PLATA PRECIOS.pdf" if mostrar_precios else "CATALOGO BAÑO DE PLATA.pdf"

        nombre_archivo = os.path.join(carpeta_destino, nombre_archivo_final)

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

        galerias_productos = []
        extensiones_validas = ['.jpg', '.jpeg', '.png', '.webp', '.avif', '.bmp', '.gif', '.tiff']

        for ruta_excel, nombre_linea, carpeta_imagenes in tareas_coleccion:
            if not os.path.exists(ruta_excel):
                continue
            try:
                if "swarovski" in coleccion_nombre.lower() or "swa" in carpeta_imagenes.lower():
                    df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
                elif "baño" in coleccion_nombre.lower() or "bano" in coleccion_nombre.lower() or "bp" in carpeta_imagenes.lower():
                    df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
                else:
                    df = pd.read_excel(ruta_excel, header=0)
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
            cardHeight = 102 if mostrar_precios else 92
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

                imagen_path = ""
                patrones_portada = ['', '.1', '.0', '_1', '_0']
                for suf in patrones_portada:
                    for ext in extensiones_validas:
                        prueba = f"{carpeta_imagenes}/{prefijo_base}{separador_char}{prod_id}{suf}{ext}"
                        if os.path.exists(prueba):
                            imagen_path = prueba
                            break
                    if imagen_path: break

                imagenes_extras = []
                for ext in extensiones_validas:
                    for i in range(2, 10):
                        prueba_extra = f"{carpeta_imagenes}/{prefijo_base}{separador_char}{prod_id}_{i}{ext}"
                        if os.path.exists(prueba_extra):
                            imagenes_extras.append(prueba_extra)
                        prueba_extra_pto = f"{carpeta_imagenes}/{prefijo_base}{separador_char}{prod_id}.{i}{ext}"
                        if os.path.exists(prueba_extra_pto):
                            imagenes_extras.append(prueba_extra_pto)

                link_galeria = pdf.add_link() if imagenes_extras else None
                link_retorno = pdf.add_link() if imagenes_extras else None

                if imagenes_extras:
                    galerias_productos.append((prod_id, nombre_linea, link_galeria, link_retorno, imagenes_extras, imagen_path, precio_final))

                if currentY + cardHeight > 265:
                    pdf.add_page()
                    currentY = 20
                    currentX = startX
                    countInRow = 0

                if link_retorno:
                    pdf.set_link(link_retorno, y=currentY)

                pdf.set_draw_color(220, 220, 220)
                pdf.set_fill_color(255, 255, 255)
                pdf.rect(currentX, currentY, cardWidth, cardHeight, 'DF')

                img_w_std = 60
                img_h_std = 60
                img_x = currentX + (cardWidth - img_w_std) / 2
                img_y = currentY + 5

                if imagen_path and os.path.exists(str(imagen_path)):
                    try:
                        img_comprimida = self.comprimir_imagen_para_pdf(str(imagen_path))
                        pdf.image(img_comprimida, x=img_x, y=img_y, w=img_w_std, h=img_h_std)
                    except Exception:
                        pass

                pdf.set_font('helvetica', 'B', 11)
                pdf.set_text_color(44, 62, 80)
                pdf.set_xy(currentX, currentY + 68)
                pdf.cell(cardWidth, 6, f"Cod. {prod_id}", new_x="RIGHT", new_y="TOP", align='C')

                y_offset_precio = 75
                if mostrar_precios:
                    pdf.set_font('helvetica', 'B', 12)
                    pdf.set_text_color(0, 128, 0)
                    pdf.set_xy(currentX, currentY + y_offset_precio)
                    pdf.cell(cardWidth, 6, f"Q{precio_final:.2f}", new_x="RIGHT", new_y="TOP", align='C')
                    y_offset_precio += 8

                if link_galeria:
                    pdf.set_font('helvetica', 'B', 9)
                    pdf.set_text_color(0, 102, 204)
                    pdf.set_xy(currentX, currentY + y_offset_precio)
                    pdf.cell(cardWidth, 6, "[Ver Galería]", link=link_galeria, align='C')

                countInRow += 1
                if countInRow >= cols:
                    countInRow = 0
                    currentX = startX
                    currentY += cardHeight + gapY
                else:
                    currentX += cardWidth + gapX

        # Generar páginas de galerías con ajuste dinámico absoluto para garantizar 1 sola página por producto
        for prod_id, linea_prod, link_obj, link_ret, lista_imgs, img_portada, precio_item in galerias_productos:
            pdf.add_page()
            pdf.set_link(link_obj, y=0)

            pdf.set_font('helvetica', 'B', 14)
            pdf.set_text_color(31, 78, 121)
            pdf.cell(0, 8, f"Galería: Cod. {prod_id} ({linea_prod})", new_x="LMARGIN", new_y="NEXT", align='L')
            
            if mostrar_precios:
                pdf.set_font('helvetica', 'B', 11)
                pdf.set_text_color(0, 128, 0)
                pdf.cell(0, 5, f"Precio: Q{precio_item:.2f}", new_x="LMARGIN", new_y="NEXT", align='L')
            
            pdf.set_font('helvetica', 'B', 9)
            pdf.set_text_color(192, 0, 0)
            pdf.cell(0, 5, "<< Regresar al Catálogo", link=link_ret, new_x="LMARGIN", new_y="NEXT", align='L')
            pdf.ln(3)

            todas_las_fotos = [img_portada] + lista_imgs if img_portada else lista_imgs
            total_fotos = len(todas_las_fotos)

            # Cálculo dinámico de filas, columnas y dimensiones para que NUNCA rebase la única página
            if total_fotos <= 2:
                cols_gal = 2
                filas_gal = 1
                cardW_gal = 82
                cardH_gal = 100
                img_size = 75
            elif total_fotos <= 4:
                cols_gal = 2
                filas_gal = 2
                cardW_gal = 82
                cardH_gal = 52
                img_size = 40
            elif total_fotos <= 6:
                cols_gal = 2
                filas_gal = 3
                cardW_gal = 82
                cardH_gal = 35
                img_size = 28
            else:
                cols_gal = 3
                filas_gal = (total_fotos + 2) // 3
                cardW_gal = 52
                cardH_gal = 32
                img_size = 25

            startX_gal = 20 if cols_gal == 2 else 15
            startY_gal = pdf.get_y()
            gap_x_gal = 8
            gap_y_gal = 5

            currX = startX_gal
            currY = startY_gal
            col_idx = 0
            row_idx = 0

            for foto_path in todas_las_fotos:
                if foto_path and os.path.exists(foto_path):
                    pdf.set_draw_color(200, 200, 200)
                    pdf.set_fill_color(255, 255, 255)
                    pdf.rect(currX, currY, cardW_gal, cardH_gal, 'DF')

                    try:
                        img_comprimida = self.comprimir_imagen_para_pdf(foto_path)
                        # Centrar imagen dentro de su tarjeta dinámica
                        off_x = (cardW_gal - img_size) / 2
                        off_y = (cardH_gal - img_size) / 2
                        pdf.image(img_comprimida, x=currX + off_x, y=currY + off_y, w=img_size, h=img_size)
                    except Exception:
                        pass

                    col_idx += 1
                    if col_idx >= cols_gal:
                        col_idx = 0
                        currX = startX_gal
                        currY += cardH_gal + gap_y_gal
                        row_idx += 1
                    else:
                        currX += cardW_gal + gap_x_gal

        pdf.output(nombre_archivo)

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGeneradorCatalogoPDF(root)
    root.mainloop()