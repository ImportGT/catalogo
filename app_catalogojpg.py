import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageDraw, ImageFont

class AppGeneradorCatalogoJPG:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Catálogos JPG - Joyería")
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
        frame_main = tk.LabelFrame(root, text=" Configuración de Catálogo JPG ", font=("Helvetica", 10, "bold"), bg="#f4f6f9", fg="#1F4E79", padx=15, pady=15)
        frame_main.pack(fill="both", expand=True, padx=20, pady=20)

        tk.Label(frame_main, text="Selecciona la Colección:", bg="#f4f6f9", font=("Helvetica", 9, "bold")).pack(anchor="w", pady=(5, 2))
        self.coleccion_var = tk.StringVar(value="Pandora")
        
        tk.Radiobutton(frame_main, text="Pandora (Completo)", variable=self.coleccion_var, value="Pandora", bg="#f4f6f9", font=("Helvetica", 9)).pack(anchor="w")
        tk.Radiobutton(frame_main, text="Swarovski", variable=self.coleccion_var, value="Swarovski", bg="#f4f6f9", font=("Helvetica", 9)).pack(anchor="w")
        tk.Radiobutton(frame_main, text="Baño de Plata", variable=self.coleccion_var, value="Bano_de_Plata", bg="#f4f6f9", font=("Helvetica", 9)).pack(anchor="w")

        self.precios_var = tk.BooleanVar(value=True)
        chk_precios = tk.Checkbutton(frame_main, text="Incluir Precios en las Imágenes", variable=self.precios_var, bg="#f4f6f9", font=("Helvetica", 9, "bold"), command=self.toggle_margen)
        chk_precios.pack(anchor="w", pady=(15, 5))

        frame_margen = tk.Frame(frame_main, bg="#f4f6f9")
        frame_margen.pack(anchor="w", pady=5)
        
        tk.Label(frame_margen, text="Porcentaje de Margen Extra (%):", bg="#f4f6f9", font=("Helvetica", 9)).pack(side="left", padx=(0, 10))
        self.entry_margen = tk.Entry(frame_margen, width=10, font=("Helvetica", 9))
        self.entry_margen.insert(0, "0")
        self.entry_margen.pack(side="left")

        btn_generar = tk.Button(root, text="Generar Catálogos en JPG", bg="#008000", fg="white", font=("Helvetica", 10, "bold"), padx=15, pady=8, command=self.ejecutar_generacion)
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
            self.generar_jpgs_estructurados(tareas, nombre_col, mostrar_precios, margen_porcentaje)
            messagebox.showinfo("¡Éxito!", f"Las imágenes JPG se han generado correctamente en la carpeta 'CATALOGOS JPG'.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al generar las imágenes: {e}")

    def generar_jpgs_estructurados(self, tareas_coleccion, coleccion_nombre, mostrar_precios=True, margen_porcentaje=0):
        carpeta_destino = os.path.join("CATALOGOS JPG", coleccion_nombre)
        if not os.path.exists(carpeta_destino):
            os.makedirs(carpeta_destino)

        extensiones_validas = ['.jpg', '.jpeg', '.png', '.webp', '.avif', '.bmp', '.gif', '.tiff']

        # Dimensiones del lienzo JPG (Formato vertical tipo plantilla web/móvil)
        canvas_width = 1200
        cols = 3
        margin_x = 50
        spacing_x = 40
        card_w = (canvas_width - (margin_x * 2) - (spacing_x * (cols - 1))) // cols
        card_h = 420 if mostrar_precios else 380
        img_box_size = 280

        try:
            font_titulo = ImageFont.truetype("arial.ttf", 36)
            font_codigo = ImageFont.truetype("arialbd.ttf", 26)
            font_precio = ImageFont.truetype("arialbd.ttf", 28)
        except IOError:
            font_titulo = font_codigo = font_precio = ImageFont.load_default()

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

            productos_linea = []
            for _, row in df.iterrows():
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

                if imagen_path and os.path.exists(imagen_path):
                    productos_linea.append((prod_id, imagen_path, precio_final))

            if not productos_linea:
                continue

            # Agrupar en páginas de 15 productos por imagen JPG para mantener buena visibilidad
            productos_por_pagina = 15
            for pagina_idx in range(0, len(productos_linea), productos_por_pagina):
                lote = productos_linea[pagina_idx:pagina_idx + productos_por_pagina]
                
                filas = (len(lote) + cols - 1) // cols
                margin_y = 120
                spacing_y = 40
                canvas_height = margin_y + (filas * (card_h + spacing_y)) + 60

                img_canvas = Image.new("RGB", (canvas_width, canvas_height), (245, 247, 250))
                draw = ImageDraw.Draw(img_canvas)

                # Encabezado limpio
                titulo_texto = f"{nombre_linea} (Parte {pagina_idx // productos_por_pagina + 1})" if len(productos_linea) > productos_por_pagina else nombre_linea
                draw.text((margin_x, 40), titulo_texto, fill=(31, 78, 121), font=font_titulo)

                curr_x = margin_x
                curr_y = margin_y
                col_idx = 0

                for prod_id, foto_path, precio_val in lote:
                    # Dibujar tarjeta blanca
                    draw.rectangle([curr_x, curr_y, curr_x + card_w, curr_y + card_h], fill=(255, 255, 255), outline=(220, 225, 230), width=2)

                    try:
                        prod_img = Image.open(foto_path)
                        if prod_img.mode in ("RGBA", "P"):
                            prod_img = prod_img.convert("RGB")
                        prod_img.thumbnail((img_box_size, img_box_size), Image.Resampling.LANCZOS)
                        
                        # Centrar imagen en su espacio asignado
                        p_w, p_h = prod_img.size
                        p_x = curr_x + (card_w - p_w) // 2
                        p_y = curr_y + 20 + (img_box_size - p_h) // 2
                        img_canvas.paste(prod_img, (p_x, p_y))
                    except Exception:
                        pass

                    # Texto de Código
                    txt_cod = f"Cod. {prod_id}"
                    draw.text((curr_x + card_w // 2, curr_y + 315), txt_cod, fill=(44, 62, 80), font=font_codigo, anchor="mm")

                    # Texto de Precio opcional
                    if mostrar_precios:
                        txt_precio = f"Q{precio_val:.2f}"
                        draw.text((curr_x + card_w // 2, curr_y + 365), txt_precio, fill=(0, 128, 0), font=font_precio, anchor="mm")

                    col_idx += 1
                    if col_idx >= cols:
                        col_idx = 0
                        curr_x = margin_x
                        curr_y += card_h + spacing_y
                    else:
                        curr_x += card_w + spacing_x

                sufijo_p = "_precios" if mostrar_precios else "_catalogo"
                nombre_limpio = nombre_linea.replace(' ', '_').lower()
                num_parte = f"_p{pagina_idx // productos_por_pagina + 1}" if len(productos_linea) > productos_por_pagina else ""
                nombre_archivo_jpg = f"{nombre_limpio}{sufijo_p}{num_parte}.jpg"

                ruta_guardado = os.path.join(carpeta_destino, nombre_archivo_jpg)
                img_canvas.save(ruta_guardado, "JPEG", quality=85, optimize=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGeneradorCatalogoJPG(root)
    root.mainloop()