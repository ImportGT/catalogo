import os
import random
import json
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageDraw, ImageFont

HISTORIAL_PRODUCTO_DIA = "historial_producto_dia.json"

class AppGeneradorPublicidad:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Publicidad Estética - Joyería")
        self.root.geometry("580x580")
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
        frame_main = tk.LabelFrame(root, text=" Creador de Publicidad Comercial Boutique ", font=("Helvetica", 10, "bold"), bg="#f4f6f9", fg="#1F4E79", padx=15, pady=15)
        frame_main.pack(fill="both", expand=True, padx=20, pady=20)

        # 1. Colección
        tk.Label(frame_main, text="1. Selecciona la Colección:", bg="#f4f6f9", font=("Helvetica", 9, "bold")).pack(anchor="w", pady=(2, 2))
        self.coleccion_var = tk.StringVar(value="Pandora")
        tk.Radiobutton(frame_main, text="Pandora", variable=self.coleccion_var, value="Pandora", bg="#f4f6f9", font=("Helvetica", 9), command=self.actualizar_lineas).pack(anchor="w")
        tk.Radiobutton(frame_main, text="Swarovski", variable=self.coleccion_var, value="Swarovski", bg="#f4f6f9", font=("Helvetica", 9), command=self.actualizar_lineas).pack(anchor="w")
        tk.Radiobutton(frame_main, text="Baño de Plata", variable=self.coleccion_var, value="Bano_de_Plata", bg="#f4f6f9", font=("Helvetica", 9), command=self.actualizar_lineas).pack(anchor="w")

        # 2. Precios y Margen
        self.precios_var = tk.BooleanVar(value=True)
        chk_precios = tk.Checkbutton(frame_main, text="2. Incluir Precios en la Publicidad", variable=self.precios_var, bg="#f4f6f9", font=("Helvetica", 9, "bold"), command=self.toggle_margen)
        chk_precios.pack(anchor="w", pady=(8, 2))

        frame_margen = tk.Frame(frame_main, bg="#f4f6f9")
        frame_margen.pack(anchor="w", pady=2)
        tk.Label(frame_margen, text="Margen Extra (%):", bg="#f4f6f9", font=("Helvetica", 9)).pack(side="left", padx=(0, 10))
        self.entry_margen = tk.Entry(frame_margen, width=8, font=("Helvetica", 9))
        self.entry_margen.insert(0, "0")
        self.entry_margen.pack(side="left")

        # 3. Tipo de Publicidad
        tk.Label(frame_main, text="3. Tipo de Publicidad:", bg="#f4f6f9", font=("Helvetica", 9, "bold")).pack(anchor="w", pady=(8, 2))
        self.tipo_pub_var = tk.StringVar(value="ultimos")
        tk.Radiobutton(frame_main, text="Últimos Ingresos (Top 5 Globales)", variable=self.tipo_pub_var, value="ultimos", bg="#f4f6f9", font=("Helvetica", 9), command=self.toggle_producto_especifico).pack(anchor="w")
        tk.Radiobutton(frame_main, text="Producto del Día (Aleatorio con galería)", variable=self.tipo_pub_var, value="dia", bg="#f4f6f9", font=("Helvetica", 9), command=self.toggle_producto_especifico).pack(anchor="w")
        tk.Radiobutton(frame_main, text="Elegir Producto Específico (por Línea y Código)", variable=self.tipo_pub_var, value="especifico", bg="#f4f6f9", font=("Helvetica", 9), command=self.toggle_producto_especifico).pack(anchor="w")

        # Controles para Producto Específico (Línea y Código)
        self.frame_esp = tk.Frame(frame_main, bg="#f4f6f9")
        
        tk.Label(self.frame_esp, text="Selecciona la Línea:", bg="#f4f6f9", font=("Helvetica", 9)).pack(anchor="w", pady=(4, 2))
        self.combo_lineas = ttk.Combobox(self.frame_esp, state="readonly", width=38, font=("Helvetica", 9))
        self.combo_lineas.pack(anchor="w", pady=(0, 6))

        frame_cod = tk.Frame(self.frame_esp, bg="#f4f6f9")
        frame_cod.pack(anchor="w", fill="x", pady=2)
        tk.Label(frame_cod, text="Código del Producto:", bg="#f4f6f9", font=("Helvetica", 9)).pack(side="left", padx=(0, 10))
        self.entry_codigo_esp = tk.Entry(frame_cod, width=12, font=("Helvetica", 9))
        self.entry_codigo_esp.pack(side="left")

        self.frame_esp.pack_forget() # Oculto por defecto
        self.actualizar_lineas()

        btn_generar = tk.Button(root, text="Generar Publicidad Estética", bg="#1F4E79", fg="white", font=("Helvetica", 10, "bold"), padx=15, pady=8, command=self.ejecutar_generacion)
        btn_generar.pack(pady=10)

    def toggle_margen(self):
        if self.precios_var.get():
            self.entry_margen.config(state="normal")
        else:
            self.entry_margen.config(state="disabled")

    def toggle_producto_especifico(self):
        if self.tipo_pub_var.get() == "especifico":
            self.frame_esp.pack(anchor="w", pady=5, fill="x")
        else:
            self.frame_esp.pack_forget()

    def actualizar_lineas(self):
        col_op = self.coleccion_var.get()
        if col_op == "Pandora":
            tareas = self.tareas_pandora
        elif col_op == "Swarovski":
            tareas = self.tareas_swarovski
        else:
            tareas = self.tareas_bano_plata

        nombres_lineas = [nombre for _, nombre, _ in tareas]
        self.combo_lineas['values'] = nombres_lineas
        if nombres_lineas:
            self.combo_lineas.current(0)

    def obtener_tareas_y_nombre(self):
        col_op = self.coleccion_var.get()
        if col_op == "Pandora":
            return self.tareas_pandora, "Pandora"
        elif col_op == "Swarovski":
            return self.tareas_swarovski, "Swarovski"
        else:
            return self.tareas_bano_plata, "Bano_de_Plata"

    def cargar_datos_coleccion(self, tareas):
        extensiones_validas = ['.jpg', '.jpeg', '.png', '.webp', '.avif', '.bmp', '.gif', '.tiff']
        todos_productos = []

        for ruta_excel, nombre_linea, carpeta_imagenes in tareas:
            if not os.path.exists(ruta_excel):
                continue
            try:
                if "swarovski" in carpeta_imagenes.lower() or "swa" in carpeta_imagenes.lower():
                    df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
                elif "baño" in carpeta_imagenes.lower() or "bano" in carpeta_imagenes.lower() or "bp" in carpeta_imagenes.lower():
                    df = pd.read_excel(ruta_excel, header=3 if "Aretes" in nombre_linea else 2)
                else:
                    df = pd.read_excel(ruta_excel, header=0)
                df.columns = [str(c).strip().upper() for c in df.columns]
            except Exception:
                continue

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
                for suf in ['', '.1', '.0', '_1', '_0']:
                    for ext in extensiones_validas:
                        prueba = f"{carpeta_imagenes}/{prefijo_base}{separador_char}{prod_id}{suf}{ext}"
                        if os.path.exists(prueba):
                            imagen_path = prueba
                            break
                    if imagen_path: break

                if not imagen_path or not os.path.exists(imagen_path):
                    continue

                imagenes_extras = []
                for ext in extensiones_validas:
                    for i in range(2, 15):
                        p1 = f"{carpeta_imagenes}/{prefijo_base}{separador_char}{prod_id}_{i}{ext}"
                        p2 = f"{carpeta_imagenes}/{prefijo_base}{separador_char}{prod_id}.{i}{ext}"
                        if os.path.exists(p1) and p1 not in imagenes_extras: imagenes_extras.append(p1)
                        if os.path.exists(p2) and p2 not in imagenes_extras: imagenes_extras.append(p2)

                galeria_completa = [imagen_path] + [img for img in imagenes_extras if img != imagen_path]

                try:
                    tiempo_mod = os.path.getmtime(imagen_path)
                except Exception:
                    tiempo_mod = 0.0

                todos_productos.append({
                    "id": prod_id,
                    "linea": nombre_linea,
                    "precio_base": precio_base,
                    "portada": imagen_path,
                    "galeria": galeria_completa,
                    "timestamp": tiempo_mod
                })

        todos_productos.sort(key=lambda x: x["timestamp"])
        return todos_productos

    def ejecutar_generacion(self):
        tareas, nombre_col = self.obtener_tareas_y_nombre()
        mostrar_precios = self.precios_var.get()
        
        try:
            margen = float(self.entry_margen.get().strip() or 0) if mostrar_precios else 0.0
        except ValueError:
            margen = 0.0

        tipo = self.tipo_pub_var.get()
        productos = self.cargar_datos_coleccion(tareas)

        if not productos:
            messagebox.showerror("Error", "No se encontraron productos con imágenes en esta colección.")
            return

        carpeta_destino = "PUBLICIDAD CREADA"
        if not os.path.exists(carpeta_destino):
            os.makedirs(carpeta_destino)

        try:
            if tipo == "ultimos":
                ultimos_5 = productos[-5:]
                self.generar_imagen_ultimos_ingresos(ultimos_5, nombre_col, mostrar_precios, margen, carpeta_destino)
                messagebox.showinfo("¡Éxito!", "Publicidad estética de Últimos Ingresos generada correctamente.")

            elif tipo == "dia":
                con_galeria = [p for p in productos if len(p["galeria"]) > 1]
                if not con_galeria:
                    messagebox.showwarning("Aviso", "No hay productos con galería adicional en esta colección. Se usará cualquiera disponible.")
                    con_galeria = productos

                prod_elegido = self.obtener_producto_dia_sin_repetir(con_galeria, nombre_col)
                self.generar_imagen_producto_galeria(prod_elegido, "PRODUCTO DEL DÍA", nombre_col, mostrar_precios, margen, carpeta_destino)
                messagebox.showinfo("¡Éxito!", f"¡Producto del Día (Cod. {prod_elegido['id']}) generado correctamente!")

            elif tipo == "especifico":
                linea_elegida = self.combo_lineas.get()
                cod_str = self.entry_codigo_esp.get().strip()
                if not cod_str.isdigit():
                    messagebox.showerror("Error", "Ingresa un código de producto válido.")
                    return
                cod_int = int(cod_str)
                
                encontrado = next((p for p in productos if p["id"] == cod_int and p["linea"] == linea_elegida), None)
                if not encontrado:
                    messagebox.showerror("Error", f"No se encontró el producto con código {cod_int} en la línea '{linea_elegida}'.")
                    return

                slogan = "¡Elegancia y estilo que enamoran en cada detalle!"
                self.generar_imagen_producto_galeria(encontrado, slogan, nombre_col, mostrar_precios, margen, carpeta_destino, es_especifico=True)
                messagebox.showinfo("¡Éxito!", f"Publicidad del producto {cod_int} generada correctamente.")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al generar la publicidad: {e}")

    def obtener_producto_dia_sin_repetir(self, lista_productos, nombre_col):
        historial = {}
        if os.path.exists(HISTORIAL_PRODUCTO_DIA):
            try:
                with open(HISTORIAL_PRODUCTO_DIA, "r", encoding="utf-8") as f:
                    historial = json.load(f)
            except Exception:
                pass

        usados = set(historial.get(nombre_col, []))
        disponibles = [p for p in lista_productos if p["id"] not in usados]

        if not disponibles:
            usados = set()
            disponibles = lista_productos

        elegido = random.choice(disponibles)
        usados.add(elegido["id"])
        historial[nombre_col] = list(usados)

        try:
            with open(HISTORIAL_PRODUCTO_DIA, "w", encoding="utf-8") as f:
                json.dump(historial, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

        return elegido

    def crear_tarjeta_estetica(self, w_card, h_card, bg_color=(255, 255, 255), radius=16):
        base = Image.new("RGBA", (w_card, h_card), (0, 0, 0, 0))
        draw = ImageDraw.Draw(base)
        for i in range(6, 0, -1):
            alpha = int(3 * (7 - i))
            draw.rounded_rectangle([i, i + 2, w_card - i, h_card - i + 2], radius=radius, fill=(200, 205, 215, alpha))
        draw.rounded_rectangle([0, 0, w_card, h_card], radius=radius, fill=bg_color, outline=(230, 235, 242), width=2)
        return base

    def generar_imagen_ultimos_ingresos(self, lista_prods, coleccion, mostrar_precios, margen, carpeta):
        w = 1200
        h = 1500
        img = Image.new("RGB", (w, h), (248, 250, 252))
        draw = ImageDraw.Draw(img)

        draw.rectangle([0, 0, w, 140], fill=(24, 43, 73))
        draw.rectangle([0, 137, w, 140], fill=(212, 175, 55))

        try:
            f_tit = ImageFont.truetype("arialbd.ttf", 40)
            f_sub = ImageFont.truetype("arial.ttf", 18)
            f_lin = ImageFont.truetype("arial.ttf", 16)
            f_cod = ImageFont.truetype("arialbd.ttf", 22)
            f_pre = ImageFont.truetype("arialbd.ttf", 24)
        except IOError:
            f_tit = f_sub = f_lin = f_cod = f_pre = ImageFont.load_default()

        draw.text((w // 2, 50), f"✨ ÚLTIMOS INGRESOS: {coleccion.upper()} ✨", fill=(255, 255, 255), font=f_tit, anchor="mm")
        draw.text((w // 2, 95), "Descubre nuestras novedades exclusivas", fill=(200, 215, 235), font=f_sub, anchor="mm")

        coords = [
            (80, 180, 560, 800),
            (640, 180, 1120, 800),
            (60, 840, 400, 1440),
            (430, 840, 770, 1440),
            (800, 840, 1140, 1440)
        ]

        for idx, prod in enumerate(lista_prods[:5]):
            x1, y1, x2, y2 = coords[idx]
            cw = x2 - x1
            ch = y2 - y1

            tarjeta = self.crear_tarjeta_estetica(cw, ch, radius=20)
            t_draw = ImageDraw.Draw(tarjeta)

            try:
                p_img = Image.open(prod["portada"])
                if p_img.mode in ("RGBA", "P"): p_img = p_img.convert("RGB")
                box_w = cw - 50
                box_h = ch - 130 if idx >= 2 else ch - 150
                p_img.thumbnail((box_w, box_h), Image.Resampling.LANCZOS)
                
                pw, ph = p_img.size
                px = (cw - pw) // 2
                py = 30
                tarjeta.paste(p_img, (px, py), p_img if p_img.mode == 'RGBA' else None)
            except Exception:
                pass

            t_draw.line([30, ch - 100, cw - 30, ch - 100], fill=(235, 240, 245), width=2)
            t_draw.text((cw // 2, ch - 75), prod["linea"], fill=(120, 130, 140), font=f_lin, anchor="mm")
            t_draw.text((cw // 2, ch - 48), f"Cod. {prod['id']}", fill=(31, 78, 121), font=f_cod, anchor="mm")
            
            if mostrar_precios:
                precio_f = prod["precio_base"] * (1 + (margen / 100))
                t_draw.text((cw // 2, ch - 22), f"Q{precio_f:.2f}", fill=(39, 134, 75), font=f_pre, anchor="mm")

            img.paste(tarjeta, (x1, y1), tarjeta)

        ruta = os.path.join(carpeta, f"Ultimos_Ingresos_{coleccion}.jpg")
        img.save(ruta, "JPEG", quality=95, optimize=True)

    def generar_imagen_producto_galeria(self, prod, titulo_o_slogan, coleccion, mostrar_precios, margen, carpeta, es_especifico=False):
        w = 1200
        h = 1500
        img = Image.new("RGB", (w, h), (248, 250, 252))
        draw = ImageDraw.Draw(img)

        # Encabezado boutique
        draw.rectangle([0, 0, w, 140], fill=(24, 43, 73))
        draw.rectangle([0, 137, w, 140], fill=(212, 175, 55))

        try:
            f_tit = ImageFont.truetype("arialbd.ttf", 36)
            f_sub = ImageFont.truetype("arial.ttf", 18)
        except IOError:
            f_tit = f_sub = ImageFont.load_default()

        draw.text((w // 2, 50), f"💎 {titulo_o_slogan} 💎" if not es_especifico else titulo_o_slogan, fill=(255, 255, 255), font=f_tit, anchor="mm")
        
        sub_txt = f"Línea: {prod['linea']}    •    Código: {prod['id']}"
        if mostrar_precios:
            precio_f = prod["precio_base"] * (1 + (margen / 100))
            sub_txt += f"    •    Precio: Q{precio_f:.2f}"
        draw.text((w // 2, 95), sub_txt, fill=(200, 215, 235), font=f_sub, anchor="mm")

        fotos = prod["galeria"]
        total = len(fotos)

        # Diseño asimétrico inspirador (Estilo Revista / Editorial con una foto protagonista grande a la izquierda)
        if total == 1:
            coords_gr = [(150, 180, 1050, 1380)]
        elif total == 2:
            coords_gr = [
                (80, 180, 680, 1380),   # Principal Izquierda Grande
                (710, 180, 1120, 1380)  # Secundaria Derecha Alta
            ]
        elif total == 3:
            coords_gr = [
                (80, 180, 680, 1380),   # Principal Izquierda Grande
                (710, 180, 1120, 765),  # Secundaria Der Sup
                (710, 795, 1120, 1380)  # Secundaria Der Inf
            ]
        elif total == 4:
            coords_gr = [
                (80, 180, 680, 1380),   # Principal Izquierda Grande
                (710, 180, 1120, 570),  # Der 1
                (710, 585, 1120, 975),  # Der 2
                (710, 990, 1120, 1380)  # Der 3
            ]
        elif total == 5:
            coords_gr = [
                (80, 180, 640, 1380),   # Principal Izq
                (670, 180, 1120, 465),  # Der 1
                (670, 480, 1120, 765),  # Der 2
                (670, 780, 1120, 1065), # Der 3
                (670, 1080, 1120, 1380) # Der 4
            ]
        else:
            # Si hay 6 o más fotos, 2 grandes principales izq y mosaico dinámico a la derecha
            coords_gr = []
            filas = (total - 1 + 1) // 2
            h_card = (1380 - 180 - ((filas - 1) * 20)) // filas
            
            # Izquierda
            coords_gr.append((80, 180, 640, 1380))
            
            # Derecha en mosaico
            for i in range(1, total):
                col = (i - 1) % 2
                row = (i - 1) // 2
                x1 = 670 + (col * 235)
                x2 = x1 + 215
                y1 = 180 + (row * (h_card + 15))
                y2 = min(1380, y1 + h_card)
                coords_gr.append((x1, y1, x2, y2))

        for idx, foto_path in enumerate(fotos):
            if idx >= len(coords_gr): break
            x1, y1, x2, y2 = coords_gr[idx]
            cw = x2 - x1
            ch = y2 - y1

            tarjeta = self.crear_tarjeta_estetica(cw, ch, radius=16)

            try:
                p_img = Image.open(foto_path)
                if p_img.mode in ("RGBA", "P"): p_img = p_img.convert("RGB")
                box_w = cw - 30
                box_h = ch - 30
                p_img.thumbnail((box_w, box_h), Image.Resampling.LANCZOS)
                
                pw, ph = p_img.size
                px = (cw - pw) // 2
                py = (ch - ph) // 2
                tarjeta.paste(p_img, (px, py), p_img if p_img.mode == 'RGBA' else None)
            except Exception:
                pass

            img.paste(tarjeta, (x1, y1), tarjeta)

        nombre_archivo = f"Producto_Dia_{prod['id']}.jpg" if not es_especifico else f"Publicidad_Cod_{prod['id']}.jpg"
        ruta = os.path.join(carpeta, nombre_archivo)
        img.save(ruta, "JPEG", quality=95, optimize=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGeneradorPublicidad(root)
    root.mainloop()