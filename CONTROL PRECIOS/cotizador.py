import json
import math
import os
import re
import sys
import io
import webbrowser
import urllib.request
import urllib.parse
import hashlib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Manejo de imágenes
from PIL import Image, ImageTk

# Generación de PDF
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

# --- DIRECTORIO ABSOLUTO DEL PROGRAMA ---
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_AUTOSAVE = os.path.join(BASE_DIR, "ultimo_prorrateo_activo.json")
CATALOGO_GLOBAL_FILE = os.path.join(BASE_DIR, "catalogo_general.json")
PROVEEDORES_FILE = os.path.join(BASE_DIR, "proveedores_custom.json")
CARPETA_CACHE_IMG = os.path.join(BASE_DIR, "cache_imagenes")

if not os.path.exists(CARPETA_CACHE_IMG):
    os.makedirs(CARPETA_CACHE_IMG)

# LISTA MAESTRA DE LÍNEAS DE PRODUCTO
LINEAS_LISTA = [
    "Pandora",
    "Swarovski",
    "Tous",
    "Baño de Plata",
    "Gorras",
]

# LISTA INICIAL DE TIENDAS / PROVEEDORES
PROVEEDORES_BASE = [
    "Friend",
    "Ako (Mayorista)",
]

# Categorías específicas por Línea
CATS_PANDORA = [
    "CH Beads",
    "CH Beads Disney",
    "CH Colgantes",
    "CH Colgantes Disney",
    "CH Muranos",
    "CH Cadenas de Seguridad",
    "CH Me",
    "CH Accesorios ME",
    "Anillos",
    "Collares",
    "Pulseras",
    "Aretes",
    "Otros",
]

CATEGORIAS_POR_LINEA = {
    "Pandora": CATS_PANDORA,
    "Swarovski": [
        "Anillos",
        "Collares / Dijes",
        "Pulseras / Brazaletes",
        "Aretes",
        "Conjuntos / Sets",
        "Figuras / Decoración",
        "Otros",
    ],
    "Tous": [
        "Ositos / Dijes",
        "Anillos",
        "Collares",
        "Pulseras",
        "Aretes",
        "Bolsos / Accesorios",
        "Otros",
    ],
    "Baño de Plata": [
        "Anillos",
        "Cadenas / Collares",
        "Pulseras",
        "Aretes",
        "Dijes",
        "Juegos / Sets",
        "Otros",
    ],
    "Gorras": [
        "Curvas",
        "Planas / Snapback",
        "Trucker / Malla",
        "Cerradas / Fitted",
        "Gorros de Lana",
        "Otros",
    ],
}


def obtener_direct_link_drive(url):
    if "drive.google.com" in url or "googleusercontent.com" in url:
        file_id = None
        if "/d/" in url:
            file_id = url.split("/d/")[1].split("/")[0]
        elif "id=" in url:
            parsed = urllib.parse.urlparse(url)
            file_id = urllib.parse.parse_qs(parsed.query).get('id', [None])[0]
        
        if file_id:
            return f"https://lh3.googleusercontent.com/d/{file_id}"
    return url


def descargar_y_guardar_localmente(url_o_path):
    if not url_o_path:
        return None

    if not (url_o_path.startswith("http://") or url_o_path.startswith("https://")):
        if os.path.exists(url_o_path):
            return url_o_path
        return None

    url_directa = obtener_direct_link_drive(url_o_path)
    hash_nombre = hashlib.md5(url_o_path.encode('utf-8')).hexdigest() + ".png"
    ruta_local_cache = os.path.join(CARPETA_CACHE_IMG, hash_nombre)

    if os.path.exists(ruta_local_cache) and os.path.getsize(ruta_local_cache) > 0:
        return ruta_local_cache

    try:
        req = urllib.request.Request(
            url_directa, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        with urllib.request.urlopen(req, timeout=12) as resp:
            data = resp.read()
            img = Image.open(io.BytesIO(data))
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            img.save(ruta_local_cache, "PNG")
            return ruta_local_cache
    except Exception as e:
        print(f"Error descargando imagen: {e}")

    return None


def descargar_imagen(url_o_path):
    ruta_local = descargar_y_guardar_localmente(url_o_path)
    if ruta_local and os.path.exists(ruta_local):
        try:
            return Image.open(ruta_local)
        except Exception as e:
            print(f"Error abriendo imagen PIL local: {e}")
    return None


class CotizadorApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Importaciones & Catálogo Multilínea (GTQ)")
        self.root.geometry("1450x980")
        self.root.resizable(True, True)

        self.tipo_cambio = 8.00
        self.nombre_proyecto = "Prorrateo_Nuevo"
        self.linea_activa = "Pandora"
        self.productos_catalogo = []
        self.productos_lote = []
        self.enlaces_temp = []
        self.ruta_imagen_temp = ""
        self.imagenes_cache = {}

        self.tiendas_disponibles = list(PROVEEDORES_BASE)
        self.cargar_proveedores_custom()

        self.params_prorrateo = {
            "pct_comision": 0.0,
            "costo_envio_usd_por_g": 0.0,
            "costo_envio_mx_por_g": 0.0,
            "val_total_prod_usd": 0.0,
            "peso_total_g": 0.0,
        }

        self.crear_interfaz()
        self.cargar_catalogo_global()
        self.cargar_ultimo_estado_automatico()

    def cargar_proveedores_custom(self):
        if os.path.exists(PROVEEDORES_FILE):
            try:
                with open(PROVEEDORES_FILE, "r", encoding="utf-8") as f:
                    provs = json.load(f)
                    for p in provs:
                        if p and p not in self.tiendas_disponibles:
                            self.tiendas_disponibles.append(p)
            except Exception as e:
                print(f"Error al cargar proveedores personalizados: {e}")

    def guardar_proveedores_custom(self):
        try:
            provs_extra = [p for p in self.tiendas_disponibles if p not in PROVEEDORES_BASE]
            with open(PROVEEDORES_FILE, "w", encoding="utf-8") as f:
                json.dump(provs_extra, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Error al guardar proveedores personalizados: {e}")

    def registrar_nuevo_proveedor_si_no_existe(self, nombre):
        nombre_clean = nombre.strip()
        if nombre_clean and nombre_clean not in self.tiendas_disponibles:
            self.tiendas_disponibles.append(nombre_clean)
            self.guardar_proveedores_custom()
            self.actualizar_combos_proveedores()

    def actualizar_combos_proveedores(self):
        if hasattr(self, 'combo_tienda_link'):
            self.combo_tienda_link.config(values=self.tiendas_disponibles)

    def crear_interfaz(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=60)

        # --- BARRA SUPERIOR ---
        frame_proyecto = ttk.LabelFrame(
            self.root, text=" 📂 Gestión y Carga de Prorrateo / Proyecto ", padding=10
        )
        frame_proyecto.pack(fill="x", padx=15, pady=5)

        ttk.Label(frame_proyecto, text="Prorrateo:").grid(row=0, column=0, sticky="w", pady=4)
        self.entry_nombre_proyecto = ttk.Entry(frame_proyecto, width=18)
        self.entry_nombre_proyecto.insert(0, self.nombre_proyecto)
        self.entry_nombre_proyecto.grid(row=0, column=1, padx=4, pady=4)

        ttk.Button(frame_proyecto, text="💾 Guardar Prorrateo", command=self.guardar_archivo_prorrateo).grid(row=0, column=2, padx=4, pady=4)
        ttk.Button(frame_proyecto, text="📂 Cargar Prorrateo Guardado", command=self.cargar_archivo_prorrateo).grid(row=0, column=3, padx=4, pady=4)
        ttk.Button(frame_proyecto, text="📄 Nuevo Prorrateo", command=self.nuevo_prorrateo).grid(row=0, column=4, padx=4, pady=4)

        # Opciones de Exportación PDF
        self.var_ocultar_precios_pdf = tk.BooleanVar(value=False)
        self.chk_ocultar_precios = ttk.Checkbutton(
            frame_proyecto, text="🙈 Ocultar Precios en PDF", variable=self.var_ocultar_precios_pdf
        )
        self.chk_ocultar_precios.grid(row=0, column=5, padx=8, pady=4)

        ttk.Button(frame_proyecto, text="📕 Exportar Catálogo PDF (Línea Activa)", command=self.generar_catalogo_pdf).grid(row=0, column=6, padx=8, pady=4)

        # --- BARRA DE ESTADO ---
        frame_status = ttk.Frame(self.root, padding=5)
        frame_status.pack(fill="x", padx=15, pady=2)

        self.lbl_status_comision = ttk.Label(
            frame_status, text="Parámetros Activos -> Comisión Tarjeta: 0.00%", font=("Arial", 9, "bold"), foreground="#1F4E79"
        )
        self.lbl_status_comision.pack(side="left", padx=10)

        self.lbl_status_flete = ttk.Label(
            frame_status, text="| Flete USD/Gramo: $0.000 | Flete MX/Gramo: Q0.000", font=("Arial", 9, "bold"), foreground="#1F4E79"
        )
        self.lbl_status_flete.pack(side="left", padx=10)

        # --- PESTAÑAS ---
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=5)

        self.tab_lote = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_lote, text=" 📦 Prorrateo de Pedido en Camino ")

        self.tab_rapido = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_rapido, text=" ⚡ Cotizador Rápido & Catálogo ")

        self.construir_tab_lote()
        self.construir_tab_rapido()

    def construir_tab_lote(self):
        frame_general = ttk.LabelFrame(self.tab_lote, text=" 1. Costos Globales del Pedido ", padding=10)
        frame_general.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_general, text="Pago Tarjeta ($USD):").grid(row=0, column=0, sticky="w", pady=4)
        self.entry_pago_tarjeta = ttk.Entry(frame_general, width=12)
        self.entry_pago_tarjeta.grid(row=0, column=1, padx=5, pady=4)
        self.entry_pago_tarjeta.bind("<KeyRelease>", lambda e: self.recalcular_lote_y_parametros())

        ttk.Label(frame_general, text="Costo Flete ($USD):").grid(row=0, column=2, sticky="w", pady=4)
        self.entry_costo_envio = ttk.Entry(frame_general, width=12)
        self.entry_costo_envio.grid(row=0, column=3, padx=5, pady=4)
        self.entry_costo_envio.bind("<KeyRelease>", lambda e: self.recalcular_lote_y_parametros())

        ttk.Label(frame_general, text="Piezas Totales Lote:").grid(row=0, column=4, sticky="w", pady=4)
        self.entry_total_piezas = ttk.Entry(frame_general, width=10)
        self.entry_total_piezas.grid(row=0, column=5, padx=5, pady=4)
        self.entry_total_piezas.bind("<KeyRelease>", lambda e: self.recalcular_lote_y_parametros())

        frame_ingreso_lote = ttk.LabelFrame(self.tab_lote, text=" 2. Agregar Producto al Lote ", padding=10)
        frame_ingreso_lote.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_ingreso_lote, text="Nombre:").grid(row=0, column=0, padx=4)
        self.entry_lote_nombre = ttk.Entry(frame_ingreso_lote, width=18)
        self.entry_lote_nombre.grid(row=0, column=1, padx=4)

        ttk.Label(frame_ingreso_lote, text="Cant:").grid(row=0, column=2, padx=4)
        self.entry_lote_cant = ttk.Entry(frame_ingreso_lote, width=6)
        self.entry_lote_cant.insert(0, "1")
        self.entry_lote_cant.grid(row=0, column=3, padx=4)

        ttk.Label(frame_ingreso_lote, text="Precio Base ($USD):").grid(row=0, column=4, padx=4)
        self.entry_lote_precio = ttk.Entry(frame_ingreso_lote, width=10)
        self.entry_lote_precio.grid(row=0, column=5, padx=4)

        ttk.Label(frame_ingreso_lote, text="Peso Unit. (g):").grid(row=0, column=6, padx=4)
        self.entry_lote_peso = ttk.Entry(frame_ingreso_lote, width=10)
        self.entry_lote_peso.grid(row=0, column=7, padx=4)

        ttk.Button(frame_ingreso_lote, text="➕ Añadir", command=self.agregar_prod_lote).grid(row=0, column=8, padx=8)

        frame_tabla_lote = ttk.Frame(self.tab_lote, padding=5)
        frame_tabla_lote.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("nombre", "cant", "precio_usd", "peso_g", "comision_gtq", "envio_usd_gtq", "envio_mx_gtq", "costo_unit_gtq", "precio_venta_gtq")
        self.tabla_lote = ttk.Treeview(frame_tabla_lote, columns=cols, show="headings", height=8)
        self.tabla_lote.heading("nombre", text="Producto")
        self.tabla_lote.heading("cant", text="Cant")
        self.tabla_lote.heading("precio_usd", text="Base ($)")
        self.tabla_lote.heading("peso_g", text="Peso (g)")
        self.tabla_lote.heading("comision_gtq", text="Comisión (Q)")
        self.tabla_lote.heading("envio_usd_gtq", text="Envío $ (Q)")
        self.tabla_lote.heading("envio_mx_gtq", text="Envío MX (Q)")
        self.tabla_lote.heading("costo_unit_gtq", text="Costo Ud (Q)")
        self.tabla_lote.heading("precio_venta_gtq", text="P. Venta Sugerido (Q)")

        self.tabla_lote.column("nombre", width=160)
        self.tabla_lote.column("cant", width=45, anchor="center")
        self.tabla_lote.column("precio_usd", width=70, anchor="e")
        self.tabla_lote.column("peso_g", width=75, anchor="center")
        self.tabla_lote.column("comision_gtq", width=80, anchor="e")
        self.tabla_lote.column("envio_usd_gtq", width=80, anchor="e")
        self.tabla_lote.column("envio_mx_gtq", width=80, anchor="e")
        self.tabla_lote.column("costo_unit_gtq", width=90, anchor="e")
        self.tabla_lote.column("precio_venta_gtq", width=120, anchor="e")

        self.tabla_lote.pack(fill="both", expand=True)

        btn_eliminar_lote = ttk.Button(frame_tabla_lote, text="❌ Eliminar de Lote", command=self.eliminar_prod_lote)
        btn_eliminar_lote.pack(anchor="e", pady=3)

    def agregar_prod_lote(self):
        try:
            nombre = self.entry_lote_nombre.get().strip()
            cant = int(self.entry_lote_cant.get())
            precio = float(self.entry_lote_precio.get())
            peso = float(self.entry_lote_peso.get())

            if not nombre:
                return

            subtotal = cant * precio
            peso_tot = cant * peso

            self.productos_lote.append({
                "nombre": nombre,
                "cant": cant,
                "precio": precio,
                "subtotal": subtotal,
                "peso": peso,
                "peso_tot": peso_tot,
            })

            self.recalcular_lote_y_parametros()
            self.entry_lote_nombre.delete(0, tk.END)
            self.entry_lote_precio.delete(0, tk.END)
            self.entry_lote_peso.delete(0, tk.END)

        except ValueError:
            messagebox.showerror("Error", "Revisa los valores numéricos ingresados.")

    def eliminar_prod_lote(self):
        seleccion = self.tabla_lote.selection()
        if not seleccion:
            return
        idx = self.tabla_lote.index(seleccion[0])
        del self.productos_lote[idx]
        self.recalcular_lote_y_parametros()

    def recalcular_lote_y_parametros(self):
        for item in self.tabla_lote.get_children():
            self.tabla_lote.delete(item)

        val_total_prod_usd = sum(p.get("subtotal", p.get("cant", 1) * p.get("precio", 0)) for p in self.productos_lote)
        peso_total_g = sum(p.get("peso_tot", p.get("cant", 1) * p.get("peso", 0)) for p in self.productos_lote)

        try:
            pago_tarjeta = float(self.entry_pago_tarjeta.get()) if self.entry_pago_tarjeta.get() else 0.0
            costo_envio = float(self.entry_costo_envio.get()) if self.entry_costo_envio.get() else 0.0
        except ValueError:
            pago_tarjeta = 0.0
            costo_envio = 0.0

        gastos_tarjeta_usd = pago_tarjeta - val_total_prod_usd
        flete_mx_total_gtq = self.calcular_costo_envio_mexico(peso_total_g)

        pct_comision = (gastos_tarjeta_usd / val_total_prod_usd) if val_total_prod_usd > 0 else 0.0
        costo_envio_usd_por_g = (costo_envio / peso_total_g) if peso_total_g > 0 else 0.0
        costo_envio_mx_por_g = (flete_mx_total_gtq / peso_total_g) if peso_total_g > 0 else 0.0

        self.params_prorrateo = {
            "pct_comision": pct_comision,
            "costo_envio_usd_por_g": costo_envio_usd_por_g,
            "costo_envio_mx_por_g": costo_envio_mx_por_g,
            "val_total_prod_usd": val_total_prod_usd,
            "peso_total_g": peso_total_g,
        }

        self.actualizar_labels_status()

        for p in self.productos_lote:
            cant = p.get("cant", 1)
            precio_unit_usd = p.get("precio", p.get("precio_unit_usd", 0.0))
            subtotal_usd = p.get("subtotal", cant * precio_unit_usd)
            peso_unit_g = p.get("peso", p.get("peso_unit_g", 0.0))
            peso_lote_g = p.get("peso_tot", cant * peso_unit_g)

            part_valor = subtotal_usd / val_total_prod_usd if val_total_prod_usd > 0 else 0
            comision_unit_usd = (gastos_tarjeta_usd * part_valor) / cant if cant > 0 else 0

            part_peso = peso_lote_g / peso_total_g if peso_total_g > 0 else 0
            envio_usd_unit_usd = (costo_envio * part_peso) / cant if cant > 0 else 0
            envio_mx_unit_gtq = (flete_mx_total_gtq * part_peso) / cant if cant > 0 else 0

            costo_base_gtq = precio_unit_usd * self.tipo_cambio
            comision_gtq = comision_unit_usd * self.tipo_cambio
            envio_usd_gtq = envio_usd_unit_usd * self.tipo_cambio

            costo_unit_gtq = costo_base_gtq + comision_gtq + envio_usd_gtq + envio_mx_unit_gtq
            precio_venta_gtq = self.calcular_precio_venta(costo_unit_gtq)

            self.tabla_lote.insert(
                "",
                tk.END,
                values=(
                    p.get("nombre"),
                    cant,
                    f"${precio_unit_usd:,.2f}",
                    f"{peso_unit_g:,.1f}g",
                    f"Q{comision_gtq:,.2f}",
                    f"Q{envio_usd_gtq:,.2f}",
                    f"Q{envio_mx_unit_gtq:,.2f}",
                    f"Q{costo_unit_gtq:,.2f}",
                    f"Q{precio_venta_gtq:,.2f}",
                ),
            )

        self.recalcular_precios_catalogo()
        self.guardar_estado_autosave()
        self.calcular_cotizacion_rapida()

    def obtener_opcion_mas_economica(self, enlaces, peso_g):
        """
        Determina la opción más económica EXCLUYENDO a 'Ako' (por ser compra al por mayor).
        """
        if not enlaces:
            return None

        enlaces_evaluables = [x for x in enlaces if "Ako" not in x.get("nombre_tienda", "")]
        if not enlaces_evaluables:
            enlaces_evaluables = enlaces

        opciones_calc = [(item, self.calcular_costo_y_precio_por_tienda(item.get("nombre_tienda", ""), item.get("precio_usd", 0.0), peso_g)[0]) for item in enlaces_evaluables]
        return min(opciones_calc, key=lambda x: x[1])[0]

    def recalcular_precios_catalogo(self):
        for p in self.productos_catalogo:
            enlaces = p.get("enlaces", [])
            tienda_seleccionada = p.get("tienda_seleccionada", p.get("tienda_origen", "Friend"))

            if enlaces:
                opcion_barata = self.obtener_opcion_mas_economica(enlaces, p.get("peso_g", 0.0))
                p["tienda_mas_barata"] = opcion_barata.get("nombre_tienda", "") if opcion_barata else tienda_seleccionada
                
                # Buscar datos de la tienda seleccionada
                item_sel = next((x for x in enlaces if x.get("nombre_tienda") == tienda_seleccionada), opcion_barata if opcion_barata else enlaces[0])
                p["precio_usd"] = item_sel.get("precio_usd", 0.0)
                p["tienda_origen"] = item_sel.get("nombre_tienda", "")
                p["link_compra"] = item_sel.get("url", "")
            else:
                p["tienda_mas_barata"] = tienda_seleccionada

            precio_usd = float(p.get("precio_usd", 0.0))
            peso_g = float(p.get("peso_g", 0.0))

            costo_total_gtq, precio_sugerido = self.calcular_costo_y_precio_por_tienda(p.get("tienda_origen", tienda_seleccionada), precio_usd, peso_g)

            p["costo_unit_gtq"] = costo_total_gtq
            p["precio_sugerido_gtq"] = precio_sugerido
            if "precio_final_gtq" not in p or p["precio_final_gtq"] <= 0:
                p["precio_final_gtq"] = precio_sugerido

        self.guardar_catalogo_global()
        self.renderizar_tabla_catalogo()

    def calcular_costo_y_precio_por_tienda(self, nombre_tienda, precio_usd, peso_g):
        """
        Fórmulas específicas por Proveedor / Tienda:
        - Friend: Costo = (Precio USD * 8) + 1.75
        - Ako: Costo = (Precio USD * 1.05 * 8) + (Peso * 0.30) + (Peso * 0.10)
        - Todos los demás nuevos proveedores: Costo = Base GTQ + Comisión Tarjeta + Flete USD + Flete México
        - Precio Sugerido en todos: Redondear a múltiplo de Q5 de: (Costo Total * 2) + 10
        """
        if "Friend" in nombre_tienda:
            costo_total_gtq = (precio_usd * 8.0) + 1.75
            base_sugerido = (costo_total_gtq * 2) + 10.0
            precio_sugerido = float(math.ceil(base_sugerido / 5.0) * 5)
            return costo_total_gtq, precio_sugerido
        elif "Ako" in nombre_tienda:
            costo_prod_gtq = (precio_usd * 1.05) * self.tipo_cambio
            flete_china_gtq = peso_g * 0.30
            flete_mexico_gtq = peso_g * 0.10
            costo_total_gtq = costo_prod_gtq + flete_china_gtq + flete_mexico_gtq
            base_sugerido = (costo_total_gtq * 2) + 10.0
            precio_sugerido = float(math.ceil(base_sugerido / 5.0) * 5)
            return costo_total_gtq, precio_sugerido
        else:
            pct_comision = self.params_prorrateo.get("pct_comision", 0.0)
            costo_flete_g = self.params_prorrateo.get("costo_envio_usd_por_g", 0.0)
            costo_flete_mx_g = self.params_prorrateo.get("costo_envio_mx_por_g", 0.0)

            base_gtq = precio_usd * self.tipo_cambio
            comision_gtq = (precio_usd * pct_comision) * self.tipo_cambio
            envio_usd_gtq = (peso_g * costo_flete_g) * self.tipo_cambio

            if costo_flete_mx_g > 0:
                envio_mx_gtq = peso_g * costo_flete_mx_g
            else:
                envio_mx_gtq = self.calcular_costo_envio_mexico(peso_g)

            costo_total_gtq = base_gtq + comision_gtq + envio_usd_gtq + envio_mx_gtq
            precio_sugerido = self.calcular_precio_venta(costo_total_gtq)
            return costo_total_gtq, precio_sugerido

    def actualizar_labels_status(self):
        pct = self.params_prorrateo.get("pct_comision", 0.0) * 100
        flete = self.params_prorrateo.get("costo_envio_usd_por_g", 0.0)
        flete_mx = self.params_prorrateo.get("costo_envio_mx_por_g", 0.0)
        self.lbl_status_comision.config(text=f"Parámetros Activos -> Comisión Tarjeta: {pct:.2f}%")
        self.lbl_status_flete.config(text=f"| Flete USD/Gramo: ${flete:.4f} | Flete MX/Gramo: Q{flete_mx:.3f}")

    def construir_tab_rapido(self):
        frame_linea = ttk.LabelFrame(self.tab_rapido, text=" 💎 Selección de Línea de Producto ", padding=10)
        frame_linea.pack(fill="x", padx=15, pady=4)

        ttk.Label(frame_linea, text="Elige la Línea:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", padx=5)

        self.combo_linea_activa = ttk.Combobox(
            frame_linea, values=LINEAS_LISTA, width=22, state="readonly", font=("Arial", 10, "bold")
        )
        self.combo_linea_activa.set(self.linea_activa)
        self.combo_linea_activa.grid(row=0, column=1, padx=10)
        self.combo_linea_activa.bind("<<ComboboxSelected>>", self.al_cambiar_linea)

        self.lbl_linea_info = ttk.Label(
            frame_linea, text="🏷️ Catálogo y Categorías aislados para la línea de productos seleccionada.", font=("Arial", 9, "italic"), foreground="#1F4E79"
        )
        self.lbl_linea_info.grid(row=0, column=2, padx=15)

        frame_rapido = ttk.LabelFrame(self.tab_rapido, text=" 1. Cotizar y Configurar Producto ", padding=12)
        frame_rapido.pack(fill="x", padx=15, pady=4)

        ttk.Label(frame_rapido, text="Categoría:").grid(row=0, column=0, sticky="w", pady=4)
        self.entry_rap_tipo = ttk.Combobox(frame_rapido, values=CATEGORIAS_POR_LINEA[self.linea_activa], width=25)
        self.entry_rap_tipo.set(CATEGORIAS_POR_LINEA[self.linea_activa][0])
        self.entry_rap_tipo.grid(row=0, column=1, padx=5, pady=4)

        ttk.Label(frame_rapido, text="No. Posición / Item:").grid(row=0, column=2, sticky="w", pady=4)
        self.entry_rap_num = ttk.Entry(frame_rapido, width=12)
        self.entry_rap_num.insert(0, "101")
        self.entry_rap_num.grid(row=0, column=3, padx=5, pady=4)

        frame_btns_img = ttk.Frame(frame_rapido)
        frame_btns_img.grid(row=0, column=4, padx=10, pady=4)

        ttk.Button(frame_btns_img, text="📷 Archivo Local", command=self.seleccionar_imagen_local).pack(side="left", padx=2)
        ttk.Button(frame_btns_img, text="🌐 Enlace Drive / URL", command=self.ingresar_url_drive).pack(side="left", padx=2)

        self.lbl_preview_img = ttk.Label(frame_rapido, text="Sin Imagen", width=14, relief="groove", anchor="center", cursor="hand2")
        self.lbl_preview_img.grid(row=0, column=5, rowspan=5, padx=10, pady=4)
        self.lbl_preview_img.bind("<Button-1>", lambda e: self.mostrar_imagen_ampliada(self.ruta_imagen_temp))

        ttk.Label(frame_rapido, text="Precio Base ($USD):").grid(row=1, column=0, sticky="w", pady=4)
        self.entry_rapido_precio = ttk.Entry(frame_rapido, width=12)
        self.entry_rapido_precio.grid(row=1, column=1, padx=5, pady=4)
        self.entry_rapido_precio.bind("<KeyRelease>", lambda e: self.calcular_cotizacion_rapida())

        ttk.Label(frame_rapido, text="Peso Unit. (Gramos):").grid(row=1, column=2, sticky="w", pady=4)
        self.entry_rapido_peso = ttk.Entry(frame_rapido, width=10)
        self.entry_rapido_peso.grid(row=1, column=3, padx=5, pady=4, sticky="w")
        self.entry_rapido_peso.bind("<KeyRelease>", lambda e: self.calcular_cotizacion_rapida())

        frame_enlaces = ttk.LabelFrame(frame_rapido, text=" 🛍️ Comparador de Tiendas / Proveedores para este Producto ", padding=6)
        frame_enlaces.grid(row=2, column=0, columnspan=5, sticky="ew", pady=5, padx=2)

        ttk.Label(frame_enlaces, text="Tienda / Proveedor:").grid(row=0, column=0, padx=2, sticky="w")
        self.combo_tienda_link = ttk.Combobox(frame_enlaces, values=self.tiendas_disponibles, width=20)
        self.combo_tienda_link.set("Friend")
        self.combo_tienda_link.grid(row=0, column=1, padx=2)

        ttk.Label(frame_enlaces, text="Precio ($USD):").grid(row=0, column=2, padx=2, sticky="w")
        self.entry_precio_tienda_link = ttk.Entry(frame_enlaces, width=10)
        self.entry_precio_tienda_link.grid(row=0, column=3, padx=2)

        ttk.Label(frame_enlaces, text="URL Enlace:").grid(row=0, column=4, padx=2, sticky="w")
        self.entry_url_tienda_link = ttk.Entry(frame_enlaces, width=25)
        self.entry_url_tienda_link.grid(row=0, column=5, padx=2)

        ttk.Button(frame_enlaces, text="➕ Agregar Opción", command=self.agregar_tienda_comparativa).grid(row=0, column=6, padx=6)

        self.listbox_enlaces = tk.Listbox(frame_enlaces, height=3, width=80)
        self.listbox_enlaces.grid(row=1, column=0, columnspan=6, pady=4, padx=2, sticky="ew")

        ttk.Button(frame_enlaces, text="❌ Quitar Opción", command=self.eliminar_tienda_comparativa).grid(row=1, column=6, padx=6, sticky="n")

        ttk.Label(frame_rapido, text="Precio Venta Real (Q GTQ):", font=("Arial", 9, "bold"), foreground="#0056b3").grid(row=3, column=0, sticky="w", pady=4)
        self.entry_precio_venta_custom = ttk.Entry(frame_rapido, width=25, font=("Arial", 10, "bold"))
        self.entry_precio_venta_custom.grid(row=3, column=1, padx=5, pady=4)

        self.btn_guardar_cat = ttk.Button(frame_rapido, text="➕ Guardar en Catálogo de Línea", command=self.guardar_producto_catalogo)
        self.btn_guardar_cat.grid(row=3, column=2, columnspan=2, padx=5, pady=4)

        self.btn_cancelar_edit = ttk.Button(frame_rapido, text="❌ Cancelar Edición", command=self.cancelar_edicion)
        self.btn_cancelar_edit.grid_remove()

        frame_res_rapido = ttk.LabelFrame(self.tab_rapido, text=" 2. Desglose en base a Reglas de la Tienda Seleccionada ", padding=8)
        frame_res_rapido.pack(fill="x", padx=15, pady=3)

        self.lbl_rap_base_gtq = ttk.Label(frame_res_rapido, text="Costo Base: Q0.00", font=("Arial", 10))
        self.lbl_rap_base_gtq.grid(row=0, column=0, sticky="w", pady=2, padx=10)

        self.lbl_rap_comision = ttk.Label(frame_res_rapido, text="Comisión Tarjeta: Q0.00", font=("Arial", 10))
        self.lbl_rap_comision.grid(row=0, column=1, sticky="w", pady=2, padx=10)

        self.lbl_rap_envio_usd = ttk.Label(frame_res_rapido, text="Flete USD: Q0.00", font=("Arial", 10))
        self.lbl_rap_envio_usd.grid(row=1, column=0, sticky="w", pady=2, padx=10)

        self.lbl_rap_envio_mx = ttk.Label(frame_res_rapido, text="Flete México: Q0.00", font=("Arial", 10))
        self.lbl_rap_envio_mx.grid(row=1, column=1, sticky="w", pady=2, padx=10)

        self.lbl_rap_costo_total = ttk.Label(frame_res_rapido, text="COSTO TOTAL: Q0.00 GTQ", font=("Arial", 10, "bold"), foreground="#C00000")
        self.lbl_rap_costo_total.grid(row=2, column=0, sticky="w", pady=4, padx=10)

        self.lbl_rap_precio_venta = ttk.Label(frame_res_rapido, text="PRECIO SUGERIDO: Q0.00 GTQ", font=("Arial", 11, "bold"), foreground="#008000")
        self.lbl_rap_precio_venta.grid(row=2, column=1, sticky="w", pady=4, padx=10)

        self.frame_tabla_cat = ttk.LabelFrame(self.tab_rapido, text=f" 3. Catálogo Exclusivo - Línea: {self.linea_activa} ", padding=8)
        self.frame_tabla_cat.pack(fill="both", expand=True, padx=15, pady=5)

        frame_busqueda = ttk.LabelFrame(self.frame_tabla_cat, text=" 🔎 Buscador en Catálogo (por Número o Imagen) ", padding=5)
        frame_busqueda.pack(fill="x", pady=4)

        ttk.Label(frame_busqueda, text="Buscar No. Item:").grid(row=0, column=0, padx=4, sticky="w")
        self.entry_buscar_num = ttk.Entry(frame_busqueda, width=12)
        self.entry_buscar_num.grid(row=0, column=1, padx=4)
        self.entry_buscar_num.bind("<KeyRelease>", lambda e: self.buscar_por_numero())

        ttk.Button(frame_busqueda, text="🔍 Buscar No.", command=self.buscar_por_numero).grid(row=0, column=2, padx=4)
        ttk.Label(frame_busqueda, text=" | ").grid(row=0, column=3, padx=2)
        ttk.Button(frame_busqueda, text="🖼️ Buscar por Imagen / Foto", command=self.buscar_por_imagen).grid(row=0, column=4, padx=6)
        ttk.Button(frame_busqueda, text="🔄 Restablecer Filtros", command=self.limpiar_busqueda).grid(row=0, column=5, padx=8)

        ttk.Label(frame_busqueda, text="Categoría:").grid(row=0, column=6, padx=4, sticky="w")
        self.combo_filtro_cat = ttk.Combobox(
            frame_busqueda, values=["Todos"] + CATEGORIAS_POR_LINEA[self.linea_activa], width=18, state="readonly"
        )
        self.combo_filtro_cat.set("Todos")
        self.combo_filtro_cat.grid(row=0, column=7, padx=4)
        self.combo_filtro_cat.bind("<<ComboboxSelected>>", lambda e: self.renderizar_tabla_catalogo())

        frame_acciones_tabla = ttk.Frame(self.frame_tabla_cat)
        frame_acciones_tabla.pack(fill="x", pady=3)

        ttk.Button(frame_acciones_tabla, text="⚖️ Comparar Tiendas del Producto", command=self.comparar_tiendas_producto).pack(side="right", padx=4)
        ttk.Button(frame_acciones_tabla, text="🔍 Ver Imagen HD", command=self.ver_imagen_seleccionada).pack(side="right", padx=4)
        ttk.Button(frame_acciones_tabla, text="🛒 Abrir Enlace de Compra", command=self.abrir_link_compra).pack(side="right", padx=4)
        ttk.Button(frame_acciones_tabla, text="✏️ Cargar para Editar", command=self.cargar_producto_para_editar).pack(side="right", padx=4)

        columnas = ("num", "tipo", "tienda_origen", "precio_usd", "peso_g", "costo_unit_gtq", "precio_sugerido_gtq", "precio_final_gtq", "ganancia_gtq", "tienda_mas_barata")
        self.tabla_cat = ttk.Treeview(self.frame_tabla_cat, columns=columnas, show="tree headings", height=5)
        self.tabla_cat.heading("#0", text="Imagen (Click)")
        self.tabla_cat.heading("num", text="No.")
        self.tabla_cat.heading("tipo", text="Categoría")
        self.tabla_cat.heading("tienda_origen", text="Tienda Elección")
        self.tabla_cat.heading("precio_usd", text="Base ($)")
        self.tabla_cat.heading("peso_g", text="Peso (g)")
        self.tabla_cat.heading("costo_unit_gtq", text="Costo Ud (Q)")
        self.tabla_cat.heading("precio_sugerido_gtq", text="P. Sugerido (Q)")
        self.tabla_cat.heading("precio_final_gtq", text="P. Venta Real (Q)")
        self.tabla_cat.heading("ganancia_gtq", text="Ganancia Neta (Q / %)")
        self.tabla_cat.heading("tienda_mas_barata", text="Opción Más Económica")

        self.tabla_cat.column("#0", width=85, anchor="center")
        self.tabla_cat.column("num", width=40, anchor="center")
        self.tabla_cat.column("tipo", width=110)
        self.tabla_cat.column("tienda_origen", width=110, anchor="center")
        self.tabla_cat.column("precio_usd", width=65, anchor="e")
        self.tabla_cat.column("peso_g", width=60, anchor="center")
        self.tabla_cat.column("costo_unit_gtq", width=80, anchor="e")
        self.tabla_cat.column("precio_sugerido_gtq", width=95, anchor="e")
        self.tabla_cat.column("precio_final_gtq", width=100, anchor="e")
        self.tabla_cat.column("ganancia_gtq", width=130, anchor="e")
        self.tabla_cat.column("tienda_mas_barata", width=150, anchor="center")

        self.tabla_cat.pack(fill="both", expand=True, pady=3)
        self.tabla_cat.bind("<Double-1>", lambda e: self.cargar_producto_para_editar())

        btn_eliminar_cat = ttk.Button(self.frame_tabla_cat, text="❌ Eliminar de Catálogo", command=self.eliminar_prod_catalogo)
        btn_eliminar_cat.pack(anchor="e", pady=2)

    def ingresar_url_drive(self):
        def guardar_url():
            url = entry_url.get().strip()
            if url:
                ruta_local = descargar_y_guardar_localmente(url)
                if ruta_local:
                    self.ruta_imagen_temp = ruta_local
                else:
                    self.ruta_imagen_temp = url
                self.mostrar_vista_previa(self.ruta_imagen_temp)
                top.destroy()

        top = tk.Toplevel(self.root)
        top.title("Ingresar Enlace de Google Drive / Imagen")
        top.geometry("500x150")
        top.grab_set()

        ttk.Label(top, text="Pega aquí el enlace de la imagen en Google Drive:", font=("Arial", 10, "bold")).pack(pady=10)
        entry_url = ttk.Entry(top, width=60)
        entry_url.pack(padx=10, pady=5)
        entry_url.focus()

        ttk.Button(top, text="✅ Aceptar y Previsualizar", command=guardar_url).pack(pady=10)

    def seleccionar_imagen_local(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.bmp *.webp")],
            title="Seleccionar Imagen del Producto",
        )
        if filepath:
            self.ruta_imagen_temp = filepath
            self.mostrar_vista_previa(filepath)

    def mostrar_vista_previa(self, ruta_o_url):
        img_pil = descargar_imagen(ruta_o_url)
        if img_pil:
            try:
                img_pil.thumbnail((70, 70))
                self.photo_preview = ImageTk.PhotoImage(img_pil)
                self.lbl_preview_img.config(image=self.photo_preview, text="")
            except Exception:
                self.lbl_preview_img.config(image="", text="Error Img")
        else:
            self.lbl_preview_img.config(image="", text="Sin Imagen")

    def buscar_por_numero(self):
        query = self.entry_buscar_num.get().strip()
        if not query:
            self.renderizar_tabla_catalogo()
            return
        self.renderizar_tabla_catalogo(filtro_num=query)

    def buscar_por_imagen(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.bmp *.webp")],
            title="Seleccionar Imagen para Buscar en Catálogo",
        )
        if not filepath:
            return

        nombre_archivo = os.path.basename(filepath).lower()
        tamano_archivo = os.path.getsize(filepath)

        prod_coincidente = None
        prods_linea = [p for p in self.productos_catalogo if p.get("linea", p.get("tienda", "Pandora")) == self.linea_activa]

        for p in prods_linea:
            img_p = p.get("imagen", "")
            if img_p and not img_p.startswith("http"):
                if os.path.exists(img_p) and os.path.basename(img_p).lower() == nombre_archivo:
                    prod_coincidente = p
                    break

        if not prod_coincidente:
            for p in prods_linea:
                img_p = p.get("imagen", "")
                if img_p and not img_p.startswith("http"):
                    if os.path.exists(img_p) and os.path.getsize(img_p) == tamano_archivo:
                        prod_coincidente = p
                        break

        if prod_coincidente:
            num_item = str(prod_coincidente.get("num", ""))
            self.entry_buscar_num.delete(0, tk.END)
            self.entry_buscar_num.insert(0, num_item)
            self.renderizar_tabla_catalogo(filtro_num=num_item)

            for child in self.tabla_cat.get_children():
                vals = self.tabla_cat.item(child, "values")
                if vals[0] == f"{num_item}":
                    self.tabla_cat.selection_set(child)
                    self.tabla_cat.focus(child)
                    self.tabla_cat.see(child)
                    break

            messagebox.showinfo("Encontrado", f"🟢 ¡Producto Ubicado!\n\nCorresponde al Código {num_item} ({prod_coincidente.get('tipo')}).")
        else:
            messagebox.showwarning("No Encontrado", "No se encontró ningún producto registrado localmente con esa imagen.")

    def limpiar_busqueda(self):
        self.entry_buscar_num.delete(0, tk.END)
        self.combo_filtro_cat.set("Todos")
        self.renderizar_tabla_catalogo()

    def agregar_tienda_comparativa(self):
        nombre = self.combo_tienda_link.get().strip()
        url = self.entry_url_tienda_link.get().strip()
        try:
            precio = float(self.entry_precio_tienda_link.get())
        except ValueError:
            messagebox.showerror("Error", "Ingresa un precio numérico válido para la tienda.")
            return

        if not nombre:
            nombre = "Proveedor General"

        self.registrar_nuevo_proveedor_si_no_existe(nombre)

        self.enlaces_temp = [x for x in self.enlaces_temp if x.get("nombre_tienda") != nombre]

        self.enlaces_temp.append({
            "nombre_tienda": nombre,
            "precio_usd": precio,
            "url": url,
        })

        self.actualizar_listbox_enlaces()

        self.entry_precio_tienda_link.delete(0, tk.END)
        self.entry_url_tienda_link.delete(0, tk.END)

        self.calcular_cotizacion_rapida()

    def eliminar_tienda_comparativa(self):
        seleccion = self.listbox_enlaces.curselection()
        if not seleccion:
            return
        idx = seleccion[0]
        del self.enlaces_temp[idx]
        self.actualizar_listbox_enlaces()
        self.calcular_cotizacion_rapida()

    def actualizar_listbox_enlaces(self):
        self.listbox_enlaces.delete(0, tk.END)
        if not self.enlaces_temp:
            return

        try:
            peso_g = float(self.entry_rapido_peso.get()) if self.entry_rapido_peso.get() else 0.0
        except ValueError:
            peso_g = 0.0

        opcion_barata = self.obtener_opcion_mas_economica(self.enlaces_temp, peso_g)

        for item in self.enlaces_temp:
            es_barata = " 🏷️ (MÁS ECONÓMICA)" if (opcion_barata and item == opcion_barata) else ""
            c_tot, p_sug = self.calcular_costo_y_precio_por_tienda(item.get("nombre_tienda", ""), item.get("precio_usd", 0.0), peso_g)
            self.listbox_enlaces.insert(
                tk.END,
                f"🏪 {item['nombre_tienda']} | USD: ${item['precio_usd']:,.2f} | Costo GTQ: Q{c_tot:,.2f}{es_barata}"
            )

    def comparar_tiendas_producto(self):
        seleccion = self.tabla_cat.selection()
        if not seleccion:
            messagebox.showwarning("Atención", "Selecciona un producto de la tabla para comparar sus tiendas.")
            return

        item_vals = self.tabla_cat.item(seleccion[0], "values")
        num_sel = int(item_vals[0].replace("#", ""))
        tipo_sel = item_vals[1]

        prod_encontrado = None
        for p in self.productos_catalogo:
            linea_p = p.get("linea", p.get("tienda", "Pandora"))
            if linea_p == self.linea_activa and p.get("num") == num_sel and p.get("tipo") == tipo_sel:
                prod_encontrado = p
                break

        if not prod_encontrado:
            return

        enlaces = prod_encontrado.get("enlaces", [])
        if not enlaces:
            messagebox.showinfo("Comparador", f"El producto Código {num_sel} sólo tiene registrada una tienda con costo de ${prod_encontrado.get('precio_usd', 0.0):,.2f} USD.")
            return

        top = tk.Toplevel(self.root)
        top.title(f"⚖️ Comparativa de Tiendas - Código {num_sel} ({tipo_sel})")
        top.geometry("780x450")
        top.resizable(True, True)
        top.grab_set()

        ttk.Label(top, text=f"Comparativa de Proveedores para {tipo_sel} Código {num_sel} (Línea {self.linea_activa})", font=("Arial", 12, "bold"), foreground="#1F4E79").pack(pady=10)

        cols = ("tienda", "precio_usd", "costo_gtq", "precio_venta_gtq", "ahorro", "link")
        tabla_comp = ttk.Treeview(top, columns=cols, show="headings", height=8)
        tabla_comp.heading("tienda", text="Tienda / Proveedor")
        tabla_comp.heading("precio_usd", text="Precio ($USD)")
        tabla_comp.heading("costo_gtq", text="Costo Total (Q GTQ)")
        tabla_comp.heading("precio_venta_gtq", text="P. Venta Sugerido (Q)")
        tabla_comp.heading("ahorro", text="Evaluación / Estado")
        tabla_comp.heading("link", text="URL Enlace")

        tabla_comp.column("tienda", width=150)
        tabla_comp.column("precio_usd", width=90, anchor="e")
        tabla_comp.column("costo_gtq", width=120, anchor="e")
        tabla_comp.column("precio_venta_gtq", width=130, anchor="e")
        tabla_comp.column("ahorro", width=150, anchor="center")
        tabla_comp.column("link", width=100, anchor="center")

        tabla_comp.pack(fill="both", expand=True, padx=15, pady=5)

        peso_g = prod_encontrado.get("peso_g", 0.0)

        opciones_calculadas = []
        for item in enlaces:
            n_tienda = item.get("nombre_tienda", "")
            p_usd = item.get("precio_usd", 0.0)
            c_tot, p_sug = self.calcular_costo_y_precio_por_tienda(n_tienda, p_usd, peso_g)
            opciones_calculadas.append((item, c_tot, p_sug))

        opcion_barata = self.obtener_opcion_mas_economica(enlaces, peso_g)

        for item, c_tot, p_sug in opciones_calculadas:
            p_usd = item.get("precio_usd", 0.0)
            if "Ako" in item.get("nombre_tienda", ""):
                evaluacion = "📦 Opción Mayorista"
            elif opcion_barata and item == opcion_barata:
                evaluacion = "🟢 MÁS ECONÓMICA"
            else:
                evaluacion = "🔴 Opción Costosa"

            str_link = "🔗 Enlace" if item.get("url") else "Sin Link"

            tabla_comp.insert(
                "",
                tk.END,
                values=(
                    item.get("nombre_tienda"),
                    f"${p_usd:,.2f}",
                    f"Q{c_tot:,.2f}",
                    f"Q{p_sug:,.2f}",
                    evaluacion,
                    str_link,
                ),
            )

        def abrir_link_desde_tabla():
            sel = tabla_comp.selection()
            if not sel:
                return
            idx = tabla_comp.index(sel[0])
            url = enlaces[idx].get("url", "")
            if url:
                if not (url.startswith("http://") or url.startswith("https://")):
                    url = "https://" + url
                webbrowser.open(url)
            else:
                messagebox.showinfo("Sin Enlace", "Esta tienda no tiene registrada una URL.")

        ttk.Button(top, text="🛒 Abrir Enlace de la Tienda Seleccionada", command=abrir_link_desde_tabla).pack(pady=10)

    def mostrar_imagen_ampliada(self, ruta_o_url, titulo="Vista Previa de Producto"):
        if not ruta_o_url:
            messagebox.showinfo("Sin Imagen", "Este producto no tiene una imagen válida asociada.")
            return

        img_pil = descargar_imagen(ruta_o_url)
        if not img_pil:
            messagebox.showinfo("Error", "No se pudo descargar o abrir la imagen.")
            return

        top = tk.Toplevel(self.root)
        top.title(titulo)
        top.geometry("700x750")
        top.resizable(True, True)
        top.grab_set()

        try:
            img_copy = img_pil.copy()
            img_copy.thumbnail((650, 650), Image.Resampling.LANCZOS)
            photo_grande = ImageTk.PhotoImage(img_copy)

            lbl_img = ttk.Label(top, image=photo_grande)
            lbl_img.image = photo_grande
            lbl_img.pack(padx=15, pady=15, expand=True)

            btn_cerrar = ttk.Button(top, text="❌ Cerrar Vista Previa", command=top.destroy)
            btn_cerrar.pack(pady=10)

        except Exception as e:
            top.destroy()
            messagebox.showerror("Error", f"No se pudo cargar la imagen ampliada.\nDetalle: {e}")

    def ver_imagen_seleccionada(self):
        seleccion = self.tabla_cat.selection()
        if not seleccion:
            messagebox.showwarning("Atención", "Selecciona un producto de la tabla para ver su imagen ampliada.")
            return

        item_vals = self.tabla_cat.item(seleccion[0], "values")
        num_sel = int(item_vals[0].replace("#", ""))
        tipo_sel = item_vals[1]

        prod_encontrado = None
        for p in self.productos_catalogo:
            linea_p = p.get("linea", p.get("tienda", "Pandora"))
            if linea_p == self.linea_activa and p.get("num") == num_sel and p.get("tipo") == tipo_sel:
                prod_encontrado = p
                break

        if prod_encontrado:
            img_path = prod_encontrado.get("imagen", "")
            self.mostrar_imagen_ampliada(img_path, titulo=f"Producto Código {num_sel} - {tipo_sel} ({self.linea_activa})")

    def abrir_link_compra(self):
        seleccion = self.tabla_cat.selection()
        if not seleccion:
            messagebox.showwarning("Atención", "Selecciona un producto de la tabla para abrir su link de compra.")
            return

        item_vals = self.tabla_cat.item(seleccion[0], "values")
        num_sel = int(item_vals[0].replace("#", ""))
        tipo_sel = item_vals[1]

        prod_encontrado = None
        for p in self.productos_catalogo:
            linea_p = p.get("linea", p.get("tienda", "Pandora"))
            if linea_p == self.linea_activa and p.get("num") == num_sel and p.get("tipo") == tipo_sel:
                prod_encontrado = p
                break

        if prod_encontrado:
            url = prod_encontrado.get("link_compra", "").strip()
            if url:
                if not (url.startswith("http://") or url.startswith("https://")):
                    url = "https://" + url
                webbrowser.open(url)
            else:
                messagebox.showinfo("Sin Enlace", f"El producto Código {num_sel} no tiene un link de compra registrado.")

    def al_cambiar_linea(self, event=None):
        nueva_linea = self.combo_linea_activa.get()
        self.linea_activa = nueva_linea

        cats_linea = CATEGORIAS_POR_LINEA.get(nueva_linea, ["General"])
        self.entry_rap_tipo.config(values=cats_linea)
        self.entry_rap_tipo.set(cats_linea[0])

        self.combo_filtro_cat.config(values=["Todos"] + cats_linea)
        self.combo_filtro_cat.set("Todos")

        self.frame_tabla_cat.config(text=f" 3. Catálogo Exclusivo - Línea: {self.linea_activa} ")

        self.cancelar_edicion()
        self.renderizar_tabla_catalogo()

    def calcular_costo_envio_mexico(self, peso_gramos):
        if peso_gramos <= 0:
            return 0.0
        peso_libras = peso_gramos / 453.592
        if peso_libras <= 1.0:
            return 35.0
        costo = 35.0
        peso_restante = peso_libras - 1.0
        bloques_media_libra = math.ceil(peso_restante / 0.5)
        costo += bloques_media_libra * 17.50
        return costo

    def calcular_precio_venta(self, costo_unitario_gtq):
        if costo_unitario_gtq <= 0:
            return 0.0
        base = (costo_unitario_gtq * 2) + 10.0
        precio_redondeado = math.ceil(base / 5.0) * 5
        return float(precio_redondeado)

    def calcular_cotizacion_rapida(self):
        try:
            precio_usd = float(self.entry_rapido_precio.get()) if self.entry_rapido_precio.get() else 0.0
            peso_g = float(self.entry_rapido_peso.get()) if self.entry_rapido_peso.get() else 0.0
        except ValueError:
            return 0.0, 0.0

        tienda_sel = self.combo_tienda_link.get().strip() or "Friend"
        costo_total_gtq, precio_sugerido_gtq = self.calcular_costo_y_precio_por_tienda(tienda_sel, precio_usd, peso_g)

        base_gtq = precio_usd * self.tipo_cambio

        if "Friend" in tienda_sel:
            self.lbl_rap_base_gtq.config(text=f"Costo Base: Q{base_gtq:,.2f}")
            self.lbl_rap_comision.config(text="Comisión Tarjeta: Q0.00 (Exento)")
            self.lbl_rap_envio_usd.config(text="Flete USD/China: Q0.00 (Gratis)")
            self.lbl_rap_envio_mx.config(text="Flete México: Q1.75")
        elif "Ako" in tienda_sel:
            costo_prod_gtq = (precio_usd * 1.05) * self.tipo_cambio
            flete_cn = peso_g * 0.30
            flete_mx = peso_g * 0.10
            self.lbl_rap_base_gtq.config(text=f"Costo Base (+5%): Q{costo_prod_gtq:,.2f}")
            self.lbl_rap_comision.config(text="Comisión Tarjeta: Q0.00 (Exento)")
            self.lbl_rap_envio_usd.config(text=f"Flete China (Q0.30/g): Q{flete_cn:,.2f}")
            self.lbl_rap_envio_mx.config(text=f"Flete México (Q0.10/g): Q{flete_mx:,.2f}")
        else:
            pct_comision = self.params_prorrateo.get("pct_comision", 0.0)
            costo_flete_g = self.params_prorrateo.get("costo_envio_usd_por_g", 0.0)
            costo_flete_mx_g = self.params_prorrateo.get("costo_envio_mx_por_g", 0.0)

            comision_gtq = (precio_usd * pct_comision) * self.tipo_cambio
            envio_usd_gtq = (peso_g * costo_flete_g) * self.tipo_cambio
            envio_mx_gtq = peso_g * costo_flete_mx_g if costo_flete_mx_g > 0 else self.calcular_costo_envio_mexico(peso_g)

            self.lbl_rap_base_gtq.config(text=f"Costo Base: Q{base_gtq:,.2f}")
            self.lbl_rap_comision.config(text=f"Comisión Tarjeta: Q{comision_gtq:,.2f}")
            self.lbl_rap_envio_usd.config(text=f"Flete USD: Q{envio_usd_gtq:,.2f}")
            self.lbl_rap_envio_mx.config(text=f"Flete México: Q{envio_mx_gtq:,.2f}")

        self.lbl_rap_costo_total.config(text=f"COSTO TOTAL ({tienda_sel}): Q{costo_total_gtq:,.2f} GTQ")
        self.lbl_rap_precio_venta.config(text=f"PRECIO SUGERIDO: Q{precio_sugerido_gtq:,.2f} GTQ")

        return costo_total_gtq, precio_sugerido_gtq

    def cargar_producto_para_editar(self):
        seleccion = self.tabla_cat.selection()
        if not seleccion:
            messagebox.showwarning("Atención", "Selecciona primero un producto de la tabla para editar.")
            return

        item_vals = self.tabla_cat.item(seleccion[0], "values")
        num_sel = int(item_vals[0].replace("#", ""))
        tipo_sel = item_vals[1]

        prod_encontrado = None
        for p in self.productos_catalogo:
            linea_p = p.get("linea", p.get("tienda", "Pandora"))
            if linea_p == self.linea_activa and p.get("num") == num_sel and p.get("tipo") == tipo_sel:
                prod_encontrado = p
                break

        if not prod_encontrado:
            return

        self.entry_rap_tipo.set(prod_encontrado.get("tipo", CATEGORIAS_POR_LINEA[self.linea_activa][0]))
        self.entry_rap_num.delete(0, tk.END)
        self.entry_rap_num.insert(0, str(prod_encontrado.get("num", "")))

        tienda_act = prod_encontrado.get("tienda_origen", "Friend")
        self.combo_tienda_link.set(tienda_act)

        self.entry_rapido_precio.delete(0, tk.END)
        self.entry_rapido_precio.insert(0, str(prod_encontrado.get("precio_usd", "")))

        self.entry_rapido_peso.delete(0, tk.END)
        self.entry_rapido_peso.insert(0, str(prod_encontrado.get("peso_g", "")))

        self.enlaces_temp = list(prod_encontrado.get("enlaces", []))
        if not self.enlaces_temp and prod_encontrado.get("link_compra"):
            self.enlaces_temp.append({
                "nombre_tienda": tienda_act,
                "precio_usd": prod_encontrado.get("precio_usd", 0.0),
                "url": prod_encontrado.get("link_compra", ""),
            })

        self.actualizar_listbox_enlaces()

        self.entry_precio_venta_custom.delete(0, tk.END)
        p_real = prod_encontrado.get("precio_final_gtq", prod_encontrado.get("precio_venta_gtq", 0.0))
        self.entry_precio_venta_custom.insert(0, f"{p_real:.2f}")

        self.ruta_imagen_temp = prod_encontrado.get("imagen", "")
        self.mostrar_vista_previa(self.ruta_imagen_temp)

        self.calcular_cotizacion_rapida()

        self.btn_guardar_cat.config(text=f"💾 Actualizar en Línea {self.linea_activa}")
        self.btn_cancelar_edit.grid(row=3, column=4, padx=5, pady=4)

    def cancelar_edicion(self):
        self.entry_rapido_precio.delete(0, tk.END)
        self.entry_rapido_peso.delete(0, tk.END)
        self.entry_precio_tienda_link.delete(0, tk.END)
        self.entry_url_tienda_link.delete(0, tk.END)
        self.enlaces_temp = []
        self.actualizar_listbox_enlaces()

        self.entry_precio_venta_custom.delete(0, tk.END)
        self.ruta_imagen_temp = ""
        self.mostrar_vista_previa("")
        self.btn_guardar_cat.config(text="➕ Guardar en Catálogo de Línea")
        self.btn_cancelar_edit.grid_remove()

    def guardar_producto_catalogo(self):
        try:
            tipo = self.entry_rap_tipo.get().strip() or "General"
            num_item = int(self.entry_rap_num.get().strip())
            precio_usd = float(self.entry_rapido_precio.get())
            peso_g = float(self.entry_rapido_peso.get())
            tienda_seleccionada = self.combo_tienda_link.get().strip() or "Friend"

            self.registrar_nuevo_proveedor_si_no_existe(tienda_seleccionada)

            existe_tienda_principal = any(x for x in self.enlaces_temp if x.get("nombre_tienda") == tienda_seleccionada)
            if not existe_tienda_principal:
                self.enlaces_temp.append({
                    "nombre_tienda": tienda_seleccionada,
                    "precio_usd": precio_usd,
                    "url": "",
                })

            opcion_barata = self.obtener_opcion_mas_economica(self.enlaces_temp, peso_g)

            costo_tot, precio_sug = self.calcular_cotizacion_rapida()

            val_custom = self.entry_precio_venta_custom.get().strip()
            if val_custom:
                precio_final = float(val_custom)
            else:
                precio_final = precio_sug

            ruta_imagen_definitiva = ""
            if self.ruta_imagen_temp:
                ruta_local = descargar_y_guardar_localmente(self.ruta_imagen_temp)
                ruta_imagen_definitiva = ruta_local if ruta_local else self.ruta_imagen_temp

            existe = False
            for p in self.productos_catalogo:
                linea_p = p.get("linea", p.get("tienda", "Pandora"))
                if linea_p == self.linea_activa and p.get("num") == num_item and p.get("tipo") == tipo:
                    p["linea"] = self.linea_activa
                    p["precio_usd"] = precio_usd
                    p["peso_g"] = peso_g
                    p["tienda_origen"] = tienda_seleccionada
                    p["tienda_mas_barata"] = opcion_barata.get("nombre_tienda", "") if opcion_barata else tienda_seleccionada
                    p["enlaces"] = list(self.enlaces_temp)
                    p["costo_unit_gtq"] = costo_tot
                    p["precio_sugerido_gtq"] = precio_sug
                    p["precio_final_gtq"] = precio_final
                    p["precio_venta_gtq"] = precio_final
                    if ruta_imagen_definitiva:
                        p["imagen"] = ruta_imagen_definitiva
                    existe = True
                    break

            if not existe:
                self.productos_catalogo.append({
                    "linea": self.linea_activa,
                    "num": num_item,
                    "tipo": tipo,
                    "nombre": f"{tipo} #{num_item}",
                    "precio_usd": precio_usd,
                    "peso_g": peso_g,
                    "tienda_origen": tienda_seleccionada,
                    "tienda_mas_barata": opcion_barata.get("nombre_tienda", "") if opcion_barata else tienda_seleccionada,
                    "enlaces": list(self.enlaces_temp),
                    "costo_unit_gtq": costo_tot,
                    "precio_sugerido_gtq": precio_sug,
                    "precio_final_gtq": precio_final,
                    "precio_venta_gtq": precio_final,
                    "imagen": ruta_imagen_definitiva,
                })

            self.guardar_catalogo_global()
            self.guardar_estado_autosave()

            self.combo_filtro_cat.set(tipo)
            self.renderizar_tabla_catalogo()

            self.cancelar_edicion()
            self.entry_rap_num.delete(0, tk.END)
            self.entry_rap_num.insert(0, str(num_item + 1))

            messagebox.showinfo("Éxito", f"Producto Código {num_item} ({tipo}) guardado en la Línea '{self.linea_activa}'.")

        except ValueError:
            messagebox.showerror("Error", "Revisa los valores numéricos ingresados.")

    def eliminar_prod_catalogo(self):
        seleccion = self.tabla_cat.selection()
        if not seleccion:
            return
        item_vals = self.tabla_cat.item(seleccion[0], "values")
        num_sel = int(item_vals[0].replace("#", ""))
        tipo_sel = item_vals[1]

        self.productos_catalogo = [
            p for p in self.productos_catalogo if not (p.get("linea", p.get("tienda", "Pandora")) == self.linea_activa and p.get("num") == num_sel and p.get("tipo") == tipo_sel)
        ]
        self.guardar_catalogo_global()
        self.guardar_estado_autosave()
        self.renderizar_tabla_catalogo()

    def renderizar_tabla_catalogo(self, filtro_num=None):
        for item in self.tabla_cat.get_children():
            self.tabla_cat.delete(item)

        self.imagenes_cache.clear()

        prods_linea = [p for p in self.productos_catalogo if p.get("linea", p.get("tienda", "Pandora")) == self.linea_activa]
        cat_filtro = getattr(self, "combo_filtro_cat", None)
        filtro_seleccionado = cat_filtro.get() if cat_filtro else "Todos"

        if filtro_seleccionado == "Todos":
            prods_filtrados = prods_linea
        else:
            prods_filtrados = [p for p in prods_linea if p.get("tipo") == filtro_seleccionado]

        if filtro_num:
            prods_filtrados = [p for p in prods_filtrados if filtro_num in str(p.get("num", ""))]

        prods_filtrados.sort(key=lambda x: x.get("num", 0), reverse=True)

        for idx, p in enumerate(prods_filtrados):
            img_path = p.get("imagen", "")
            photo_obj = None

            if img_path:
                img_pil = descargar_imagen(img_path)
                if img_pil:
                    try:
                        img_pil.thumbnail((45, 45))
                        photo_obj = ImageTk.PhotoImage(img_pil)
                        self.imagenes_cache[f"{p.get('num')}_{idx}"] = photo_obj
                    except Exception as e:
                        print(f"Error renderizando imagen #{p.get('num')}: {e}")
                        photo_obj = None

            p_final = p.get("precio_final_gtq", p.get("precio_venta_gtq", 0.0))
            p_sug = p.get("precio_sugerido_gtq", 0.0)
            costo_tot = p.get("costo_unit_gtq", 0.0)

            ganancia_gtq = p_final - costo_tot
            pct_ganancia = (ganancia_gtq / costo_tot * 100) if costo_tot > 0 else 0.0
            str_ganancia = f"Q{ganancia_gtq:,.2f} ({pct_ganancia:,.1f}%)"

            enlaces = p.get("enlaces", [])
            if enlaces:
                opcion_barata = self.obtener_opcion_mas_economica(enlaces, p.get("peso_g", 0.0))
                if opcion_barata:
                    str_tienda_barata = f"🟢 {opcion_barata.get('nombre_tienda')} (${opcion_barata.get('precio_usd'):,.2f})"
                else:
                    str_tienda_barata = f"🟢 {p.get('tienda_origen', 'Friend')} (${p.get('precio_usd'):,.2f})"
            else:
                str_tienda_barata = f"🟢 {p.get('tienda_origen', 'Friend')} (${p.get('precio_usd'):,.2f})"

            valores = (
                f"{p.get('num')}",
                p.get("tipo"),
                p.get("tienda_origen", "Friend"),
                f"${p.get('precio_usd'):,.2f}",
                f"{p.get('peso_g')}g",
                f"Q{costo_tot:,.2f}",
                f"Q{p_sug:,.2f}",
                f"Q{p_final:,.2f}",
                str_ganancia,
                str_tienda_barata,
            )

            if photo_obj:
                self.tabla_cat.insert("", tk.END, image=photo_obj, values=valores)
            else:
                self.tabla_cat.insert("", tk.END, text="📷 Sin Img", values=valores)

    def guardar_archivo_prorrateo(self):
        nombre_proj = self.entry_nombre_proyecto.get().strip() or "Prorrateo"
        nombre_limpio = re.sub(r'[\\\\/*?:"<>|]', "", nombre_proj)

        datos_guardar = {
            "nombre_proyecto": nombre_proj,
            "pago_tarjeta_usd": self.entry_pago_tarjeta.get(),
            "costo_envio_usd": self.entry_costo_envio.get(),
            "total_piezas": self.entry_total_piezas.get(),
            "params_prorrateo": self.params_prorrateo,
            "productos_lote": self.productos_lote,
            "productos_catalogo": self.productos_catalogo,
        }

        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialfile=f"{nombre_limpio}.json",
            filetypes=[("Archivos JSON de Prorrateo", "*.json")],
            title="Guardar Archivo de Prorrateo Como...",
        )
        if not filepath:
            return

        try:
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(datos_guardar, f, ensure_ascii=False, indent=4)

            messagebox.showinfo("Éxito", f"El prorrateo '{nombre_proj}' se guardó correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo.\nDetalle: {e}")

    def cargar_archivo_prorrateo(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Archivos JSON de Prorrateo", "*.json")],
            title="Seleccionar Archivo de Prorrateo Guardado",
        )
        if not filepath:
            return

        try:
            with open(filepath, "r", encoding="utf-8") as f:
                datos = json.load(f)

            self.nombre_proyecto = datos.get("nombre_proyecto", "Prorrateo Cargado")
            self.entry_nombre_proyecto.delete(0, tk.END)
            self.entry_nombre_proyecto.insert(0, self.nombre_proyecto)

            self.entry_pago_tarjeta.delete(0, tk.END)
            self.entry_pago_tarjeta.insert(0, str(datos.get("pago_tarjeta_usd", "")))

            self.entry_costo_envio.delete(0, tk.END)
            self.entry_costo_envio.insert(0, str(datos.get("costo_envio_usd", "")))

            self.entry_total_piezas.delete(0, tk.END)
            self.entry_total_piezas.insert(0, str(datos.get("total_piezas", "")))

            prods = datos.get("productos_lote", datos.get("productos", []))
            self.productos_lote = []

            for p in prods:
                nombre = p.get("nombre", "Sin Nombre")
                cant = int(p.get("cant", 1))
                precio = float(p.get("precio_unit_usd", p.get("precio", 0.0)))
                peso_g = float(p.get("peso_unit_g", p.get("peso_unit", p.get("peso", 0.0))))

                self.productos_lote.append({
                    "nombre": nombre,
                    "cant": cant,
                    "precio": precio,
                    "subtotal": cant * precio,
                    "peso": peso_g,
                    "peso_tot": cant * peso_g,
                })

            if not self.productos_catalogo and "productos_catalogo" in datos:
                self.productos_catalogo = datos.get("productos_catalogo", [])
                for p in self.productos_catalogo:
                    if "linea" not in p:
                        p["linea"] = p.get("tienda", "Pandora")

            self.recalcular_lote_y_parametros()
            messagebox.showinfo("Éxito", f"Prorrateo '{self.nombre_proyecto}' cargado correctamente.")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo.\nDetalle: {e}")

    def nuevo_prorrateo(self):
        self.nombre_proyecto = "Prorrateo_Nuevo"
        self.entry_nombre_proyecto.delete(0, tk.END)
        self.entry_nombre_proyecto.insert(0, self.nombre_proyecto)

        self.entry_pago_tarjeta.delete(0, tk.END)
        self.entry_costo_envio.delete(0, tk.END)
        self.entry_total_piezas.delete(0, tk.END)

        self.productos_lote = []
        self.params_prorrateo = {
            "pct_comision": 0.0,
            "costo_envio_usd_por_g": 0.0,
            "costo_envio_mx_por_g": 0.0,
            "val_total_prod_usd": 0.0,
            "peso_total_g": 0.0,
        }

        self.recalcular_lote_y_parametros()

    def guardar_catalogo_global(self):
        try:
            with open(CATALOGO_GLOBAL_FILE, "w", encoding="utf-8") as f:
                json.dump(self.productos_catalogo, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Error al guardar catálogo global: {e}")

    def cargar_catalogo_global(self):
        if os.path.exists(CATALOGO_GLOBAL_FILE):
            try:
                with open(CATALOGO_GLOBAL_FILE, "r", encoding="utf-8") as f:
                    self.productos_catalogo = json.load(f)

                for p in self.productos_catalogo:
                    if "linea" not in p:
                        p["linea"] = p.get("tienda", "Pandora")

                self.renderizar_tabla_catalogo()
            except Exception as e:
                print(f"Error al cargar catálogo global: {e}")

    def guardar_estado_autosave(self):
        try:
            with open(CONFIG_AUTOSAVE, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "nombre_proyecto": self.entry_nombre_proyecto.get().strip(),
                        "pago_tarjeta_usd": self.entry_pago_tarjeta.get(),
                        "costo_envio_usd": self.entry_costo_envio.get(),
                        "total_piezas": self.entry_total_piezas.get(),
                        "params_prorrateo": self.params_prorrateo,
                        "productos_lote": self.productos_lote,
                        "productos_catalogo": self.productos_catalogo,
                    },
                    f,
                    ensure_ascii=False,
                    indent=4,
                )
        except Exception:
            pass

    def cargar_ultimo_estado_automatico(self):
        if os.path.exists(CONFIG_AUTOSAVE):
            try:
                with open(CONFIG_AUTOSAVE, "r", encoding="utf-8") as f:
                    datos = json.load(f)

                self.nombre_proyecto = datos.get("nombre_proyecto", "Prorrateo_Nuevo")
                self.entry_nombre_proyecto.delete(0, tk.END)
                self.entry_nombre_proyecto.insert(0, self.nombre_proyecto)

                self.entry_pago_tarjeta.insert(0, datos.get("pago_tarjeta_usd", ""))
                self.entry_costo_envio.insert(0, datos.get("costo_envio_usd", ""))
                self.entry_total_piezas.insert(0, datos.get("total_piezas", ""))

                self.params_prorrateo = datos.get("params_prorrateo", self.params_prorrateo)
                self.productos_lote = datos.get("productos_lote", [])

                if not self.productos_catalogo and "productos_catalogo" in datos:
                    self.productos_catalogo = datos.get("productos_catalogo", [])
                    for p in self.productos_catalogo:
                        if "linea" not in p:
                            p["linea"] = p.get("tienda", "Pandora")

                self.recalcular_lote_y_parametros()
            except Exception:
                pass

    def generar_catalogo_pdf(self):
        prods_linea = [p for p in self.productos_catalogo if p.get("linea", p.get("tienda", "Pandora")) == self.linea_activa]

        if not prods_linea:
            messagebox.showwarning("Atención", f"No hay productos archivados para la línea '{self.linea_activa}'.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile=f"Catalogo_Linea_{self.linea_activa}_{self.entry_nombre_proyecto.get().strip()}.pdf",
            filetypes=[("Documento PDF", "*.pdf")],
            title=f"Guardar Catálogo PDF (Línea {self.linea_activa}) Como...",
        )
        if not filepath:
            return

        try:
            doc = SimpleDocTemplate(filepath, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
            story = []
            styles = getSampleStyleSheet()

            style_titulo = ParagraphStyle("TituloCat", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=22, textColor=HexColor("#1F4E79"), alignment=1, spaceAfter=5)
            style_subtitulo = ParagraphStyle("SubTituloCat", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=14, textColor=HexColor("#555555"), alignment=1, spaceAfter=15)
            
            style_indice_head = ParagraphStyle("IndiceHead", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=14, textColor=HexColor("#1F4E79"), spaceBefore=10, spaceAfter=6)
            style_indice_link = ParagraphStyle("IndiceLink", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11, textColor=HexColor("#0056b3"), spaceAfter=4)
            
            style_categoria = ParagraphStyle("CatSeccion", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=15, textColor=HexColor("#C00000"), spaceBefore=15, spaceAfter=8)
            style_prod_nombre = ParagraphStyle("ProdNombre", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11, textColor=HexColor("#2C3E50"), spaceAfter=4)
            style_precio = ParagraphStyle("ProdPrecio", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=13, textColor=HexColor("#008000"))

            # --- ENCABEZADO ---
            story.append(Paragraph(f"CATÁLOGO DE PRODUCTOS", style_titulo))
            story.append(Paragraph(f"LÍNEA: {self.linea_activa.upper()}", style_subtitulo))
            story.append(Spacer(1, 10))

            categorias = list(set(p.get("tipo", "General") for p in prods_linea))
            categorias.sort()

            # --- ÍNDICE CON HIPERVÍNCULOS DE NAVEGACIÓN ---
            story.append(Paragraph("📌 ÍNDICE DE CATEGORÍAS", style_indice_head))
            for cat in categorias:
                tag_cat = re.sub(r'[^a-zA-Z0-9]', '', cat)
                story.append(Paragraph(f'• <a href="#{tag_cat}" color="#0056b3"><u>{cat.upper()}</u></a>', style_indice_link))
            
            story.append(Spacer(1, 15))

            ocultar_precios = self.var_ocultar_precios_pdf.get()

            # --- SECCIONES DEL CATÁLOGO ---
            for cat in categorias:
                tag_cat = re.sub(r'[^a-zA-Z0-9]', '', cat)
                
                # Ancla para el hipervínculo del índice
                story.append(Paragraph(f'<a name="{tag_cat}"/>• CATEGORÍA: {cat.upper()}', style_categoria))
                
                prods_cat = [p for p in prods_linea if p.get("tipo", "General") == cat]
                prods_cat.sort(key=lambda x: x.get("num", 0), reverse=True)

                filas_tabla = []
                row_actual = []

                for p in prods_cat:
                    p_num = p.get("num", 0)
                    p_precio = p.get("precio_final_gtq", p.get("precio_venta_gtq", 0.0))
                    img_path = p.get("imagen", "")

                    img_obj = Paragraph("<b>[Sin Imagen]</b>", styles["Normal"])
                    if img_path:
                        img_pil = descargar_imagen(img_path)
                        if img_pil:
                            try:
                                temp_img_file = os.path.join(CARPETA_CACHE_IMG, f"temp_pdf_{p_num}.png")
                                img_pil.save(temp_img_file)
                                img_obj = RLImage(temp_img_file, width=100, height=100)
                            except Exception:
                                img_obj = Paragraph("<b>[Sin Imagen]</b>", styles["Normal"])

                    contenido_celda = [
                        img_obj,
                        Spacer(1, 4),
                        Paragraph(f"<b>Código {p_num}</b>", style_prod_nombre),
                    ]

                    # Si NO está marcada la casilla de ocultar, se añade el precio
                    if not ocultar_precios:
                        contenido_celda.append(Paragraph(f"Precio: Q{p_precio:,.2f}", style_precio))

                    row_actual.append(contenido_celda)

                    if len(row_actual) == 2:
                        filas_tabla.append(row_actual)
                        row_actual = []

                if row_actual:
                    row_actual.append("")
                    filas_tabla.append(row_actual)

                if filas_tabla:
                    tabla_pdf = Table(filas_tabla, colWidths=[260, 260])
                    tabla_pdf.setStyle(
                        TableStyle([
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("INNERGRID", (0, 0), (-1, -1), 0.5, HexColor("#D3D3D3")),
                            ("BOX", (0, 0), (-1, -1), 1, HexColor("#1F4E79")),
                            ("TOPPADDING", (0, 0), (-1, -1), 8),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                        ])
                    )
                    story.append(tabla_pdf)
                    story.append(Spacer(1, 15))

            doc.build(story)
            messagebox.showinfo("Éxito", f"El catálogo PDF de la línea '{self.linea_activa}' se generó correctamente.")

        except Exception as e:
            messagebox.showerror("Error al generar PDF", f"Detalle del error:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = CotizadorApp(root)
    root.mainloop()