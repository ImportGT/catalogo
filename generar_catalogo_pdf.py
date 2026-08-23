import os
import pandas as pd
from PIL import Image
from fpdf import FPDF

class PDFCatalogo(FPDF):
    def header(self):
        self.set_font('helvetica', 'B', 14)
        self.set_text_color(31, 78, 121)
        self.cell(0, 10, 'CATÁLOGO DE PRODUCTOS - ANILLOS PANDORA', new_x="LMARGIN", new_y="NEXT", align='C')
        self.ln(5)

def generar_pdf_desde_excel(ruta_excel, nombre_linea, mostrar_precios=True, margen_porcentaje=0):
    carpeta_salida = "catalogos"
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)

    nombre_archivo = os.path.join(carpeta_salida, f"Catalogo_{nombre_linea.replace(' ', '_')}.pdf")

    try:
        df = pd.read_excel(ruta_excel)
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return

    pdf = PDFCatalogo(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    
    startX = 14
    startY = 20
    cols = 4
    cardWidth = 42
    cardHeight = 52 if mostrar_precios else 44
    gapX = 5
    gapY = 6
    
    currentX = startX
    currentY = startY
    countInRow = 0

    if df.empty:
        print(f"Aviso: El archivo Excel de {nombre_linea} no contiene registros.")
        return

    for index, row in df.iterrows():
        raw_id = row.get('NUMERO', '')
        
        # Eliminar decimales convirtiendo a entero (ej: 565.0 -> 565)
        try:
            prod_id = int(float(raw_id))
        except (ValueError, TypeError):
            prod_id = raw_id

        precio_base = row.get('PRECIO', 0)
        if pd.isna(precio_base):
            precio_base = 0.0

        precio_final = precio_base * (1 + (margen_porcentaje / 100))

        # --- BUSCAR LA IMAGEN DE PORTADA (.1.webp o .1.jpg) ---
        imagen_path = ""
        posibles_rutas = [
            f"imagenes/anillos/anillos_{prod_id}.1.webp",
            f"imagenes/anillos/anillos_{prod_id}.1.jpg",
            f"imagenes/anillos/anillos_{prod_id}.webp",
            f"imagenes/anillos/anillos_{prod_id}.jpg"
        ]

        for ruta in posibles_rutas:
            if os.path.exists(ruta):
                imagen_path = ruta
                break

        if currentY + cardHeight > 280:
            pdf.add_page()
            currentY = 20
            currentX = startX
            countInRow = 0

        # Dibujar tarjeta de producto
        pdf.set_draw_color(220, 220, 220)
        pdf.set_fill_color(255, 255, 255)
        pdf.rect(currentX, currentY, cardWidth, cardHeight, 'DF')

        # Insertar imagen de portada si existe
        if imagen_path and os.path.exists(str(imagen_path)):
            try:
                pdf.image(str(imagen_path), x=currentX + 6, y=currentY + 4, w=30, h=30)
            except Exception as ex:
                print(f"No se pudo cargar la imagen para el ID {prod_id}: {ex}")

        # Código del producto
        pdf.set_font('helvetica', 'B', 10)
        pdf.set_text_color(44, 62, 80)
        pdf.set_xy(currentX, currentY + 35)
        pdf.cell(cardWidth, 6, f"Cod. {prod_id}", 0, 0, 'C')

        # Precio
        if mostrar_precios:
            pdf.set_font('helvetica', 'B', 10)
            pdf.set_text_color(0, 128, 0)
            pdf.set_xy(currentX, currentY + 42)
            pdf.cell(cardWidth, 6, f"Q{precio_final:.2f}", 0, 0, 'C')

        countInRow += 1
        if countInRow >= cols:
            countInRow = 0
            currentX = startX
            currentY += cardHeight + gapY
        else:
            currentX += cardWidth + gapX

    pdf.output(nombre_archivo)
    print(f"¡Catálogo completo generado con éxito: {nombre_archivo}!")

if __name__ == "__main__":
    generar_pdf_desde_excel("ANILLOS.xlsx", "Pandora_Anillos", mostrar_precios=True, margen_porcentaje=0)