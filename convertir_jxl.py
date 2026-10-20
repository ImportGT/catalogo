import os
import subprocess

def convertir_y_reemplazar_jxl():
    print("Iniciando la conversión masiva de .jxl a .jpg en todas las subcarpetas...")
    contador = 0
    
    # os.walk recorre automáticamente la carpeta actual y todas las subcarpetas hacia abajo
    for raiz, directorios, archivos in os.walk("."):
        for archivo in archivos:
            if archivo.lower().endswith(".jxl"):
                ruta_jxl = os.path.join(raiz, archivo)
                ruta_jpg = os.path.splitext(ruta_jxl)[0] + ".jpg"
                
                convertido_exitosamente = False
                
                # Si el .jpg ya existe, damos por sentado que está listo
                if os.path.exists(ruta_jpg):
                    convertido_exitosamente = True
                else:
                    try:
                        # Intenta convertir usando djxl.exe ubicado en la carpeta principal
                        subprocess.run(["djxl.exe", ruta_jxl, ruta_jpg], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        if os.path.exists(ruta_jpg):
                            convertido_exitosamente = True
                            print(f"[CONVERTIDO] En '{raiz}': {archivo} -> .jpg")
                    except Exception as e:
                        print(f"[ERROR] No se pudo convertir {archivo} en {raiz}: {e}")
                
                # Si se convirtió o ya existía el jpg, eliminamos el .jxl original para evitar duplicados
                if convertido_exitosamente:
                    try:
                        os.remove(ruta_jxl)
                        print(f"[ELIMINADO] Original .jxl borrado: {archivo}")
                        contador += 1
                    except Exception as e:
                        print(f"[ADVERTENCIA] No se pudo eliminar {archivo}: {e}")

    print(f"\n¡Proceso finalizado! Se limpiaron y actualizaron {contador} archivos en total.")

if __name__ == "__main__":
    convertir_y_reemplazar_jxl()
    input("\nPresiona Enter para cerrar esta ventana...")