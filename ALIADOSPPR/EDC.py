import os
import sys
import csv
import tkinter as tk
from tkinter import filedialog, Tk, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm
import math
from reportlab.pdfbase.pdfmetrics import stringWidth
from PyPDF2 import PdfMerger
from datetime import datetime

def resource_path(relative_path):
    """Obtiene la ruta del recurso desde el ejecutable."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def leer_archivo(archivo):
    """Lee un archivo txt delimitado por '|', intentando primero con UTF-8 y luego con latin1."""
    try:
        with open(archivo, 'r', encoding='utf-8') as f:
            lineas = [line.strip().split('|') for line in f]
    except UnicodeDecodeError:
        with open(archivo, 'r', encoding='latin1') as f:
            lineas = [line.strip().split('|') for line in f]
    
    return lineas

def generar_pdfs(datos, nombre_pdf, total_hojas, fondo_portada=None, fondo_overflow=None, fondo_final=None):
    """Genera un PDF con los datos y un fondo opcional."""
    c = canvas.Canvas(nombre_pdf, pagesize=letter)
    width, height = letter
    
    #y_position = height - 50
    max_cols = max(len(fila) for fila in datos)
    contcargos = 0
    contreg = 0
    periodo = ""
    fecha = ""
    poliza = ""
    final = 0
    conthojas = 1
    primerahoja = 0
    min_font_size = 7
    font_size = 9
    max_width = 102
    for fila in datos:
        fila.extend([''] * (max_cols - len(fila)))
        
        if(primerahoja == 0):
            if fila[0] == "01":
                if fondo_portada:
                    fondo_img = ImageReader(fondo_portada)
                c.drawImage(fondo_img, 0, 0, width, height)
                c.setFont("Arial", 10) 
                c.drawRightString(20.5 * cm, 24.19* cm, "PERIODO: " + str(fila[2])) #Periodo
                c.drawString(18.75 * cm, 25.25* cm, "Hoja " + str(conthojas) + " de " + str(total_hojas[2])) #Hojas
                conthojas += 1
                periodo = str(fila[2])
            if fila[0] == "02":
                #Información general de la póliza
                c.setFont("Arial", 10) 
                c.drawString(3.35 * cm, 22.4* cm, str(fila[2]).strip()) #Contratante
                c.drawString(3.00 * cm, 21.65* cm, str(fila[3]).split(",")[0].strip()) #Calle
                c.drawString(3.00 * cm, 21.15* cm, "Col." + str(fila[3]).split(",")[1].strip() + ", " +  str(fila[3]).split(",")[2].strip()) #CP
                c.drawString(3.00 * cm, 20.65* cm, str(fila[3]).split(",")[3].strip()) #Ciudad y Estado
                c.drawString(3.35 * cm, 20.13* cm, str(fila[4]).strip()) #Asegurado
                c.drawString(16.5 * cm, 22.4* cm, str(fila[5]).strip()) #Poliza
                c.drawString(16.5 * cm, 21.95* cm, str(fila[6]).strip()) #Forma de pago
                c.drawString(16.5 * cm, 21.05* cm, "$ " + str(fila[8]).strip()) #Prima anual
                c.drawString(16.5 * cm, 20.6* cm, str(fila[9]).strip()) #Moneda

                while font_size >= min_font_size:
                    text_width = stringWidth(str(fila[7]).strip(), "Arial", font_size)
                    if text_width <= max_width:
                        break
                    font_size -= 0.25  # Disminuye poco a poco

                c.setFont("Arial", font_size)
                c.drawString(16.5 * cm, 21.5* cm, str(fila[7]).strip()) #Vigencia

                c.drawString(16.5 * cm, 20.15* cm, str(fila[10]).strip()) #Pagado hasta
            if fila[0] == "03":
                #Promotor
                c.setFont("Arial", 10) 
                c.drawString(3.35 * cm, 19.425* cm, str(fila[2]).strip()) #Agente  
                c.drawString(3.35 * cm, 18.7* cm, str(fila[3]).strip()) #Promotor      
            if fila[0] == "04":
                #Resumen general del fondo PPR
                c.setFont("Arial", 10) 
                if contcargos == 0:
                    #Saldo Anterior
                    c.drawString(6.5 * cm,  16.55* cm, "$ " + str(fila[2]).strip()) #Cargo
                    c.drawString(11.5 * cm, 16.55* cm, "$ " + str(fila[3]).strip()) #Abono 
                    c.drawString(16.5 * cm, 16.55* cm, "$ " + str(fila[4]).strip()) #Saldo 
                    contcargos += 1
                elif contcargos == 1:
                    #Aportaciones Adicionales
                    c.drawString(6.5 * cm,  15.75* cm, "$ " + str(fila[2]).strip()) #Cargo
                    c.drawString(11.5 * cm, 15.75* cm, "$ " + str(fila[3]).strip()) #Abono 
                    c.drawString(16.5 * cm, 15.75* cm, "$ " + str(fila[4]).strip()) #Saldo 
                    contcargos += 1  
                elif contcargos == 2:
                    #Retiros
                    c.drawString(6.5 * cm,  14.95* cm, "$ " + str(fila[2]).strip()) #Cargo
                    c.drawString(11.5 * cm, 14.95* cm, "$ " + str(fila[3]).strip()) #Abono 
                    c.drawString(16.5 * cm, 14.95* cm, "$ " + str(fila[4]).strip()) #Saldo 
                    contcargos += 1            
                elif contcargos == 3:
                    #Cargos
                    c.drawString(6.5 * cm,  14.15* cm, "$ " + str(fila[2]).strip()) #Cargo
                    c.drawString(11.5 * cm, 14.15* cm, "$ " + str(fila[3]).strip()) #Abono 
                    c.drawString(16.5 * cm, 14.15* cm, "$ " + str(fila[4]).strip()) #Saldo 
                    contcargos += 1
                elif contcargos == 4:
                    #Rendimiento
                    c.drawString(6.5 * cm,  13.35* cm, "$ " + str(fila[2]).strip()) #Cargo
                    c.drawString(11.5 * cm, 13.35* cm, "$ " + str(fila[3]).strip()) #Abono 
                    c.drawString(16.5 * cm, 13.35* cm, "$ " + str(fila[4]).strip()) #Saldo 
                    contcargos += 1   
                elif contcargos == 5:
                    #Retencion de ISR
                    c.drawString(6.5 * cm,  12.55* cm, "$ " + str(fila[2]).strip()) #Cargo
                    c.drawString(11.5 * cm, 12.55* cm, "$ " + str(fila[3]).strip()) #Abono 
                    c.drawString(16.5 * cm, 12.55* cm, "$ " + str(fila[4]).strip()) #Saldo 
                    contcargos += 1    
                else:
                    #Saldo Actual
                    c.drawString(6.5 * cm,  11.75* cm, "$ " + str(fila[2]).strip()) #Cargo
                    c.drawString(11.5 * cm, 11.75* cm, "$ " + str(fila[3]).strip()) #Abono 
                    c.drawString(16.5 * cm, 11.75* cm, "$ " + str(fila[4]).strip()) #Saldo        
            if fila[0] == "05":
                c.setFont("Arial", 10) 
                c.drawString(1 * cm, .55* cm, str(fila[2])) #Fecha Inferior 
                fecha = str(fila[2]) 
                poliza = str(fila[1]) 
                c.showPage()
                primerahoja = 1

        else: 
            if total_hojas[1] > 31:
                if fondo_overflow:
                    fondo_img = ImageReader(fondo_overflow)
                if fila[0] == "06": 
                    c.drawImage(fondo_img, 0, 0, width, height)
                    c.setFont("Arial", 10) 
                    c.drawString(9.43 * cm, 24.19* cm, "PERIODO: " + str(fila[2])) #Periodo 
                if fila[0] == "07":
                    c.setFont("Arial", 10) 
                    c.drawString(17.8 * cm, 23.72 * cm, "Póliza: " + str(fila[2])) #Poliza
                    c.drawString(1 * cm, .55* cm, fecha) #Fecha Inferior
                if fila[0] == "08":
                #Detalles de movimientos
                    c.setFont("Arial", 10) 
                #Saldo Anterior
                    c.drawString( 1.2 * cm, (21.9 - (0.5*contreg)) * cm, str(fila[2]).strip()) #Fecha
                    c.drawString( 4.2 * cm, (21.9 - (0.5*contreg)) * cm, str(fila[3]).strip()) #Movimiento 
                    c.drawString(  13 * cm, (21.9 - (0.5*contreg)) * cm, "$ " + str(fila[4]).strip()) #Cargo
                    c.drawString(15.6 * cm, (21.9 - (0.5*contreg)) * cm, "$ " + str(fila[5]).strip()) #Abono 
                    c.drawString(18.1 * cm, (21.9 - (0.5*contreg)) * cm, "$ " + str(fila[6]).strip()) #Saldo 
                    contreg += 1
                if contreg == 31:
                    c.drawString(18.75 * cm, 25.25* cm, "Hoja " + str(conthojas) + " de " + str(total_hojas[2])) #Hojas
                    conthojas += 1
                    c.showPage()
                    total_hojas[1] = total_hojas[1] - 31
                    contreg = 0
            else:
                if fondo_final:
                    fondo_img = ImageReader(fondo_final) 
                if final == 0:   
                    c.drawImage(fondo_img, 0, 0, width, height)
                    final = 1
                if fila[0] == "08":
                #Detalles de movimientos
                    c.setFont("Arial", 10) 
                #Saldo Anterior
                    c.drawString( 0.9 * cm, (20.85 - (0.5*contreg)) * cm, str(fila[2]).strip()) #Fecha
                    c.drawString( 3.6 * cm, (20.85 - (0.5*contreg)) * cm, str(fila[3]).strip()) #Movimiento 
                    c.drawString(11.75 * cm, (20.85 - (0.5*contreg)) * cm, "$ " + str(fila[4]).strip()) #Cargo
                    c.drawString(  15 * cm, (20.85 - (0.5*contreg)) * cm, "$ " + str(fila[5]).strip()) #Abono 
                    c.drawString(18.3 * cm, (20.85 - (0.5*contreg)) * cm, "$ " + str(fila[6]).strip()) #Saldo 
                    contreg += 1
                
                if fila[0] == "09":
                    #Saldo actual
                    c.setFont("Arial", 10)
                    c.drawString( 18.3 * cm, 5.4 * cm, "$ " + str(fila[2]).strip()) #Saldo actual
                if fila[0] == "10":
                    #Valor de rescate
                    c.setFont("Arial", 10)
                    c.drawString(.5 * cm, 23.2* cm, "PERIODO: " + periodo) #Periodo 
                    c.drawString(17.8 * cm, 22.7 * cm, poliza) #Poliza
                    c.drawString( 18.3 * cm, 4.65 * cm, "$ " + str(fila[2]).strip()) #Saldo actual
                    c.drawString(.45 * cm, 1.15* cm, fecha) #Fecha Inferior
                    c.drawString(19.25 * cm, 24.5* cm, "Hoja " + str(conthojas) + " de " + str(total_hojas[2])) #Hojas
                    conthojas += 1
                    
                    c.showPage()
                    

    c.save()
    return conthojas

def procesar_txt_a_pdfs(archivo_txt, output_folder, fondo_portada=None, fondo_overflow=None, fondo_final=None):
    lineas = leer_archivo(archivo_txt)
    pdf_actual = 1
    grupo_actual = None
    datos_grupo = []
    total_hojas = []
    cuenta_reg = 0
    cuenta_poliza = 0
    cuenta_hojas = 0
    hoy = datetime.now()
    # Mes con dos dígitos
    mes = hoy.strftime("%m")
    # Año en cuatro dígitos
    anio = hoy.strftime("%Y")
    carpeta_salida = os.path.abspath(os.path.join(output_folder, ".."))
    archivo_salida = os.path.join(carpeta_salida, 'datos_extraidos.csv')

    with open(archivo_salida, mode='w', newline='', encoding='utf-8') as archivo_salida:
        escritor = csv.writer(archivo_salida)
        escritor.writerow(['CC0000', 'PRODUCTO', 'TIPO_EDC', 'POLIZA', 'PAGINAS', 'NOMBRE COMPLETO', 'DIRECCION', 'DIRECCION_2','ESTADO', 'CP', '', 'FOLIO' ])
        
        for fila in lineas:
            if grupo_actual is None:
                grupo_actual = fila[1]
                
            if fila[1] != grupo_actual:
                cuenta_hojas = generar_pdfs(datos_grupo, f"{output_folder}\AliadosPPR_0329_{datos_grupo[0][1]}_0_0_EC_{mes}{anio}.pdf", total_hojas, fondo_portada, fondo_overflow, fondo_final)
                print(f"Archivo creado: AliadosPPR_0329_{datos_grupo[0][1]}_0_0_EC_{mes}{anio}.pdf")
                escritor.writerow(['CC0000', '0329', 'Aliados_PPR', datos_grupo[0][1], cuenta_hojas-1, datos_grupo[1][2], 
                                   str(datos_grupo[1][3]).split(',')[0], str(datos_grupo[1][3]).split(',')[1] + str(datos_grupo[1][3]).split(',')[3], 
                                   str(datos_grupo[1][3]).split(',')[3], str(datos_grupo[1][3]).split(',')[2].split(' ')[2], '', f"{(cuenta_poliza+1):07d}"])
                pdf_actual += 1
                grupo_actual = fila[1]
                datos_grupo = []
                cuenta_reg = 0
                cuenta_poliza += 1
                
            if fila[0] == '08':
                cuenta_reg += 1

            total_hojas = [fila[1], cuenta_reg, math.ceil(cuenta_reg/31)+1]
            datos_grupo.append(fila)
            
        if datos_grupo:
            cuenta_hojas = generar_pdfs(datos_grupo, f"{output_folder}\AliadosPPR_0329_{datos_grupo[0][1]}_0_0_EC_{mes}{anio}.pdf", total_hojas, fondo_portada, fondo_overflow, fondo_final)
            print(f"Archivo creado: AliadosPPR_0329_{datos_grupo[0][1]}_0_0_EC_{mes}{anio}.pdf")
            escritor.writerow(['CC0000', '0329', 'Aliados_PPR', datos_grupo[0][1], cuenta_hojas-1, datos_grupo[1][2], 
                    str(datos_grupo[1][3]).split(',')[0], str(datos_grupo[1][3]).split(',')[1] + str(datos_grupo[1][3]).split(',')[3], 
                    str(datos_grupo[1][3]).split(',')[3], str(datos_grupo[1][3]).split(',')[2].split(' ')[2], '', f"{(cuenta_poliza+1):07d}"])

def seleccionar_carpeta():                                          #Definir funcion de seleccion de carpeta
    root = tk.Tk()                                                  #Define la raiz del programa
    root.withdraw()                     
    carpeta = filedialog.askdirectory()                             #define la carpeta como ventana de dialogo
    return carpeta

def combinar_pdfs_en_carpeta(destino):
    
    i=1
    if not destino:
        messagebox.showinfo("Cancelado", "No se seleccionó ninguna carpeta.")
        return

    # Buscar solo archivos PDF
    archivos_pdf = sorted([
        f for f in os.listdir(destino)
        if f.lower().endswith('.pdf')
    ])

    if not archivos_pdf:
        messagebox.showinfo("Sin PDFs", "No se encontraron archivos PDF en la carpeta seleccionada.")
        return

    # Crear merger
    merger = PdfMerger()

    for archivo in archivos_pdf:
        ruta_pdf = os.path.join(destino, archivo)
        print("Archivo compilado " + str(i) + " de " + str(len(archivos_pdf)) + ": " + archivo)
        i += 1
        merger.append(ruta_pdf)

    # Carpeta de salida (una carpeta arriba)
    carpeta_salida = os.path.abspath(os.path.join(destino, ".."))
    archivo_salida = os.path.join(carpeta_salida, "VB.pdf")

    # Guardar el PDF combinado
    merger.write(archivo_salida)
    merger.close()

    messagebox.showinfo("Éxito", f"PDF de muestras guardado en:\n{archivo_salida}")

def main():
    archivo_txt = filedialog.askopenfilename(title="Selecciona el archivo txt", filetypes=[("txt files", "*.txt")])
    output_folder = seleccionar_carpeta()
    fondo_portada = ".\Recursos\Estado de Cuenta A PPR Portada.png"
    fondo_overflow = ".\Recursos\Estado de Cuenta A PPR Overflow.png"
    fondo_final = ".\Recursos\Estado de Cuenta A PPR Final.png"

    # Registro de Fonts Arial y ArialBold
    arial_font_path = resource_path('Arial.ttf')
    arial_bold_font_path = resource_path('ArialBD.ttf')

    pdfmetrics.registerFont(TTFont('Arial', arial_font_path))
    pdfmetrics.registerFont(TTFont('ArialBD', arial_bold_font_path))
    procesar_txt_a_pdfs(archivo_txt, output_folder, fondo_portada, fondo_overflow, fondo_final)
    combinar_pdfs_en_carpeta(output_folder)

if __name__ == "__main__":
    main()
