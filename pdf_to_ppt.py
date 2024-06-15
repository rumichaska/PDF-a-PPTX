""" Importación de módulos
"""

import os
import fnmatch
import fitz
import cv2
import numpy as np

from pptx import Presentation
from pptx.util import Inches

# FUNCIONES ----

def pdf_to_pptx(pdf_path, filename, zoom_factor=2):
    """Covierte un archivo pdf en pptx"""

    # Patrón del nombre del pdf
    filename_pattern = filename[0:len(filename) - 4]

    # Crear directorio de imágenes de las diapositivas
    page_dir = os.path.exists("./pages/")
    if not page_dir:
        os.makedirs("./pages/")

    # Crear subdirectorio del pdf a convertir
    pdf_slide_dir = os.path.exists(f"./pages/{filename_pattern}")
    if not pdf_slide_dir:
        os.makedirs(f"./pages/{filename_pattern}")

    # Abre el archivo PDF
    pdf_document = fitz.open(f"{pdf_path}{filename}")

    # Iteración por páginas
    for page_number in range(len(pdf_document)):
        # Extrae la página como imagen con mayor resolución
        page = pdf_document.load_page(page_number)
        matrix = fitz.Matrix(zoom_factor, zoom_factor)
        pix = page.get_pixmap(matrix=matrix)

        # Guardar temporalmente la imágen de la página
        new_page_number = page_number + 1
        slide_number = f"0{new_page_number}" if new_page_number < 10 else f"{new_page_number}"
        img_path = f"./pages/{filename_pattern}/page_{slide_number}.png"
        pix.save(img_path)


def get_tables(png_path, out_path, filename):
    """Extrae tablas a partir de imágenes según un color"""

    # Crear directorio de imágenes de las diapositivas
    table_dir = os.path.exists(out_path)
    if not table_dir:
        os.makedirs(out_path)

    # Cargar imágen
    new_image = cv2.imread(f"{png_path}{filename}")

    # Convertir imágen a formato HSV para detectar color patrón
    hsv = cv2.cvtColor(new_image, cv2.COLOR_BGR2HSV)

    # Definir el rango de colores HSV patrón
    lower_purple = np.array([120, 40, 40])
    upper_purple = np.array([150, 255, 255])

    # Create a mask for purple color
    mask_purple = cv2.inRange(hsv, lower_purple, upper_purple)

    # Find contours in the mask
    contours_purple, _ = cv2.findContours(mask_purple, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # Filter out small contours that are not likely to be the purple border
    contours_purple = [cnt for cnt in contours_purple if cv2.contourArea(cnt) > 100]

    # If no purple contour is found, do not extract anything
    if contours_purple:
        # Assuming the largest purple contour is the desired table
        largest_contour_purple = max(contours_purple, key=cv2.contourArea)
        x, y, w, h = cv2.boundingRect(largest_contour_purple)

        # Adjust coordinates to exclude the purple border
        border_margin = 12  # Adjust this margin as needed to exclude the border
        x = max(0, x + border_margin)
        y = max(0, y + border_margin)
        w = max(0, w - 2 * border_margin)
        h = max(0, h - 2 * border_margin)

        # Crop the image to the bounding box of the largest purple contour
        table_image_new = new_image[y:y+h, x:x+w]

        # Save the cropped table image
        cropped_table_path = f"{out_path}{filename}"
        cv2.imwrite(cropped_table_path, table_image_new)

# CONVERTIR SALAS DE PDF A PPTX ----

# Listar archivos de las salas
list_files = fnmatch.filter(os.listdir("./content/"), "*.pdf")

# Generación de salas
# # NOTE: Ajustar el `zoom_factor` para cambiar la resolución
for file in list_files:
    # Convertir archivos
    pdf_to_pptx("./content/", file, zoom_factor=3)

    # Directorio del archivo convertido
    dir_in = f"./pages/{file[0:len(file) - 4]}/"
    dir_out = f"./content/{file[0:len(file) - 4]}/"

    # Diapositiva
    list_slides = fnmatch.filter(os.listdir(dir_in), "*.png")

    # Extraer tablas
    for table in list_slides:
        get_tables(dir_in, dir_out, table)

    # Generar pptx
    list_tables = fnmatch.filter(os.listdir(dir_out), "*.png")
    list_out = list(set(list_slides) - set(list_tables))
    list_out.sort()

    # Crea una presentación en blanco
    presentation = Presentation()

    # Define el tamaño de la imagen y la posición
    left = top = Inches(0)

    for i in list_out:
        # Añade una diapositiva en blanco a la presentación
        slide = presentation.slides.add_slide(presentation.slide_layouts[5])
        # Añadir la imágena a la diapositiva creada
        slide.shapes.add_picture(f"{dir_in}{i}", left, top, height=Inches(7.5))

    # Guarda la presentación
    presentation.save(f"./content/{file[0:len(file) - 4]}.pptx")
