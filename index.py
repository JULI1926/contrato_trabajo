import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from docx import Document
from tkcalendar import DateEntry
from docx.shared import RGBColor
import locale
import os

# Variable global para almacenar la ruta del archivo cargado
archivo_cargado = None

def cargar_documento():
    global archivo_cargado
    archivo_cargado = filedialog.askopenfilename(
        title="Seleccionar documento Word",
        filetypes=[("Documentos Word", "*.docx")]
    )
    if archivo_cargado:
        archivo_label.config(text=f"Documento cargado: {archivo_cargado}")
    else:
        archivo_label.config(text="No se ha cargado ningún documento.")

def cargar_documento_por_defecto():
    global archivo_cargado
    archivo_cargado = os.path.join(os.getcwd(), "CONTRATO DE TRABAJO INDEFINIDO.docx")
    if os.path.exists(archivo_cargado):
        archivo_label.config(text=f"Documento cargado: {archivo_cargado}")
    else:
        archivo_label.config(text="No se encontró el documento por defecto.")

def reemplazar_texto_en_parrafo(parrafo, reemplazos):
    for texto_a_reemplazar, nuevo_texto in reemplazos.items():
        if texto_a_reemplazar in parrafo.text:
            for run in parrafo.runs:
                if texto_a_reemplazar in run.text:
                    run.text = run.text.replace(texto_a_reemplazar, nuevo_texto)
                    run.font.color.rgb = RGBColor(0, 0, 0)  # Establecer el color del texto a negro

def reemplazar_texto_en_documento(documento, reemplazos):
    # Reemplazar texto en párrafos
    for parrafo in documento.paragraphs:
        for run in parrafo.runs:
            for clave, valor in reemplazos.items():
                if clave in run.text:
                    print(f"Reemplazando {clave} con {valor}")
                    run.text = run.text.replace(clave, valor)
                else:
                    print(f"No se encontró {clave} en el párrafo: {run.text}")

    # Reemplazar texto en tablas
    for tabla in documento.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    for run in parrafo.runs:
                        for clave, valor in reemplazos.items():
                            if clave in run.text:
                                print(f"Reemplazando {clave} con {valor} en una celda")
                                run.text = run.text.replace(clave, valor)
                                run.font.color.rgb = RGBColor(0, 0, 0)  # Establecer el color del texto a negro
                            else:
                                print(f"No se encontró {clave} en la celda: {run.text}")

    documento.save("documento_modificado.docx")


def reemplazar_texto():
    global archivo_cargado
    if archivo_cargado:
        documento = Document(archivo_cargado)
        fecha = fecha_nacimiento.get_date()
        locale.setlocale(locale.LC_TIME, 'es_ES')  # Establecer el locale en español
        reemplazos = {
            "[Empleador]": entrada_empleador.get(),
            "[N.I.T]": entrada_nit.get(),
            "[REPRESENTANTE LEGAL]": entrada_representante_legal.get(),
            "[C.C.]" : entrada_cc_representante_legal.get(),
            "[TRABAJADOR]": entrada_trabajador.get(),
            "[C.CNo]": entrada_cc_trabajador.get(),
            "[CIUDAD]": entrada_ciudad.get(),
            "[DEPARTAMENTO]": entrada_departamento.get(),
            "[DIA]": str(fecha.day),
            "[MES]": fecha.strftime('%B').upper(),
            "[ANO]": str(fecha.year),
            "[ESTADO CIVIL]": estado_civil.get()
           
            
        }

        reemplazar_texto_en_documento(documento, reemplazos)

        # Guardar el documento
        archivo_guardado = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Documentos Word", "*.docx")]
        )
        if archivo_guardado:
            documento.save(archivo_guardado)
            messagebox.showinfo("Éxito", "El texto ha sido reemplazado y el documento guardado.")
        else:
            messagebox.showwarning("Advertencia", "No se ha guardado el documento.")
    else:
        messagebox.showwarning("Advertencia", "No se ha cargado ningún documento.")


root = tk.Tk()
root.configure(bg='#b0d4ec')
root.title("Reemplazar Texto en Documento Word")

# Crear un marco para agrupar los widgets
frame = tk.Frame(root, padx=10, pady=10)
frame.grid(row=0, column=0, padx=10, pady=10)

# Subtítulo Datos del Empleador
tk.Label(root, text="DATOS DEL EMPLEADOR", font=("Helvetica", 12, "bold")).grid(row=1, column=0, columnspan=4, padx=5, pady=10)


# Datos del Empleador
tk.Label(root, text="NOMBRE DEL EMPLEADOR", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=1, column=2, padx=5, pady=5, sticky="e")
entrada_empleador = tk.Entry(root)
entrada_empleador.grid(row=1, column=3, padx=5, pady=5)

tk.Label(root, text="N.I.T EMPLEADOR:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=1, column=4, padx=5, pady=5, sticky="e")
entrada_nit = tk.Entry(root)
entrada_nit.grid(row=1, column=5, padx=5, pady=5)

# Espaciado entre filas
root.grid_rowconfigure(2, minsize=20)


# Subtítulo Datos del Representante Legal
tk.Label(root, text="DATOS DEL REPRESENTANTE LEGAL", font=("Helvetica", 11, "bold")).grid(row=3, column=0, columnspan=4, padx=5, pady=10)

# Datos del Representante Legal
tk.Label(root, text="REPRESENTANTE LEGAL:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=3, column=2, padx=5, pady=5, sticky="e")
entrada_representante_legal = tk.Entry(root)
entrada_representante_legal.grid(row=3, column=3, padx=5, pady=5)

tk.Label(root, text="CC REPRESENTANTE LEGAL:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=3, column=4, padx=5, pady=5, sticky="e")
entrada_cc_representante_legal = tk.Entry(root)
entrada_cc_representante_legal.grid(row=3, column=5, padx=5, pady=5)

# Espaciado entre filas
root.grid_rowconfigure(4, minsize=20)



# Subtítulo Datos del Trabajador
tk.Label(root, text="DATOS DEL TRABAJADOR", font=("Helvetica", 11, "bold")).grid(row=6, column=0, columnspan=4, padx=5, pady=10)

# Datos del Trabajador
tk.Label(root, text="NOMBRE TRABAJADOR:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=6, column=2, padx=5, pady=5, sticky="e")
entrada_trabajador = tk.Entry(root)
entrada_trabajador.grid(row=6, column=3, padx=5, pady=5)

tk.Label(root, text="CC DEL TRABAJADOR:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=6, column=4, padx=5, pady=5, sticky="e")
entrada_cc_trabajador = tk.Entry(root)
entrada_cc_trabajador.grid(row=6, column=5, padx=5, pady=5)

# Espaciado entre filas
root.grid_rowconfigure(6, minsize=20)



#Fecha y lugar de Nacimiento

# Label para Fecha y lugar de Nacimiento
tk.Label(root, text="Fecha y lugar de Nacimiento", bg='#b0d4ec',font=("Helvetica", 12, "bold")).grid(row=7, column=0, columnspan=4, padx=5, pady=10)

# Label para Fecha de Nacimiento
tk.Label(root, text="Fecha de Nacimiento:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=7, column=2, padx=5, pady=5, sticky="e")
fecha_nacimiento = DateEntry(root, date_pattern='dd/MM/yyyy')
fecha_nacimiento.grid(row=7, column=3, padx=5, pady=5)

tk.Label(root, text="DEPARTAMENTO:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=7, column=4, padx=5, pady=5, sticky="e")
entrada_departamento = tk.Entry(root)
entrada_departamento.grid(row=7, column=5, padx=5, pady=5)

tk.Label(root, text="CIUDAD:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=8, column=2, padx=5, pady=5, sticky="e")
entrada_ciudad = tk.Entry(root)
entrada_ciudad.grid(row=8, column=3, padx=5, pady=5)

#Estado Civil
tk.Label(root, text="ESTADO CIVIL:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=8, column=4, padx=5, pady=5, sticky="e")
estado_civil = ttk.Combobox(root, values=["SOLTERO", "SOLTERA", "CASADO", "CASADA", "VIUDO", "VIUDA", "SEPARADO", "SEPARADA", "UNION LIBRE"], state="readonly")
estado_civil.set("SOLTERO")  # Valor por defecto
estado_civil.grid(row=8, column=5, padx=5, pady=5)


# Label para Fecha y lugar de Nacimiento
tk.Label(root, text="Dirección", bg='#b0d4ec',font=("Helvetica", 12, "bold")).grid(row=9, column=0, columnspan=4, padx=5, pady=10)

tk.Label(root, text="DIRECCIÓN:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=9, column=2, padx=5, pady=5, sticky="e")
entrada_ciudad = tk.Entry(root)
entrada_ciudad.grid(row=9, column=3, padx=5, pady=5)

tk.Label(root, text="TELÉFONO:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=9, column=4, padx=5, pady=5, sticky="e")
entrada_ciudad = tk.Entry(root)
entrada_ciudad.grid(row=9, column=5, padx=5, pady=5)

root.grid_rowconfigure(10, minsize=20)

############################  DATOS DEL CONTRATO  #########################
                                        

# Subtítulo Datos del Contrato   
tk.Label(root, text="DATOS DEL CONTRATO", font=("Helvetica", 12, "bold")).grid(row=11, column=0, columnspan=4, padx=5, pady=10)

# Datos del Contrrato
tk.Label(root, text="NOMBRE TRABAJADOR:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=11, column=2, padx=5, pady=5, sticky="e")
entrada_trabajador = tk.Entry(root)
entrada_trabajador.grid(row=11, column=3, padx=5, pady=5)

tk.Label(root, text="CC DEL TRABAJADOR:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=11, column=4, padx=5, pady=5, sticky="e")
entrada_cc_trabajador = tk.Entry(root)
entrada_cc_trabajador.grid(row=11, column=5, padx=5, pady=5)

# Espaciado entre filas
root.grid_rowconfigure(12, minsize=20)



#Fecha y lugar de Nacimiento

# Label para Fecha y lugar de Nacimiento
tk.Label(root, text="Fecha y lugar de Nacimiento", bg='#b0d4ec',font=("Helvetica", 13, "bold")).grid(row=11, column=0, columnspan=4, padx=5, pady=10)

# Label para Fecha de Nacimiento
tk.Label(root, text="Fecha de Nacimiento:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=13, column=2, padx=5, pady=5, sticky="e")
fecha_nacimiento = DateEntry(root, date_pattern='dd/MM/yyyy')
fecha_nacimiento.grid(row=13, column=3, padx=5, pady=5)

tk.Label(root, text="DEPARTAMENTO:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=14, column=4, padx=5, pady=5, sticky="e")
entrada_departamento = tk.Entry(root)
entrada_departamento.grid(row=14, column=5, padx=5, pady=5)

tk.Label(root, text="CIUDAD:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=15, column=2, padx=5, pady=5, sticky="e")
entrada_ciudad = tk.Entry(root)
entrada_ciudad.grid(row=15, column=3, padx=5, pady=5)

#Estado Civil
tk.Label(root, text="ESTADO CIVIL:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=16, column=4, padx=5, pady=5, sticky="e")
estado_civil = ttk.Combobox(root, values=["SOLTERO", "SOLTERA", "CASADO", "CASADA", "VIUDO", "VIUDA", "SEPARADO", "SEPARADA", "UNION LIBRE"], state="readonly")
estado_civil.set("SOLTERO")  # Valor por defecto
estado_civil.grid(row=16, column=5, padx=5, pady=5)


# Label para Fecha y lugar de Nacimiento
tk.Label(root, text="Dirección", bg='#b0d4ec',font=("Helvetica", 12, "bold")).grid(row=17, column=0, columnspan=4, padx=5, pady=10)

tk.Label(root, text="DIRECCIÓN:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=17, column=2, padx=5, pady=5, sticky="e")
entrada_ciudad = tk.Entry(root)
entrada_ciudad.grid(row=17, column=3, padx=5, pady=5)

tk.Label(root, text="TELÉFONO:", bg='#b0d4ec', font=("Helvetica", 10, "bold italic")).grid(row=17, column=4, padx=5, pady=5, sticky="e")
entrada_ciudad = tk.Entry(root)
entrada_ciudad.grid(row=17, column=5, padx=5, pady=(10, 20))




# Cargar el documento
cargar_btn = tk.Button(root, text="Cargar Documento", command=cargar_documento)
cargar_btn.grid(row=18, column=0, columnspan=2, pady=10)

# Botón para reemplazar el texto
reemplazar_btn = tk.Button(root, text="Reemplazar Texto", command=reemplazar_texto)
reemplazar_btn.grid(row=19, column=0, columnspan=2, pady=10)

# Crear y colocar el Label para mostrar el archivo cargado
archivo_label = tk.Label(root, text="No se ha cargado ningún documento.")
archivo_label.grid(row=20, column=0, columnspan=2, pady=10)

# Cargar el documento por defecto al iniciar la aplicación
cargar_documento_por_defecto()

root.mainloop()