import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from docx import Document
from tkcalendar import DateEntry
from docx.shared import RGBColor
from num2words import num2words
import locale
import os

# Variable global para almacenar la ruta del archivo cargado
archivo_cargado = None

def solo_letras(char):
    if char.isalpha() or char == "":
        return True
    else:
        messagebox.showerror("Entrada inválida", "Solo se permiten letras.")
        return False
    
def solo_numeros(char):
    # Verifica si el carácter ingresado es un número o si está vacío (para permitir borrar)
    if char.isdigit() or char == "":
        return True
    else:
        # Si no es un número, muestra un mensaje de error
        messagebox.showerror("Entrada inválida", "Solo se permiten números.")
        return False

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


def reemplazar_texto_en_documento(documento, reemplazos):
    for parrafo in documento.paragraphs:
        for clave, valor in reemplazos.items():
            if clave in parrafo.text:
                print(f"Reemplazando {clave} con {valor} en el párrafo: {parrafo.text}")
                parrafo.text = parrafo.text.replace(clave, valor)


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
                                
    documento.save("documento_modificado.docx")



def reemplazar_salario_en_documento(doc_path, salario):
    # Convierte el salario a palabras
    salario_palabras = num2words(salario, lang='es').replace('coma', 'mil')
    salario_texto = f"{salario:,} ({salario_palabras.upper()} PESOS M/CTE.)"
    
    # Diccionario de reemplazos
    reemplazos = {"[SALARIO]": salario_texto}
    
    # Imprimir el diccionario de reemplazos para verificar
    print("Diccionario de reemplazos:", reemplazos)

    return reemplazos
    
    '''
    # Cargar el documento
    documento = Document(doc_path)
    
    # Verificar que el documento se ha cargado correctamente
    print("Documento cargado correctamente.")
    
    # Reemplazar el texto en el documento
    reemplazar_texto_en_documento(documento, reemplazos)
    
    '''


'''
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
'''

def reemplazar_texto():
    global archivo_cargado
    salario = int(salario_trabajador.get().replace('.', ''))
    if archivo_cargado:
        documento = Document(archivo_cargado)
        fecha = fecha_nacimiento.get_date()

        # Obtener los reemplazos de salario
        reemplazos_salario = reemplazar_salario_en_documento(archivo_cargado, salario)   


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
            "[ESTADO CIVIL]": estado_civil.get(),           
            "[DIRECCION]": entrada_direccion.get(),  
            "[TELEFONO]": entrada_telefono.get(), 
            "[CARGO]": entrada_cargo.get()
            
            
        }

        # Combinar los diccionarios de reemplazos
        reemplazos.update(reemplazos_salario)

        # Imprimir el diccionario de reemplazos combinado para verificar
        print("Diccionario de reemplazos (combinado):", reemplazos)

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
root.title("CONTRATO DE TRABAJO AYUDA SOCIAL Y LABORAL")

# Estilo personalizado para tttk.Entry
style = ttk.Style()
style.configure("Rounded.TEntry", padding=6, relief="flat", borderwidth=2, bordercolor="#b0d4ec")
style.map("Rounded.TEntry",
          fieldbackground=[('readonly', '#b0d4ec'), ('focus', '#e0f7fa')],
          background=[('active', '#b0d4ec')],
          bordercolor=[('focus', '#b0d4ec')])

# Crear un marco para agrupar los widgets
frame = tk.Frame(root, padx=10, pady=10, bg='#b0d4ec')
frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# Configurar las columnas y filas para que se expandan
for i in range(6):
    root.grid_columnconfigure(i, weight=1)
for i in range(21):
    root.grid_rowconfigure(i, weight=1)


# Registrar la función de validación
vcmd = (root.register(solo_letras), '%P')
vcmdnum = (root.register(solo_numeros), '%P')

# Subtítulo Datos del Empleador
tk.Label(root, text="DATOS DEL EMPLEADOR", font=("Helvetica", 14, "bold")).grid(row=1, column=2, columnspan=4, padx=5, pady=10)

# Datos del Empleador
tk.Label(root, text="NOMBRE DEL EMPLEADOR", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=2, column=1, padx=5, pady=5, sticky="e")
entrada_empleador = ttk.Entry(root, style="Rounded.TEntry", validate="key", validatecommand=vcmd)
entrada_empleador.grid(row=2, column=2, padx=5, pady=5, sticky="ew")

tk.Label(root, text="N.I.T EMPLEADOR:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=2, column=3, padx=5, pady=5, sticky="e")
entrada_nit = ttk.Entry(root, style="Rounded.TEntry")
entrada_nit.grid(row=2, column=4, padx=5, pady=5, sticky="ew")

# Espaciado entre filas
root.grid_rowconfigure(3, minsize=20)

# Subtítulo Datos del Representante Legal
tk.Label(root, text="DATOS DEL REPRESENTANTE LEGAL", font=("Helvetica", 16, "bold")).grid(row=4, column=2, columnspan=4, padx=5, pady=10)

# Datos del Representante Legal
tk.Label(root, text="REPRESENTANTE LEGAL:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=5, column=1, padx=5, pady=5, sticky="e")
entrada_representante_legal = ttk.Entry(root, style="Rounded.TEntry", validate="key", validatecommand=vcmd)
entrada_representante_legal.grid(row=5, column=2, padx=5, pady=5, sticky="ew")

tk.Label(root, text="CC REPRESENTANTE LEGAL:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=5, column=3, padx=5, pady=5, sticky="e")
entrada_cc_representante_legal = ttk.Entry(root, style="Rounded.TEntry", validate="key", validatecommand=vcmdnum)
entrada_cc_representante_legal.grid(row=5, column=4, padx=5, pady=5, sticky="ew")

# Espaciado entre filas
root.grid_rowconfigure(6, minsize=20)

# Subtítulo Datos del Trabajador
tk.Label(root, text="DATOS DEL TRABAJADOR", font=("Helvetica", 16, "bold")).grid(row=7, column=2, columnspan=4, padx=5, pady=10)

# Datos del Trabajador
tk.Label(root, text="NOMBRE TRABAJADOR:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=8, column=1, padx=5, pady=5, sticky="e")
entrada_trabajador = ttk.Entry(root, style="Rounded.TEntry")
entrada_trabajador.grid(row=8, column=2, padx=5, pady=5, sticky="ew")

tk.Label(root, text="CC DEL TRABAJADOR:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=8, column=3, padx=5, pady=5, sticky="e")
entrada_cc_trabajador = ttk.Entry(root, style="Rounded.TEntry")
entrada_cc_trabajador.grid(row=8, column=4, padx=5, pady=5, sticky="ew")

# Espaciado entre filas
root.grid_rowconfigure(9, minsize=20)

# Fecha y lugar de Nacimiento
tk.Label(root, text="Fecha y lugar de Nacimiento", bg='#b0d4ec', font=("Helvetica", 12, "bold")).grid(row=10, column=2, columnspan=4, padx=5, pady=10)

tk.Label(root, text="Fecha de Nacimiento:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=11, column=1, padx=5, pady=5, sticky="e")
fecha_nacimiento = DateEntry(root, style="Rounded.TEntry", date_pattern='dd/MM/yyyy')
fecha_nacimiento.grid(row=11, column=2, padx=5, pady=5, sticky="ew")

tk.Label(root, text="DEPARTAMENTO:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=11, column=3, padx=5, pady=5, sticky="e")
entrada_departamento = ttk.Entry(root, style="Rounded.TEntry")
entrada_departamento.grid(row=11, column=4, padx=5, pady=5, sticky="ew")

tk.Label(root, text="CIUDAD:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=12, column=1, padx=5, pady=5, sticky="e")
entrada_ciudad = ttk.Entry(root, style="Rounded.TEntry")
entrada_ciudad.grid(row=12, column=2, padx=5, pady=5, sticky="ew")

tk.Label(root, text="ESTADO CIVIL:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=12, column=3, padx=5, pady=5, sticky="e")
estado_civil = ttk.Combobox(root, values=["SOLTERO", "SOLTERA", "CASADO", "CASADA", "VIUDO", "VIUDA", "SEPARADO", "SEPARADA", "UNION LIBRE"], state="readonly")
estado_civil.set("SOLTERO")  # Valor por defecto
estado_civil.grid(row=12, column=4, padx=5, pady=5, sticky="ew")

# Dirección
tk.Label(root, text="Dirección y Teléfono", bg='#b0d4ec', font=("Helvetica", 12, "bold")).grid(row=13, column=2, columnspan=4, padx=5, pady=10)

tk.Label(root, text="DIRECCIÓN:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=14, column=1, padx=5, pady=5, sticky="e")
entrada_direccion = ttk.Entry(root, style="Rounded.TEntry")
entrada_direccion.grid(row=14, column=2, padx=5, pady=5, sticky="ew")

tk.Label(root, text="TELÉFONO:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=14, column=3, padx=5, pady=5, sticky="e")
entrada_telefono = ttk.Entry(root, style="Rounded.TEntry")
entrada_telefono.grid(row=14, column=4, padx=5, pady=5, sticky="ew")

root.grid_rowconfigure(15, minsize=20)

# Datos del Contrato
tk.Label(root, text="DATOS DEL CONTRATO", font=("Helvetica", 12, "bold")).grid(row=16, column=2, columnspan=4, padx=5, pady=10)

tk.Label(root, text="CARGO QUE DESEMPEÑARÁ:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=17, column=1, padx=5, pady=5, sticky="e")
entrada_cargo = ttk.Entry(root, style="Rounded.TEntry")
entrada_cargo.grid(row=17, column=2, padx=5, pady=5, sticky="ew")

tk.Label(root, text="SALARIO DEL TRABAJADOR:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=17, column=3, padx=5, pady=5, sticky="e")
salario_trabajador = ttk.Entry(root, style="Rounded.TEntry")
salario_trabajador.grid(row=17, column=4, padx=5, pady=5, sticky="ew")

# Espaciado entre filas
root.grid_rowconfigure(18, minsize=20)

tk.Label(root, text="DEPARTAMENTO:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=19, column=3, padx=5, pady=5, sticky="e")
entrada_departamento_contrato = ttk.Entry(root, style="Rounded.TEntry")
entrada_departamento_contrato.grid(row=19, column=4, padx=5, pady=5, sticky="ew")

tk.Label(root, text="CIUDAD:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=19, column=1, padx=5, pady=5, sticky="e")
entrada_ciudad_contrato = ttk.Entry(root, style="Rounded.TEntry")
entrada_ciudad_contrato.grid(row=19, column=2, padx=5, pady=5, sticky="ew")

root.grid_rowconfigure(21, minsize=20)

# Cargar el documento
cargar_btn = tk.Button(root, text="Cargar Documento", command=cargar_documento)
cargar_btn.grid(row=22, column=2, columnspan=2, pady=10, sticky="ew")

# Botón para reemplazar el texto
reemplazar_btn = tk.Button(root, text="Reemplazar Texto", command=reemplazar_texto)
reemplazar_btn.grid(row=23, column=2, columnspan=2, pady=10, sticky="ew")

# Crear y colocar el Label para mostrar el archivo cargado
archivo_label = tk.Label(root, text="No se ha cargado ningún documento.")
archivo_label.grid(row=24, column=2, columnspan=2, pady=10, sticky="ew")

# Cargar el documento por defecto al iniciar la aplicación
cargar_documento_por_defecto()

root.mainloop()