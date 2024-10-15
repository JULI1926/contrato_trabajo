import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from docx import Document
from tkcalendar import DateEntry
from docx.shared import RGBColor
from num2words import num2words
from datetime import datetime, timedelta
import locale
import json
import os

# Definir las variables de reemplazo como globales
reemplazos = {}

# Variable global para almacenar la ruta del archivo cargado
archivo_cargado = None

# Cargar el archivo JSON
def cargar_datos_json(ruta_archivo):
    """Carga el archivo JSON y devuelve los datos."""
    with open('municipios.json', 'r', encoding='utf-8') as archivo_json:
        return json.load(archivo_json)
    

def procesar_datos(datos_json):
    """Procesa los datos JSON y devuelve una lista de departamentos y un diccionario de municipios."""
    # Extraer los departamentos únicos
    departamentos = sorted(set(dato["departamento"] for dato in datos_json))

    # Crear un diccionario con los municipios por departamento
    municipios_por_departamento = {}
    for dato in datos_json:
        depto = dato["departamento"]
        if depto not in municipios_por_departamento:
            municipios_por_departamento[depto] = []
        municipios_por_departamento[depto].append(dato["municipio"])

    return departamentos, municipios_por_departamento

salario_inicial = None

def actualizar_salario(event):
    global salario_inicial
    if salario_inicial is None:
        try:
            salario_inicial = int(salario_trabajador.get())
        except ValueError:
            salario_inicial = 0

    seleccion = jornada_trabajo.get()
    if seleccion == "TIEMPO COMPLETO":
        nuevo_salario = salario_inicial
    elif seleccion == "MEDIO TIEMPO":
        nuevo_salario = salario_inicial // 2
    elif seleccion == "POR HORAS":
        nuevo_salario = salario_inicial // 230
    else:
        nuevo_salario = salario_inicial

    salario_trabajador.delete(0, tk.END)
    salario_trabajador.insert(0, f"{nuevo_salario}")



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
                                
    

                                
    nombre_archivo = f"Contrato de Trabajo {reemplazos['[TERMINO]']} de {reemplazos['[TRABAJADOR]']}.docx"
    documento.save(nombre_archivo)



def reemplazar_salario_en_documento(doc_path, salario):
    # Convierte el salario a palabras
    if jornada_trabajo.get() == "POR HORAS":
    #"TIEMPO COMPLETO", "MEDIO TIEMPO", "POR HORAS"
        salario_palabras = num2words(salario, lang='es').replace('coma', 'mil')
        salario_texto = f"{salario:,} ({salario_palabras.upper()} PESOS M/CTE POR HORA.)"
    else:
        salario_palabras = num2words(salario, lang='es').replace('coma', 'mil')
        salario_texto = f"{salario:,} ({salario_palabras.upper()} PESOS M/CTE MENSUAL.)"
    
    # Diccionario de reemplazos
    reemplazos = {"[SALARIO]": salario_texto}
    
    # Imprimir el diccionario de reemplazos para verificar
    print("Diccionario de reemplazos:", reemplazos)

    return reemplazos
    
def calcular_fecha_fin(fecha_inicio, duracion):
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    fecha_inicio_dt = datetime.strptime(fecha_inicio, '%d/%m/%Y')
    fecha_fin_dt = fecha_inicio_dt + timedelta(days=duracion)
    return fecha_fin_dt.strftime('%d de %B del %Y').upper()



def deshabilitar_duracion_contrato(event):
    selected_option = termino_contrato.get()
    if selected_option in ["INDEFINIDO", "POR DURACION DE OBRA O LABOR"]:
        entrada_duracion_contrato.grid_remove()  # Oculta el campo
        # entrada_duracion_contrato.config(state='disabled')  # Alternativa: Deshabilita el campo
    else:
        entrada_duracion_contrato.grid()  # Muestra el campo
        # entrada_duracion_contrato.config(state='normal')  # Alternativa: Habilita el campo


# Función para validar la duración del período de prueba
def validar_duracion_prueba(event=None):
    termino = termino_contrato.get()
    try:
        duracion_contrato = int(entrada_duracion_contrato.get())
        duracion_prueba = int(entrada_duracion_prueba.get())
    except ValueError:
        return

    if termino == "A TÉRMINO FIJO":
        if duracion_prueba > duracion_contrato / 5:
            messagebox.showerror("Error", "No puede exceder la quinta parte de la duración del contrato")
            entrada_duracion_prueba.delete(0, "end")
    elif termino == "INDEFINIDO":
        if duracion_prueba > 60:
            messagebox.showerror("Error", "No puede exceder los 60 días (2 meses) de período de prueba")
            entrada_duracion_prueba.delete(0, "end")
    elif termino == "POR DURACION DE OBRA O LABOR":
        if duracion_prueba > duracion_contrato:
            messagebox.showerror("Error", "No puede exceder la duración del contrato")
            entrada_duracion_prueba.delete(0, "end")
        elif duracion_prueba > 60:
            messagebox.showerror("Error", "No puede exceder los 60 días (2 meses) de período de prueba")
            entrada_duracion_prueba.delete(0, "end")

def reemplazar_texto():
    global archivo_cargado
    global reemplazos
    
    fecha_inicio = fecha_inicio_contrato.get()
    
     # Inicializar la duración solo si el término del contrato es "A TÉRMINO FIJO"
    if termino_contrato.get() == "A TÉRMINO FIJO":
        if entrada_duracion_contrato.winfo_ismapped():
            duracion_texto = entrada_duracion_contrato.get()
            if duracion_texto:
                duracion = int(duracion_texto)
            else:
                duracion = 0  # O cualquier valor predeterminado que desees usar
        else:
            duracion = 0  # O cualquier valor predeterminado que desees usar
        fecha_fin = calcular_fecha_fin(fecha_inicio, duracion)
    else:
        fecha_fin = ""  # O cualquier valor predeterminado que desees usar
    

    # Inicializar la lista de campos faltantes
    campos_faltantes = []

    # Verificar si el campo de salario no está vacío
    salario_text = salario_trabajador.get().replace('.', '')
    if not salario_text:
        campos_faltantes.append("Salario")
    else:
        try:
            salario = int(salario_text)
        except ValueError:
            messagebox.showwarning("Advertencia", "El valor del salario no es válido.")
            return

    if archivo_cargado:
        documento = Document(archivo_cargado)
        fecha = fecha_nacimiento.get_date()

        # Obtener los reemplazos de salario solo si salario está definido
        if 'salario' in locals():
            reemplazos_salario = reemplazar_salario_en_documento(archivo_cargado, salario)
        else:
            reemplazos_salario = {}

        locale.setlocale(locale.LC_TIME, 'es_ES')  # Establecer el locale en español

        # Verificar cada campo
        if not entrada_empleador.get():
            campos_faltantes.append("Empleador")
        if not entrada_nit.get():
            campos_faltantes.append("N.I.T")
        if not entrada_representante_legal.get():
            campos_faltantes.append("Representante Legal")
        if not entrada_cc_representante_legal.get():
            campos_faltantes.append("C.C. Representante Legal")
        if not entrada_trabajador.get():
            campos_faltantes.append("Trabajador")
        if not entrada_cc_trabajador.get():
            campos_faltantes.append("C.C. Trabajador")
        if not entrada_ciudad.get():
            campos_faltantes.append("Ciudad")
        if not entrada_departamento.get():
            campos_faltantes.append("Departamento")
        if not estado_civil.get():
            campos_faltantes.append("Estado Civil")
        if not entrada_direccion.get():
            campos_faltantes.append("Dirección")
        if not entrada_telefono.get():
            campos_faltantes.append("Teléfono")
        if not entrada_cargo.get():
            campos_faltantes.append("Cargo")
        if not entrada_ciudad_contrato.get():
            campos_faltantes.append("Ciudad Contrato")
        if not entrada_departamento_contrato.get():
            campos_faltantes.append("Departamento Contrato")
        if not jornada_trabajo.get():
            campos_faltantes.append("Jornada")
        if not termino_contrato.get():
            campos_faltantes.append("Término del Contrato")
        if not fecha_inicio_contrato.get_date():
            campos_faltantes.append("Fecha de Inicio del Contrato")

        # Si hay campos faltantes, mostrar una alerta y no realizar los reemplazos
        if campos_faltantes:
            mensaje_error = "Los siguientes campos están vacíos:\n" + "\n".join(campos_faltantes)
            tk.messagebox.showerror("Error", mensaje_error)
            return

        reemplazos = {
            "[Empleador]": entrada_empleador.get().upper(),
            "[N.I.T]": entrada_nit.get(),
            "[REPRESENTANTE LEGAL]": entrada_representante_legal.get().upper(),
            "[C.C.]" : entrada_cc_representante_legal.get(),
            "[TRABAJADOR]": entrada_trabajador.get().upper(),
            "[C.CNo]": entrada_cc_trabajador.get(),
            "[CIUDAD]": str(entrada_ciudad.get()).upper(),
            "[DEPARTAMENTO]": str(entrada_departamento.get()).upper(),
            "[DIA]": str(fecha.day),
            "[MES]": fecha.strftime('%B').upper(),
            "[ANO]": str(fecha.year),
            "[ESTADO CIVIL]": estado_civil.get().upper(),           
            "[DIRECCION]": entrada_direccion.get().upper(),  
            "[TELEFONO]": entrada_telefono.get(), 
            "[CARGO]": entrada_cargo.get().upper(),
            "[CD_CONT]": str(entrada_ciudad_contrato.get()).upper(),    
            "[DPTO_CONT]": str(entrada_departamento_contrato.get()).upper(),
            "[JORNADA]": jornada_trabajo.get().upper(),
            "[TERMINO]": termino_contrato.get().upper(),
            "[FECHA_INICIO]": fecha_inicio_contrato.get_date().strftime('%d de %B del %Y').upper(),
            "[FECHA_FIN]": fecha_fin.upper(),
            "[FECHA_FIRMA]": fecha_firma_contrato.get_date().strftime('%d de %B del %Y').upper(),

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
#root.geometry("1920x1080")
root.configure(bg='#b0d4ec')
root.title("CONTRATO DE TRABAJO AYUDA SOCIAL Y LABORAL")

# Estilo personalizado para ttk.Entry
style = ttk.Style()
style.configure("Rounded.TEntry", padding=6, relief="flat", borderwidth=2, bordercolor="#b0d4ec")
style.map("Rounded.TEntry",
          fieldbackground=[('readonly', '#b0d4ec'), ('focus', '#e0f7fa')],
          background=[('active', '#b0d4ec')],
          bordercolor=[('focus', '#b0d4ec')])

# Crear un Frame contenedor para el Canvas y el Scrollbar
contenedor = tk.Frame(root, bg='#b0d4ec')
contenedor.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# Crear un Canvas
canvas = tk.Canvas(contenedor, bg='#b0d4ec')
canvas.pack(side="left", fill="both", expand=True)

# Añadir un Scrollbar al Canvas
scrollbar = ttk.Scrollbar(contenedor, orient="vertical", command=canvas.yview)
scrollbar.pack(side="right", fill="y")

# Configurar el Canvas para usar el Scrollbar
canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

# Crear otro Frame dentro del Canvas
frame_con_scroll = tk.Frame(canvas, bg='#b0d4ec')
canvas.create_window((0, 0), window=frame_con_scroll, anchor="nw")

# Crear un Frame adicional para centrar el contenido
frame_centrado = tk.Frame(frame_con_scroll, bg='#b0d4ec')
frame_centrado.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# # Crear un marco para agrupar los widgets dentro del Frame con scroll
# frame = tk.Frame(frame_con_scroll, padx=10, pady=10, bg='#b0d4ec')
# frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")


# Función para el desplazamiento con la rueda del mouse
def _on_mouse_wheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")


canvas.bind_all("<MouseWheel>", _on_mouse_wheel)

# Configurar las columnas y filas para que se expandan
for i in range(10):
    frame_con_scroll.grid_columnconfigure(i, weight=1)
for i in range(40):
    frame_con_scroll.grid_rowconfigure(i, weight=1)


# Registrar la función de validación
vcmd = (root.register(solo_letras), '%P')
vcmdnum = (root.register(solo_numeros), '%P')

# Subtítulo Datos del Empleador
tk.Label(frame_con_scroll, text="DATOS DEL EMPLEADOR", font=("Helvetica", 14, "bold")).grid(row=1, column=2, columnspan=4, padx=5, pady=10)

# Datos del Empleador
tk.Label(frame_con_scroll, text="NOMBRE DEL EMPLEADOR", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=2, column=1, padx=5, pady=5, sticky="e")
entrada_empleador = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14), validate="key")
entrada_empleador.grid(row=2, column=2, padx=5, pady=5, sticky="ew")

tk.Label(frame_con_scroll, text="N.I.T EMPLEADOR:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=2, column=3, padx=5, pady=5, sticky="e")
entrada_nit = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14))
entrada_nit.grid(row=2, column=4, padx=5, pady=5, sticky="ew")

# Espaciado entre filas
root.grid_rowconfigure(3, minsize=20)

# Subtítulo Datos del Representante Legal
tk.Label(frame_con_scroll, text="DATOS DEL REPRESENTANTE LEGAL", font=("Helvetica", 16, "bold")).grid(row=4, column=2, columnspan=4, padx=5, pady=10)

# Datos del Representante Legal
tk.Label(frame_con_scroll, text="REPRESENTANTE LEGAL:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=5, column=1, padx=5, pady=5, sticky="e")
entrada_representante_legal = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14), validate="key")
entrada_representante_legal.grid(row=5, column=2, padx=5, pady=5, sticky="ew")

tk.Label(frame_con_scroll, text="CC REPRESENTANTE LEGAL:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=5, column=3, padx=5, pady=5, sticky="e")
entrada_cc_representante_legal = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14), validate="key")
entrada_cc_representante_legal.grid(row=5, column=4, padx=5, pady=5, sticky="ew")

# Espaciado entre filas
root.grid_rowconfigure(6, minsize=20)

# Subtítulo Datos del Trabajador
tk.Label(frame_con_scroll, text="DATOS DEL TRABAJADOR", font=("Helvetica", 16, "bold")).grid(row=7, column=2, columnspan=4, padx=5, pady=10)

# Datos del Trabajador
tk.Label(frame_con_scroll, text="NOMBRE TRABAJADOR:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=8, column=1, padx=5, pady=5, sticky="e")
entrada_trabajador = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14))
entrada_trabajador.grid(row=8, column=2, padx=5, pady=5, sticky="ew")

tk.Label(frame_con_scroll, text="CC DEL TRABAJADOR:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=8, column=3, padx=5, pady=5, sticky="e")
entrada_cc_trabajador = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14))
entrada_cc_trabajador.grid(row=8, column=4, padx=5, pady=5, sticky="ew")

# Espaciado entre filas
root.grid_rowconfigure(9, minsize=20)

# Fecha y lugar de Nacimiento
tk.Label(frame_con_scroll, text="Fecha y lugar de Nacimiento", bg='#b0d4ec', font=("Helvetica", 12, "bold")).grid(row=10, column=2, columnspan=4, padx=5, pady=10)

tk.Label(frame_con_scroll, text="Fecha de Nacimiento:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=11, column=1, padx=5, pady=5, sticky="e")
fecha_nacimiento = DateEntry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14), date_pattern='dd/MM/yyyy')
fecha_nacimiento.delete(0, "end")
fecha_nacimiento.insert(0, "dd/MM/AAAA")
fecha_nacimiento.grid(row=11, column=2, padx=5, pady=5, sticky="ew")


def autompletar_municipios(departamentos, municipios_por_departamento): 

    global entrada_departamento, entrada_ciudad 
    
    # Función para actualizar el combobox de municipios cuando cambie el departamento
    def actualizar_municipios(event):
        departamento_seleccionado = entrada_departamento.get()
        entrada_ciudad["values"] = municipios_por_departamento.get(departamento_seleccionado, [])
        entrada_ciudad.set('')  # Limpiar la selección de municipio al cambiar el departamento

    
    # Label y combobox para el departamento
    tk.Label(frame_con_scroll, text="DEPARTAMENTO:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=11, column=3, padx=5, pady=5, sticky="e")
    entrada_departamento = ttk.Combobox(frame_con_scroll, values=departamentos, font=("Helvetica", 14))
    entrada_departamento.grid(row=11, column=4, padx=5, pady=5, sticky="ew")
    entrada_departamento.bind("<<ComboboxSelected>>", actualizar_municipios)

    # Label y combobox para el municipio
    tk.Label(frame_con_scroll, text="MUNICIPIO:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=12, column=1, padx=5, pady=5, sticky="e")
    entrada_ciudad = ttk.Combobox(frame_con_scroll, font=("Helvetica", 14))
    entrada_ciudad.grid(row=12, column=2, padx=5, pady=5, sticky="ew")

def main():
    # Cargar y procesar los datos JSON
    datos_json = cargar_datos_json('ruta/al/archivo.json')
    departamentos, municipios_por_departamento = procesar_datos(datos_json)

    # Inicializar la interfaz
    autompletar_municipios(departamentos, municipios_por_departamento)

if __name__ == "__main__":
    main()

# Configurar la fuente para los elementos del Combobox
root.option_add('*TCombobox*Listbox.font', ("Helvetica", 14))
root.option_add('*TCombobox.font', ("Helvetica", 14)) 

tk.Label(frame_con_scroll, text="ESTADO CIVIL:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=12, column=3, padx=5, pady=5, sticky="e")
estado_civil = ttk.Combobox(frame_con_scroll, values=["SOLTERO", "SOLTERA", "CASADO", "CASADA", "VIUDO", "VIUDA", "SEPARADO", "SEPARADA", "UNION LIBRE"], state="readonly")
estado_civil.set("Seleccione una opción ...")  # Valor por defecto
estado_civil.grid(row=12, column=4, padx=5, pady=5, sticky="ew")

# Dirección
tk.Label(frame_con_scroll, text="Dirección y Teléfono", bg='#b0d4ec', font=("Helvetica", 12, "bold")).grid(row=13, column=2, columnspan=4, padx=5, pady=10)

tk.Label(frame_con_scroll, text="DIRECCIÓN:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=14, column=1, padx=5, pady=5, sticky="e")
entrada_direccion = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14))
entrada_direccion.grid(row=14, column=2, padx=5, pady=5, sticky="ew")

tk.Label(frame_con_scroll, text="TELÉFONO:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=14, column=3, padx=5, pady=5, sticky="e")
entrada_telefono = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14))
entrada_telefono.grid(row=14, column=4, padx=5, pady=5, sticky="ew")

tk.Label(frame_con_scroll, text="TELÉFONO CONTACTO ADICIONAL:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=15, column=1, padx=5, pady=5, sticky="e")
entrada_telefono = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14))
entrada_telefono.grid(row=15, column=2, padx=5, pady=5, sticky="ew")

root.grid_rowconfigure(16, minsize=20)

# Datos del Contrato
tk.Label(frame_con_scroll, text="DATOS DEL CONTRATO", font=("Helvetica", 16, "bold")).grid(row=17, column=2, columnspan=4, padx=5, pady=10)

tk.Label(frame_con_scroll, text="CARGO QUE DESEMPEÑARÁ:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=18, column=1, padx=5, pady=5, sticky="e")
entrada_cargo = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14))
entrada_cargo.grid(row=18, column=2, padx=5, pady=5, sticky="ew")

tk.Label(frame_con_scroll, text="SALARIO BASE DEL TRABAJADOR:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=18, column=3, padx=5, pady=5, sticky="e")
salario_trabajador = ttk.Entry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14))
salario_trabajador.grid(row=18, column=4, padx=5, pady=5, sticky="ew")

# Espaciado entre filas
root.grid_rowconfigure(18, minsize=20)



def autompletar_municipios_contrato(departamentos, municipios_por_departamento):  
    global entrada_departamento_contrato, entrada_ciudad_contrato
    
    # Función para actualizar el combobox de municipios cuando cambie el departamento
    def actualizar_municipios(event):
        departamento_seleccionado = entrada_departamento_contrato.get()
        entrada_ciudad_contrato["values"] = municipios_por_departamento.get(departamento_seleccionado, [])
        entrada_ciudad_contrato.set('')  # Limpiar la selección de municipio al cambiar el departamento

    # Label y combobox para el departamento
    tk.Label(frame_con_scroll, text="DEPARTAMENTO DE LABOR:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=19, column=1, padx=5, pady=5, sticky="e")
    entrada_departamento_contrato = ttk.Combobox(frame_con_scroll, values=departamentos, font=("Helvetica", 14))
    entrada_departamento_contrato.grid(row=19, column=2, padx=5, pady=5, sticky="ew")
    entrada_departamento_contrato.bind("<<ComboboxSelected>>", actualizar_municipios)

    # Label y combobox para el municipio
    tk.Label(frame_con_scroll, text="MUNICIPIO DE LABOR:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=19, column=3, padx=5, pady=5, sticky="e")
    entrada_ciudad_contrato = ttk.Combobox(frame_con_scroll, font=("Helvetica", 14))
    entrada_ciudad_contrato.grid(row=19, column=4, padx=5, pady=5, sticky="ew")

def main():
    # Cargar y procesar los datos JSON
    datos_json = cargar_datos_json('ruta/al/archivo.json')
    departamentos, municipios_por_departamento = procesar_datos(datos_json)

    # Inicializar la interfaz
    autompletar_municipios_contrato(departamentos, municipios_por_departamento)

if __name__ == "__main__":
    main()


root.grid_rowconfigure(20, minsize=20)

tk.Label(frame_con_scroll, text="JORNADA DE TRABAJO:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=21, column=1, padx=5, pady=5, sticky="e")
jornada_trabajo = ttk.Combobox(frame_con_scroll, values=["TIEMPO COMPLETO", "MEDIO TIEMPO", "POR HORAS"], state="readonly")
jornada_trabajo.set("Seleccione una opción ...")  # Valor por defecto
jornada_trabajo.grid(row=21, column=2, padx=5, pady=5, sticky="ew")
jornada_trabajo.bind("<<ComboboxSelected>>", actualizar_salario)


tk.Label(frame_con_scroll, text="TÉRMINO DEL CONTRATO:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=21, column=3, padx=5, pady=5, sticky="e")
termino_contrato = ttk.Combobox(frame_con_scroll, values=["INDEFINIDO", "A TÉRMINO FIJO", "POR DURACION DE OBRA O LABOR"], state="readonly")
termino_contrato.set("Seleccione una opción ...")  # Valor por defecto
termino_contrato.grid(row=21, column=4, padx=5, pady=5, sticky="ew")
termino_contrato.bind("<<ComboboxSelected>>" , lambda event: (deshabilitar_duracion_contrato(), validar_duracion_prueba()))

# Espaciado entre filas
root.grid_rowconfigure(22, minsize=20)

tk.Label(frame_con_scroll, text="Fecha de Inicio de Contrato:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=23, column=1, padx=5, pady=5, sticky="e")
fecha_inicio_contrato = DateEntry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14), date_pattern='dd/MM/yyyy')
fecha_inicio_contrato.delete(0, "end")
fecha_inicio_contrato.insert(0, "dd/MM/AAAA")
fecha_inicio_contrato.grid(row=23, column=2, padx=5, pady=5, sticky="ew")

tk.Label(frame_con_scroll, text="Fecha de Firma de Contrato:", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=23, column=3, padx=5, pady=5, sticky="e")
fecha_firma_contrato = DateEntry(frame_con_scroll, style="Rounded.TEntry", font=("Helvetica", 14), date_pattern='dd/MM/yyyy')
fecha_firma_contrato.delete(0, "end")
fecha_firma_contrato.insert(0, "dd/MM/AAAA")
fecha_firma_contrato.grid(row=23, column=4, padx=5, pady=5, sticky="ew")

 # Label y combobox para el municipio
tk.Label(frame_con_scroll, text="DURACION DEL CONTRATO (EN DIAS):", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=24, column=1, padx=5, pady=5, sticky="e")
entrada_duracion_contrato = ttk.Entry(frame_con_scroll, font=("Helvetica", 14))
entrada_duracion_contrato.grid(row=24, column=2, padx=5, pady=5, sticky="ew")

 # Label y combobox para el municipio
tk.Label(frame_con_scroll, text="DURACION DEL PERIODO DE PRUEBA (EN DIAS):", bg='#b0d4ec', font=("Helvetica", 14, "bold italic")).grid(row=24, column=3, padx=5, pady=5, sticky="e")
entrada_duracion_prueba = ttk.Entry(frame_con_scroll, font=("Helvetica", 14))
entrada_duracion_prueba.grid(row=24, column=4, padx=5, pady=5, sticky="ew")
entrada_duracion_prueba.bind("<FocusOut>", validar_duracion_prueba)

# Cargar el documento
cargar_btn = tk.Button(frame_con_scroll, text="Cargar Documento", command=cargar_documento)
cargar_btn.grid(row=25, column=2, columnspan=2, pady=10, sticky="ew")

# Botón para reemplazar el texto
reemplazar_btn = tk.Button(frame_con_scroll, text="Reemplazar Texto", command=reemplazar_texto)
reemplazar_btn.grid(row=26, column=2, columnspan=2, pady=10, sticky="ew")

# Crear y colocar el Label para mostrar el archivo cargado
archivo_label = tk.Label(frame_con_scroll, text="No se ha cargado ningún documento.")
archivo_label.grid(row=27, column=2, columnspan=2, pady=10, sticky="ew")

# Cargar el documento por defecto al iniciar la aplicación
cargar_documento_por_defecto()

# Asegurarse de que el contenedor se expanda
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)
contenedor.grid_rowconfigure(0, weight=1)
contenedor.grid_columnconfigure(0, weight=1)

root.mainloop()