import os

def listar_estructura(directorio, nivel=0, profundidad=2):
    if nivel > profundidad:
        return
    for elemento in os.listdir(directorio):
        ruta = os.path.join(directorio, elemento)
        print("    " * nivel + "|-- " + elemento)
        if os.path.isdir(ruta):
            listar_estructura(ruta, nivel + 1, profundidad)

# Cambia '.' por la ruta raíz de tu proyecto si es necesario
listar_estructura(".", profundidad=2)
