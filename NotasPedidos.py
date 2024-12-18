import os
import re
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

mapeo_proveedores = {
    'A0018': 'Drogueria Kellerhoff SA',
    'A1114': 'Monroe Americana SA',
    'A0307': 'Drogueria Cofarsur SACIF',
    'A0418': 'Suizo Argentina SA',
    'A0567': 'Drogueria Cofarsur SACIF'
}

def extraer_codigo_factura(operacion):
    """
    Extrae el código de factura después de "Fact.:", devolviendo solo los primeros 5 caracteres
    """
    if pd.notnull(operacion):
        # Busca "Fact.:" seguido de cualquier cadena alfanumérica
        match = re.search(r'Fact\.:\s*([A-Z0-9]+)', str(operacion))
        if match:
            return match.group(1)[:5]  # Devuelve los primeros 5 caracteres del código capturado
    return None

def mapear_proveedor(row):
    """
    Mapea el proveedor basado en los primeros 5 caracteres del código de factura
    """
    try:
        codigo = extraer_codigo_factura(row.get('Operación', None))  # Extrae el código de la operación
        if codigo in mapeo_proveedores:
            return mapeo_proveedores[codigo]  # Mapea al proveedor si coincide
        return row.get('Proveedor/Cliente', None)  # Si no coincide, devuelve el original
    except Exception as e:
        print(f"Error en mapear_proveedor para la fila: {row}. Error: {e}")
        return None

# PROCESAMOS LA CARPETA CON NOTA DE PEDIDOS
def procesar_archivos(carpeta):
    columnas = [
        "FECHA", "COMPRADOR", "LABORATORIO", "IMPORTE PEDIDO", "CANTIDAD PEDIDO", "DROGUERIA", "Num Cuenta",
        "Can", "Codebar", "Producto", "Precio", 
        "drog", "Desc.", "Costo", "Imp. Total"
    ]
    
    datos_consolidados = []

    for archivo in os.listdir(carpeta):
        if archivo.endswith(".xlsx"):
            ruta_archivo = os.path.join(carpeta, archivo)
            xls = pd.ExcelFile(ruta_archivo)
            
            for hoja in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=hoja)

                if df.empty:
                    print(f"La hoja {hoja} está vacía en el archivo {archivo}.")
                    continue
                print(f"Hoja: {hoja} de {archivo}")
                try:
                    fecha = df.iloc[0, 8]
                    comprador = df.iloc[4, 2]
                    importe_pedido = df.iloc[4,5]
                    cantidad_pedido = df.iloc[5,5]
                    laboratorio = df.iloc[0, 4]
                    drogueria = df.iloc[2, 4]
                    # Dividir DROGUERIA en nombre y número de cuenta
                    nombre_drogueria, num_cuenta = dividir_drogueria(drogueria)
                    
                    tabla = df.iloc[9:-1, [0, 1, 2, 3, 4, 5, 6, 7, 8]]
                    tabla.columns = ["Can", "Codebar", "Producto", "Cantidad", "Precio", "drog", "Desc.", "Costo", "Imp. Total"]
                    for _, fila in tabla.iterrows():
                        if fila["Can"] == 0:
                            continue
                        codebar = str(fila["Codebar"])
                        datos_consolidados.append([
                            fecha, comprador, laboratorio,importe_pedido , cantidad_pedido, nombre_drogueria, num_cuenta,
                            fila["Can"], fila["Codebar"], fila["Producto"],
                            fila["Precio"], fila["drog"], 
                            fila["Desc."], fila["Costo"], fila["Imp. Total"]
                        ])
                except Exception as e:
                    print(f"Error procesando hoja en {archivo}: {e}")

    consolidado_df = pd.DataFrame(datos_consolidados, columns=columnas)
    
    # Pedir al usuario dónde guardar el archivo
    archivo_salida = filedialog.asksaveasfilename(
        title="Guardar archivo consolidado",
        defaultextension=".xlsx",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )

    if archivo_salida:
        consolidado_df.to_excel(archivo_salida, index=False)
        return archivo_salida
    else:
        messagebox.showwarning("Cancelado", "No se guardó el archivo.")
        return None


def dividir_drogueria(drogueria):
    """
    Divide la columna DROGUERIA en nombre y número de cuenta.
    La parte del número debe ser un número entero o decimal válido.
    """
    if pd.notnull(drogueria):
        drogueria = str(drogueria).strip()  # Aseguramos que no haya espacios iniciales o finales
        # Expresión regular para separar texto y número al final
        match = re.match(r'^(.*?)(\d+)$', drogueria)
        
        if match:
            nombre = match.group(1).strip()  # Todo antes del número
            num_cuenta = int(match.group(2))  # Convertir el número a entero
            return nombre, num_cuenta
        
        return drogueria, None  # Si no hay número al final
    return None, None
# Función para seleccionar la carpeta y procesar los archivos
def seleccionar_carpeta():
    carpeta = filedialog.askdirectory(title="Seleccionar Carpeta")
    if carpeta:
        try:
            archivo_salida = procesar_archivos(carpeta)
            messagebox.showinfo("Éxito", f"Datos consolidados guardados en:\n{archivo_salida}")
        except Exception as e:
            messagebox.showerror("Error", f"Hubo un problema procesando los archivos:\n{e}")

# Función para filtrar datos basado en el proveedor/cliente
def filtrar_por_proveedor(archivo):
    try:
        # Detectar el tipo de archivo y cargar los datos
        if archivo.endswith(".csv"):
            try:
                # Intentar leer con separador de espacios
                df = pd.read_csv(
                    archivo,
                    sep=';',  # Separador de uno o más espacios
                    engine='python',  # Motor de Python para manejo más flexible
                    encoding="ISO-8859-1",
                    on_bad_lines="skip"
                )

                # Imprimir columnas para verificar
            except Exception as e:
                # Si falla, intentar con delimitador alternativo
                print(f"Error al leer el archivo: {e}")
                raise
        elif archivo.endswith(".xlsx"):
            df = pd.read_excel(archivo)
        else:
            raise ValueError("Formato de archivo no soportado. Seleccione un archivo .csv o .xlsx.")
        # Verificar si la columna "Proveedor/Cliente" existe
        if "Operación" not in df.columns:
            raise ValueError("El archivo no contiene la columna 'Proveedor/Cliente'.")
        # Aplicar mapeo para rellenar "Proveedor/Cliente"
        df["Proveedor/Cliente"] = df.apply(mapear_proveedor, axis=1)

        try:
            df["Cantidad"] = df["Cantidad"].apply(convertir_numero)
            df["Nro.Lote"] = df["Nro.Lote"].apply(convertir_numero)
            df["Costo"] = df["Costo"].apply(convertir_numero)
            df["Total Costo"] = df["Total Costo"].apply(convertir_numero)
            df["Unitario"] = df["Unitario"].apply(convertir_numero)
            df["Total"] = df["Total"].apply(convertir_numero)
        except Exception as e:
            print(f"Error en conversión de columnas numéricas: {e}")
        #columnas a tipos de datos adecuados con manejo de errores
        df["Cod.Barras"] = df["Cod.Barras"].apply(convertir_numero)
        df["Cantidad"]= df["Cantidad"].apply(convertir_numero)
        df["Nro.Lote"]=df["Nro.Lote"].apply(convertir_numero)
        df["Costo"] = df["Costo"].apply(convertir_numero)
        df["Total Costo"] = df["Total Costo"].apply(convertir_numero)
        df["Unitario"] = df["Unitario"].apply(convertir_numero)
        df["Total"] = df["Total"].apply(convertir_numero)

        # Agregar las 3 nuevas columnas
        df["Cargado"] = "Cargado"
        df["Cantidad Cargada"] = df["Cantidad"]
        df["Fecha Cargada"] = df["Fecha"]

        # Filtrar las filas donde la columna NO comienza con "SANCHEZ ANTONIOLLI"
        filtro = ~df["Operación"].str.startswith("PD X ", na=False)


        datos_filtrados = df[filtro]

        if datos_filtrados.empty:
            messagebox.showinfo("Sin resultados", "No quedaron datos después del filtrado.")
            return

        # Pedir al usuario dónde guardar el archivo filtrado en formato Excel
        archivo_salida = filedialog.asksaveasfilename(
            title="Guardar archivo filtrado",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )

        if archivo_salida:
            datos_filtrados.to_excel(archivo_salida, index=False)
            messagebox.showinfo("Éxito", f"Archivo filtrado guardado en:\n{archivo_salida}")
        else:
            messagebox.showwarning("Cancelado", "No se guardó el archivo.")

    except Exception as e:
        messagebox.showerror("Error", f"Hubo un problema procesando el archivo:\n{e}")

# Función para seleccionar un archivo y aplicar el filtro
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(
        title="Seleccionar Archivo",
        filetypes=[("Archivos CSV o Excel", "*.csv *.xlsx")]
    )
    if archivo:
        filtrar_por_proveedor(archivo)

def convertir_numero(valor):
    try:
        # Reemplazar coma por punto y convertir a float
        return float(str(valor).replace(',', '.'))
    except:
        return 0
# Interfaz con customtkinter
def main():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")

    app = ctk.CTk()
    app.title("Consolidar y Filtrar Excel")
    app.geometry("400x300")

    etiqueta = ctk.CTkLabel(app, text="Herramientas para Archivos Excel", font=("Arial", 16))
    etiqueta.pack(pady=20)

    boton_consolidar = ctk.CTkButton(app, text="Consolidar Archivos", command=seleccionar_carpeta)
    boton_consolidar.pack(pady=20)

    boton_filtrar = ctk.CTkButton(app, text="Arreglo Quantio", command=seleccionar_archivo)
    boton_filtrar.pack(pady=20)

    app.mainloop()

if __name__ == "__main__":
    main()