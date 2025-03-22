import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from sklearn.impute import SimpleImputer
from sklearn.covariance import EllipticEnvelope
from pyod.models.knn import KNN

def seleccionar_hoja(nombre_archivo):
    # Obtener la lista de hojas en el archivo Excel
    hojas = pd.ExcelFile(nombre_archivo).sheet_names
    
    # Mostrar la lista de nombres de las hojas
    print("Hojas disponibles:")
    for i, hoja in enumerate(hojas):
        print(f"{i + 1}: {hoja}")
    
    # Preguntar al usuario cuál hoja desea cargar
    nombre_hoja = input("Escribe el nombre de la hoja que deseas cargar: ")
    
    # Verificar que el nombre de la hoja sea válido
    if nombre_hoja in hojas:
        return nombre_hoja
    else:
        print("Nombre de hoja no válido. Por favor, intenta de nuevo.")
        return seleccionar_hoja(nombre_archivo)

# Modificación en la función de lectura del archivo para incluir la selección de hoja
def leer_archivo():
    # Pedir nombre del archivo al usuario y agregar la extensión .xlsx
    nombre_archivo = input("Ingresa el nombre del archivo (sin extensión): ") + '.xlsx'
    
    # Seleccionar la hoja a cargar
    nombre_hoja = seleccionar_hoja(nombre_archivo)
    
    # Leer la hoja seleccionada del archivo .xlsx
    df = pd.read_excel(nombre_archivo, sheet_name=nombre_hoja)
    
    # Convertir la primera columna (que contiene las fechas) a tipo datetime
    df[df.columns[0]] = pd.to_datetime(df[df.columns[0]], format='%d:%m:%Y %H:%M:%S')

    # Retornar el DataFrame ya sea filtrado o completo
    return df

def aplicar_intervalo_fechas(df):
    """Función que pide un intervalo de fechas y filtra el DataFrame"""
    fecha_inicial_str = input("Ingresa la fecha inicial (dd:mm:YY hh:mm:ss): ")
    fecha_final_str = input("Ingresa la fecha final (dd:mm:YY hh:mm:ss): ")
    fecha_inicial = pd.to_datetime(fecha_inicial_str, format='%d:%m:%Y %H:%M:%S')
    fecha_final = pd.to_datetime(fecha_final_str, format='%d:%m:%Y %H:%M:%S')
    df_filtrado = df[(df[df.columns[0]] >= fecha_inicial) & (df[df.columns[0]] <= fecha_final)]
    print(f"Datos filtrados entre {fecha_inicial} y {fecha_final}")
    return df_filtrado

def solicitar_confirmacion(mensaje):
    """Función para confirmar una elección del usuario (retorna True si es 'y', False si es 'n')"""
    while True:
        respuesta = input(f"{mensaje} (y/n): ").lower()
        if respuesta in ["y", "n"]:
            return respuesta == 'y'
        else:
            print("Opción no válida. Por favor, ingresa 'y' para sí o 'n' para no.")

def preguntar_intervalo_fechas(df):
    """Función para preguntar si se desea aplicar intervalo de fechas y realizar el filtrado"""
    df_filtrado = df
    cambio_hecho = False  # Bandera para saber si ya se aplicó un filtrado inicial

    while True:
        if not cambio_hecho:  # Primera vez, preguntar si se desea filtrar
            if solicitar_confirmacion("¿Desea filtrar por intervalo de fechas?"):
                if solicitar_confirmacion("¿Estás seguro de aplicar este filtro?"):
                    df_filtrado = aplicar_intervalo_fechas(df)
                    cambio_hecho = True  # Indicar que se hizo el primer filtrado
                else:
                    continue
            else:
                if solicitar_confirmacion("¿Deseas ver el archivo completo sin filtrar?"):
                    print(df)
                    return df
                else:
                    continue

        # Después de la primera vez, preguntar si se desea cambiar el intervalo
        if cambio_hecho:
            if solicitar_confirmacion("¿Deseas cambiar el intervalo de fechas?"):
                df_filtrado = aplicar_intervalo_fechas(df)
            else:
                print(df_filtrado)
                return df_filtrado
            
def listar_columnas_con_estado(df):
    columnas = df.columns[df.columns != 'Fecha']
    total_filas = len(df)
    
    print("Columnas disponibles para análisis:")
    for col in columnas:
        num_validos = df[col].notna().sum()
        porcentaje_validos = num_validos / total_filas
        
        # Generar mensaje según el estado de la columna
        if num_validos == 0:
            mensaje = "[Columna Vacía]"
        elif porcentaje_validos < 0.25:
            mensaje = "[Columna con pocos datos]"
        else:
            mensaje = ""
        
        print(f"- {col} {mensaje}".strip())
            
def mostrar_menu(df_final):
    """Función que muestra un menú con opciones para análisis"""
    while True:
        print("\nMenú de Opciones:")
        print("1. Análisis Estadístico Univariado")
        print("2. Análisis Estadístico entre 2 Variables")
        print("3. Análisis de Series Temporales entre 4 Variables")
        print("4. Salir")
        
        opcion = input("Selecciona una opción (1-5): ")

        if opcion == "1":
            print("Has seleccionado Estadística Univariada.")
            # Lógica para Estadística Univariada
            estadistica_univariada(df_final)
        elif opcion == "2":
            print("Has seleccionado Análisis Estadístico entre 2 Variables")
            # Lógica para Estadística Multivariada
            analisis_multivariado(df_final)
        elif opcion == "3":
            print("Has seleccionado Análisis de Series Temporales entre 4 Variables")
            # Lógica para Capacidad Hidrociclón
        elif opcion == "4":
            print("Saliendo del programa.")
            break
        else:
            print("Opción no válida. Por favor selecciona una opción entre 1 y 4.")

def unidades_de_medida(variable):
    # Unidades de medida
    unidades = {
        "Peso": ["kg", "g", "toneladas"],
        "Presión": ["Pa", "bar", "psi"],
        "Velocidad": ["m/s", "km/h", "rpm"],
        "Temperatura": ["°C", "°F", "K"],
        "Porcentaje": ["%"],
        "TPH": ["TPH"],
    }

    print("Selecciona una unidad de medida para la variable:")
    for categoria, unidades_lista in unidades.items():
        print(f"{categoria}: {', '.join(unidades_lista)}")

    unidad = input("Ingresa la unidad de medida para la variable "+str(variable)+": ").strip()

    return unidad

def mostrar_graficos_y_tabla(variable, unidad, df, outliers):
    # Calcular estadísticas
    stats = df[variable].describe()

    # Filtrar los datos que no son outliers
    datos_filtrados = df[~df.index.isin(outliers.index)]

    # Crear subgráficas
    fig, axs = plt.subplots(2, 2, figsize=(12, 10))

    # Gráfico de serie de tiempo: Datos filtrados y outliers en diferentes colores
    axs[0, 0].scatter(datos_filtrados['Fecha'], datos_filtrados[variable], color='blue', label='Datos Filtrados', alpha=0.6)
    axs[0, 0].scatter(outliers['Fecha'], outliers[variable], color='red', label='Outliers', alpha=0.6)
    axs[0, 0].set_title('Serie de Tiempo')
    axs[0, 0].set_xlabel('Fecha')
    axs[0, 0].set_ylabel(variable + ' ' + str(unidad))
    axs[0, 0].legend()

    # Gráfico de torta
    porcentajes = [len(datos_filtrados), len(outliers)]
    axs[0, 1].pie(porcentajes, labels=['Datos Filtrados', 'Outliers'], autopct='%1.1f%%')
    axs[0, 1].set_title('Porcentaje de Datos Filtrados vs Outliers')

    # Histograma: Solo los datos filtrados
    axs[1, 0].hist(datos_filtrados[variable], bins=20, color='green', alpha=0.7)
    axs[1, 0].axvline(datos_filtrados[variable].mean(), color='red', linestyle='dashed', linewidth=1, label='Media')
    axs[1, 0].axvline(datos_filtrados[variable].median(), color='blue', linestyle='dashed', linewidth=1, label='Mediana')
    axs[1, 0].set_title('Histograma de Datos Filtrados')
    axs[1, 0].set_xlabel(variable + ' ' + str(unidad))
    axs[1, 0].set_ylabel('Frecuencia')
    axs[1, 0].legend()

    # Tabla de estadísticas
    axs[1, 1].axis('tight')
    axs[1, 1].axis('off')
    table_data = stats.reset_index().values
    axs[1, 1].table(cellText=table_data, colLabels=['Estadística', variable], cellLoc='center', loc='center')

    plt.tight_layout()
    plt.show()

def detectar_outliers_mcd(df, variable, considerar_cero=False, cont=0.1):

    cont = float(input("Ingrese la contaminación para MCD  de la variable "+str(variable)+" (default es 0.1): ").strip())

    # Detectar outliers usando EllipticEnvelope (equivalente a MinCovDet)
    ee = EllipticEnvelope(contamination=cont)
    yhat = ee.fit_predict(df[[variable]])
    
    # yhat tiene valores -1 para outliers y 1 para inliers
    if considerar_cero:
        outliers = df[(yhat == -1) | (df[variable] == 0)]
        datos_filtrados = df[(yhat != -1) & (df[variable] != 0)]  # Excluir filas con valor 0 y los outliers
    else:
        outliers = df[yhat == -1]
        datos_filtrados = df[yhat != -1]

    print(f"Se han identificado {len(outliers)} outliers.")
    
    return outliers, datos_filtrados

def detectar_outliers_knn(df, variable, considerar_cero=False, cont=0.1, n_neighbors=5):

    cont = float(input("Ingrese la contaminación para KNN de la variable "+str(variable)+" (default es 0.1): ").strip())

    # Preparar los datos
    tx = df[[variable]].values  # Convertir a matriz para el modelo

    # Detectar outliers usando KNN de pyod
    knn = KNN(contamination=cont, method='mean', n_neighbors=n_neighbors)
    knn.fit(tx)
    predicted = pd.Series(knn.predict(tx), index=df.index)  # 1 indica outliers, 0 indica inliers
    
    # Incluir el 0 como outlier si el usuario lo ha decidido
    if considerar_cero:
        outliers = df[(predicted == 1) | (df[variable] == 0)]
        datos_filtrados = df[(predicted == 0) & (df[variable] != 0)]  # Excluir filas con valor 0
    else:
        outliers = df[predicted == 1]
        datos_filtrados = df[predicted == 0]

    print(f"Se han identificado {len(outliers)} outliers.")
    
    return outliers, datos_filtrados

def estadistica_univariada(df):
    # Llamar a la función para listar columnas antes de solicitar la variable
    listar_columnas_con_estado(df)

    # Solicitar al usuario el nombre de la variable
    variable = input("Ingresa el nombre de la variable: ").strip()
    if variable not in df.columns:
        print("Nombre de variable no válido. Fin del programa.")
        return

    # Revisar valores NaN en la variable seleccionada
    num_nan = df[variable].isna().sum()
    if num_nan > 0:
        print(f"La variable '{variable}' tiene {num_nan} valores NaN.")

        # Opciones para manejar valores NaN
        while True:
            opcion = input("¿Cómo desea manejar los valores NaN? (eliminar/rrellenar/imputar): ").strip().lower()
            if opcion == 'eliminar':
                df = df[df[variable].notna()]
                print("Se han eliminado las filas con valores NaN.")
                break
            elif opcion == 'rellenar':
                mediana = df[variable].median()
                df[variable].fillna(mediana, inplace=True)
                print(f"Se han rellenado los valores NaN con la mediana: {mediana}.")
                break
            elif opcion == 'imputar':
                imputer = SimpleImputer(strategy='median')
                df.loc[:, variable] = imputer.fit_transform(df[[variable]])
                print("Se han imputado los valores NaN usando la mediana.")
                break
            else:
                print("Opción no válida. Por favor, elija 'eliminar', 'rellenar' o 'imputar'.")

    unidad = unidades_de_medida(variable)

    # Análisis de outliers
    outlier_opcion = input("¿Desea identificar outliers? (KNN/MCD/Ninguno): ").strip().lower()

    # Calcular la mediana de la variable seleccionada
    mediana = df[variable].median()
    
    # Contar cuántas filas tienen valor 0
    count_zero = (df[variable] == 0).sum()
    
    # Mostrar la mediana y la cantidad de ceros
    print(f"La mediana de {variable} es {mediana}.")
    print(f"Número de filas con valor 0: {count_zero}.")
    
    # Preguntar al usuario si desea considerar el 0 como outlier
    considerar_cero = input("¿Desea considerar el 0 como un outlier? (s/n): ").strip().lower()
    if considerar_cero == 's':
        considerar_cero = True
    elif considerar_cero == 'n':
        considerar_cero = False
    else:
        print("Opción no válida. No se considerará el 0 como outlier.")
        considerar_cero = False
    
    if outlier_opcion == 'knn':
        outliers, datos_filtrados = detectar_outliers_knn(df, variable, considerar_cero)
    elif outlier_opcion == 'mcd':
        outliers, datos_filtrados = detectar_outliers_mcd(df, variable, considerar_cero)
    else:
        print("No se identificarán outliers.")

    mostrar_graficos_y_tabla(variable, unidad, datos_filtrados, outliers)


def analisis_multivariado(df, considerar_cero_x=False, considerar_cero_y=False):
    # Listar las columnas disponibles para elegir las variables
    listar_columnas_con_estado(df)
    
    # Solicitar al usuario que elija las variables X y Y
    variable_x = input("Elige la variable para el eje X: ").strip()

    # Revisar valores NaN en la variable seleccionada
    num_nan = df[variable_x].isna().sum()
    if num_nan > 0:
        print(f"La variable '{variable_x}' tiene {num_nan} valores NaN.")

        # Opciones para manejar valores NaN
        while True:
            opcion = input("¿Cómo desea manejar los valores NaN? (eliminar/rrellenar/imputar): ").strip().lower()
            if opcion == 'eliminar':
                df = df[df[variable_x].notna()]
                print("Se han eliminado las filas con valores NaN.")
                break
            elif opcion == 'rellenar':
                mediana = df[variable_x].median()
                df[variable_x].fillna(mediana, inplace=True)
                print(f"Se han rellenado los valores NaN con la mediana: {mediana}.")
                break
            elif opcion == 'imputar':
                imputer = SimpleImputer(strategy='median')
                df.loc[:, variable_x] = imputer.fit_transform(df[[variable_x]])
                print("Se han imputado los valores NaN usando la mediana.")
                break
            else:
                print("Opción no válida. Por favor, elija 'eliminar', 'rellenar' o 'imputar'.")

    variable_y = input("Elige la variable para el eje Y: ").strip()

    # Revisar valores NaN en la variable seleccionada
    num_nan = df[variable_y].isna().sum()
    if num_nan > 0:
        print(f"La variable '{variable_y}' tiene {num_nan} valores NaN.")

        # Opciones para manejar valores NaN
        while True:
            opcion = input("¿Cómo desea manejar los valores NaN? (eliminar/rrellenar/imputar): ").strip().lower()
            if opcion == 'eliminar':
                df = df[df[variable_y].notna()]
                print("Se han eliminado las filas con valores NaN.")
                break
            elif opcion == 'rellenar':
                mediana = df[variable_y].median()
                df[variable_y].fillna(mediana, inplace=True)
                print(f"Se han rellenado los valores NaN con la mediana: {mediana}.")
                break
            elif opcion == 'imputar':
                imputer = SimpleImputer(strategy='median')
                df.loc[:, variable_y] = imputer.fit_transform(df[[variable_y]])
                print("Se han imputado los valores NaN usando la mediana.")
                break
            else:
                print("Opción no válida. Por favor, elija 'eliminar', 'rellenar' o 'imputar'.")
    
    # Solicitar al usuario la unidad de medida para cada variable
    unidad_x = unidades_de_medida(variable_x)
    unidad_y = unidades_de_medida(variable_y)

    # Preguntar al usuario qué método de outliers desea utilizar
    metodo_outliers = input("¿Qué método deseas utilizar para la detección de outliers? (knn/mcd/ninguno): ").lower()
    
    # Inicializar variables para los datos filtrados
    outliers_x = outliers_y = pd.DataFrame()
    datos_filtrados_x = datos_filtrados_y = df
    
    # Aplicar el método de outliers según la elección del usuario
    if metodo_outliers == 'knn':
        outliers_x, datos_filtrados_x = detectar_outliers_knn(df, variable_x, considerar_cero=considerar_cero_x)
        outliers_y, datos_filtrados_y = detectar_outliers_knn(df, variable_y, considerar_cero=considerar_cero_y)
    elif metodo_outliers == 'mcd':
        outliers_x, datos_filtrados_x = detectar_outliers_mcd(df, variable_x, considerar_cero=considerar_cero_x)
        outliers_y, datos_filtrados_y = detectar_outliers_mcd(df, variable_y, considerar_cero=considerar_cero_y)
    
    # Si el método es "ninguno", no filtramos outliers
    if metodo_outliers != 'ninguno':
        # Unir los datos filtrados de ambas variables
        datos_filtrados = df[(df.index.isin(datos_filtrados_x.index)) & (df.index.isin(datos_filtrados_y.index))]
    else:
        datos_filtrados = df
    
    # Crear el layout con gridspec
    fig = plt.figure(figsize=(10, 8))
    grid = fig.add_gridspec(4, 4, hspace=0.5, wspace=0.5)

    # Histograma horizontal de la variable X (en la parte superior)
    ax_histx = fig.add_subplot(grid[0, 0:3])
    sns.histplot(datos_filtrados[variable_x], bins=20, kde=False, color='blue', ax=ax_histx, alpha=0.7)
    ax_histx.axvline(datos_filtrados[variable_x].mean(), color='red', linestyle='dashed', linewidth=1, label='Media')
    ax_histx.axvline(datos_filtrados[variable_x].median(), color='blue', linestyle='dashed', linewidth=1, label='Mediana')
    
    sns.kdeplot(datos_filtrados[variable_x], color='green', ax=ax_histx, label='Tendencia', linewidth=2)
    ax_histx.set_title(f'Histograma de {variable_x} ({unidad_x})')
    ax_histx.set_xlabel('')
    ax_histx.set_ylabel('Frecuencia')
    ax_histx.legend()

    # Histograma vertical de la variable Y (en el lado derecho)
    ax_histy = fig.add_subplot(grid[1:4, 3])
    sns.histplot(datos_filtrados[variable_y], bins=20, kde=False, color='green', ax=ax_histy, alpha=0.7)
    ax_histy.axhline(datos_filtrados[variable_y].mean(), color='red', linestyle='dashed', linewidth=1, label='Media')
    ax_histy.axhline(datos_filtrados[variable_y].median(), color='blue', linestyle='dashed', linewidth=1, label='Mediana')

    sns.kdeplot(datos_filtrados[variable_x], color='green', ax=ax_histx, label='Tendencia', linewidth=2)
    
    ax_histy.set_title(f'Histograma de {variable_y} ({unidad_y})')
    ax_histy.set_xlabel('Frecuencia')
    ax_histy.set_ylabel(variable_y)
    ax_histy.legend()

    # Gráfico de densidad (mapa de calor) en la parte inferior izquierda
    ax_heatmap = fig.add_subplot(grid[1:4, 0:3])
    hb = ax_heatmap.hexbin(datos_filtrados[variable_x], datos_filtrados[variable_y], gridsize=30, cmap='Blues', mincnt=1)
    ax_heatmap.set_title(f'Densidad de {variable_x} y {variable_y}')
    ax_heatmap.set_xlabel(f'{variable_x} ({unidad_x})')
    ax_heatmap.set_ylabel(f'{variable_y} ({unidad_y})')

    plt.colorbar(hb, ax=ax_heatmap, label='Conteo de Datos')
    
    plt.tight_layout()
    plt.show()

def main():
    """Función principal del programa"""
    df_resultado = leer_archivo()
    
    if df_resultado is not None:
        print("Archivo cargado correctamente")
    else:
        print("No se seleccionó ningún archivo para imprimir.")
        exit()

    df_final = preguntar_intervalo_fechas(df_resultado)

    # Mostrar menú de opciones después de cargar el archivo y aplicar filtros
    mostrar_menu(df_final)

    print("Fin del programa.")
    exit()

# Ejecutar el programa
if __name__ == "__main__":
    main()