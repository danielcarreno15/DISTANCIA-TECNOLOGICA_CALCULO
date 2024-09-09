# Cargar paquetes necesarios
install.packages("readxl")
install.packages("writexl")
install.packages("tidyverse")  # Este incluye tidyr y dplyr
install.packages("stringr")
install.packages("openxlsx")

library(readxl)
library(writexl)
library(tidyr)
library(dplyr)
library(stringr)  
library(openxlsx)


# Especifica la ruta del archivo
file.choose()


# Especifica la ruta del archivo
ruta_excel <-"C:\\Users\\danie\\Downloads\\AGRO INNOVATION INTERNATIONAL.xlsx"


# Leer el archivo Excel
data <- read_excel(ruta_excel)

# Verificar que data es un data frame
if (!is.data.frame(data)) {
  stop("El objeto data no es un data frame")
}

# Mostrar los primeros registros para verificar la estructura
print(head(data))

# Verificar los nombres de las columnas
print(colnames(data))

# Asegurarte de que la columna Application Date existe
if (!"Application Date" %in% colnames(data)) {
  stop("La columna 'Application Date' no existe en el data frame")
}

# Convertir Application Date a fecha
data <- data %>%
  mutate(Application_Date = as.Date(`Application Date`, format = "%d.%m.%Y"))

# Función para filtrar por rango de años y contar patentes
filtrar_y_contar <- function(data, start_year, end_year) {
  data_filtrada <- data %>%
    filter(format(Application_Date, "%Y") >= start_year & format(Application_Date, "%Y") <= end_year)
  
  conteo_por_ano <- data_filtrada %>%
    group_by(ano = format(Application_Date, "%Y")) %>%
    summarize(conteo = n())
  
  total_patentes <- nrow(data_filtrada)
  rango <- paste(start_year, end_year, sep = "-")
  
  list(data = data_filtrada, conteo_por_ano = conteo_por_ano, total_patentes = total_patentes, rango = rango)
}

# Filtrar por los rangos especificados
rango_1 <- filtrar_y_contar(data, 2012, 2016)
rango_2 <- filtrar_y_contar(data, 2013, 2017)
rango_3 <- filtrar_y_contar(data, 2014, 2018)


#CÓMO HALLAR LA DISTANCIA TECNOLÓGICA

# Crear un nuevo data frame con los resultados resumidos
df_rango_1 <- rango_1$data
df_rango_2 <- rango_2$data
df_rango_3 <- rango_3$data

# Comprobar si la columna "Application Id" existe
if(!"Application Id" %in% names(df_rango_1)) {
  stop("La columna 'Application Id' no existe en df_rango_1")
}
if(!"Application Id" %in% names(df_rango_2)) {
  stop("La columna 'Application Id' no existe en df_rango_2")
}
if(!"Application Id" %in% names(df_rango_3)) {
  stop("La columna 'Application Id' no existe en df_rango_3")
}

# Función para separar los códigos IPC en columnas y mantener las otras columnas
separar_codigos <- function(df) {
  # Separar los códigos IPC en columnas
  df_codigos <- df %>%
    separate_rows(`I P C`, sep = ";") %>%
    mutate(`I P C` = str_trim(`I P C`))
  
  # Extraer los primeros tres dígitos de cada código IPC
  df_codigos <- df_codigos %>%
    mutate(Codigo = substr(`I P C`, 1, 3))
  
  # Combinar los códigos IPC separados en una única columna para cada "Application Id"
  df_codigos_unidos <- df_codigos %>%
    group_by(`Application Id`) %>%
    summarise(Codigos_Unicos = paste(unique(Codigo), collapse = ";")) %>%
    ungroup()
  
  # Combinar el DataFrame original con los códigos separados
  df_final <- df %>%
    select(-`I P C`) %>%
    left_join(df_codigos_unidos, by = "Application Id")
  
  return(df_final)
}

# Aplicar la función a los DataFrames
df_rango_1_niveles <- separar_codigos(df_rango_1)
df_rango_2_niveles <- separar_codigos(df_rango_2)
df_rango_3_niveles <- separar_codigos(df_rango_3)




# Función para crear la matriz de comparación con columna de sumatoria
crear_matriz_comparacion <- function(df_niveles, codigos_objetivo) {
  # Obtener los nombres de las columnas que contienen los códigos IPC
  columnas_codigos <- grep("Codigos_Unicos", names(df_niveles), value = TRUE)
  
  # Obtener los Application Ids
  application_ids <- df_niveles$`Application Id`
  
  # Crear una matriz vacía con las dimensiones adecuadas
  matriz <- matrix(0, nrow = length(codigos_objetivo), ncol = nrow(df_niveles))
  
  # Rellenar la matriz con 1 si hay coincidencia de códigos IPC
  for (i in 1:nrow(df_niveles)) {
    codigos_presentes <- str_split(df_niveles[i, columnas_codigos], ";") %>% unlist() %>% na.omit()
    for (j in 1:length(codigos_objetivo)) {
      if (codigos_objetivo[j] %in% codigos_presentes) {
        matriz[j, i] <- 1
      }
    }
  }
  
  # Convertir la matriz en un data frame
  df_matriz <- as.data.frame(matriz)
  
  # Añadir los nombres de columna (Application Ids)
  colnames(df_matriz) <- application_ids
  
  # Añadir los códigos como una columna al data frame
  df_matriz <- cbind(Codigo = codigos_objetivo, df_matriz)
  
  # Verificar si df_matriz tiene más de una columna después de excluir la columna "Codigo"
  if (ncol(df_matriz) > 2) {  # Si tiene más de dos columnas (incluyendo "Codigo" y al menos una columna de datos)
    # Añadir una columna con la sumatoria por fila
    df_matriz$Total <- rowSums(df_matriz[, -1])  # Excluir la columna "Codigo" en el cálculo de sumas
  } else {
    # Si solo hay una columna de datos, la suma de esa fila es simplemente el valor de esa columna
    df_matriz$Total <- df_matriz[, 2]
  }
  
  return(df_matriz)
}

# Definir los códigos objetivo
codigos_objetivo <- c(
  "A01", "A21", "A22", "A23", "A24", "A41", "A42", "A43", "A44",  
  "A45", "A46", "A47", "A61", "A62", "A63", "A99", "B01", "B02", "B03", "B04",  
  "B05", "B06", "B07", "B08", "B09", "B21", "B22", "B23", "B24", "B25", "B26",  
  "B27", "B28", "B29", "B30", "B31", "B32", "B33", "B41", "B42", "B43",  
  "B44", "B60", "B61", "B62", "B63", "B64", "B65", "B66", "B67", "B68", "B81",  
  "B82", "B99", "C01", "C02", "C03", "C04", "C05", "C06", "C07", "C08", "C09",  
  "C10", "C11", "C12", "C13", "C14", "C21", "C22", "C23", "C25", "C30", "C40",  
  "C99", "D01", "D02", "D03", "D04", "D05", "D06", "D07", "D21", "D99", "E01",  
  "E02", "E03", "E04", "E05", "E06", "E21", "E99", "F01", "F02", "F03", "F04",  
  "F15", "F16", "F17", "F21", "F22", "F23", "F24", "F25", "F26", "F27", "F28",  
  "F41", "F42", "F99", "G01", "G02", "G03", "G04", "G05", "G06", "G07", "G08",  
  "G09", "G10", "G11", "G12", "G16", "G21", "G99", "H01", "H02", "H03", "H04",  
  "H05", "H10", "H99") 



# Crear las matrices de comparación
matriz_rango_1 <- crear_matriz_comparacion(df_rango_1_niveles, codigos_objetivo)
matriz_rango_2 <- crear_matriz_comparacion(df_rango_2_niveles, codigos_objetivo)
matriz_rango_3 <- crear_matriz_comparacion(df_rango_3_niveles, codigos_objetivo)

# Crear un nuevo archivo Excel y agregar hojas con los resultados
wb <- createWorkbook()

agregar_hoja <- function(wb, rango, nombre_hoja) {
  addWorksheet(wb, nombre_hoja)
  if (!is.null(rango$data) && nrow(rango$data) > 0) {
    writeData(wb, nombre_hoja, rango$data, startCol = 1, startRow = 1)
    writeData(wb, nombre_hoja, paste("Total patentes en el rango", rango$rango, ":", rango$total_patentes), startCol = 1, startRow = nrow(rango$data) + 2)
    if (!is.null(rango$conteo_por_ano) && nrow(rango$conteo_por_ano) > 0) {
      writeData(wb, nombre_hoja, rango$conteo_por_ano, startCol = 1, startRow = nrow(rango$data) + 4)
    }
  } else {
    writeData(wb, nombre_hoja, "No hay datos disponibles", startCol = 1, startRow = 1)
  }
}

# Agregar hojas con los datos filtrados
agregar_hoja(wb, rango_1, "2012-2016")
agregar_hoja(wb, rango_2, "2013-2017")
agregar_hoja(wb, rango_3, "2014-2018")


# Añadir cada matriz como una hoja separada
addWorksheet(wb, "Matriz_Rango_1")
writeData(wb, "Matriz_Rango_1", matriz_rango_1)

addWorksheet(wb, "Matriz_Rango_2")
writeData(wb, "Matriz_Rango_2", matriz_rango_2)

addWorksheet(wb, "Matriz_Rango_3")
writeData(wb, "Matriz_Rango_3", matriz_rango_3)

# Guardar el archivo excel
saveWorkbook(wb, "C:\\Users\\danie\\Documents\\EMPRESAS DT\\EMPRESA_EJEMPLO\\Nueva.xlsx", overwrite = TRUE)



# Función para procesar cada archivo Excel para el almacenamiento
procesar_y_almacenar_excel <- function(ruta_excel, ruta_almacenamiento) {
  # Leer el archivo Excel generado
  data_matriz_1 <- read_excel(ruta_excel, sheet = "Matriz_Rango_1")
  data_matriz_2 <- read_excel(ruta_excel, sheet = "Matriz_Rango_2")
  data_matriz_3 <- read_excel(ruta_excel, sheet = "Matriz_Rango_3")
  
  # Verificar que data_matriz son data frames
  if (!is.data.frame(data_matriz_1) | !is.data.frame(data_matriz_2) | !is.data.frame(data_matriz_3)) {
    stop("Alguno de los objetos data no es un data frame")
  }
  
  # Extraer los códigos IPC de las hojas matriz
  codigos_ipc_1 <- data_matriz_1$Codigo
  codigos_ipc_2 <- data_matriz_2$Codigo
  codigos_ipc_3 <- data_matriz_3$Codigo
  
  # Extraer la columna "Total" de las hojas matriz
  total_1 <- data_matriz_1$Total
  total_2 <- data_matriz_2$Total
  total_3 <- data_matriz_3$Total
  
  # Crear un nuevo data frame con los códigos IPC como columnas
  df_resultado_1 <- tibble(
    Empresa = tools::file_path_sans_ext(basename(ruta_excel)),
    !!!setNames(as.list(total_1), codigos_ipc_1)
  )
  
  df_resultado_2 <- tibble(
    Empresa = tools::file_path_sans_ext(basename(ruta_excel)),
    !!!setNames(as.list(total_2), codigos_ipc_2)
  )
  
  df_resultado_3 <- tibble(
    Empresa = tools::file_path_sans_ext(basename(ruta_excel)),
    !!!setNames(as.list(total_3), codigos_ipc_3)
  )
  
  # Leer el contenido actual del archivo de almacenamiento
  if (file.exists(ruta_almacenamiento)) {
    wb_almacenamiento <- loadWorkbook(ruta_almacenamiento)
  } else {
    # Si el archivo no existe, crearlo y agregar las hojas necesarias
    wb_almacenamiento <- createWorkbook()
    addWorksheet(wb_almacenamiento, "Almacenamiento_Rango_1")
    addWorksheet(wb_almacenamiento, "Almacenamiento_Rango_2")
    addWorksheet(wb_almacenamiento, "Almacenamiento_Rango_3")
  }
  
  # Escribir los datos en cada hoja (verificando si ya existen)
  if ("Almacenamiento_Rango_1" %in% names(wb_almacenamiento)) {
    current_data_1 <- read.xlsx(wb_almacenamiento, sheet = "Almacenamiento_Rango_1")
    writeData(wb_almacenamiento, "Almacenamiento_Rango_1", 
              rbind(current_data_1, df_resultado_1), colNames = TRUE)
  }
  
  if ("Almacenamiento_Rango_2" %in% names(wb_almacenamiento)) {
    current_data_2 <- read.xlsx(wb_almacenamiento, sheet = "Almacenamiento_Rango_2")
    writeData(wb_almacenamiento, "Almacenamiento_Rango_2", 
              rbind(current_data_2, df_resultado_2), colNames = TRUE)
  }
  
  if ("Almacenamiento_Rango_3" %in% names(wb_almacenamiento)) {
    current_data_3 <- read.xlsx(wb_almacenamiento, sheet = "Almacenamiento_Rango_3")
    writeData(wb_almacenamiento, "Almacenamiento_Rango_3", 
              rbind(current_data_3, df_resultado_3), colNames = TRUE)
  }
  
  # Guardar el archivo Excel de almacenamiento
  saveWorkbook(wb_almacenamiento, ruta_almacenamiento, overwrite = TRUE)
}

# Ruta del archivo Excel de almacenamiento
ruta_almacenamiento <- "C:\\Users\\danie\\Documents\\EMPRESAS DT\\EMPRESA_EJEMPLO\\ALMACENAMIENTO_DT.xlsx"

# Procesar y almacenar la información del archivo generado
procesar_y_almacenar_excel("C:\\Users\\danie\\Documents\\EMPRESAS DT\\EMPRESA_EJEMPLO\\nueva.xlsx", ruta_almacenamiento)

# Obtener el directorio de trabajo actual
print(getwd())
