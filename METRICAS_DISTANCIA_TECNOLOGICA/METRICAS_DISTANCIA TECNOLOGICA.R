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
ruta_excel <-"C:\\Users\\danie\\Downloads\\PERFILES_TECNOLOGICOS_EMPRESAS.xlsx"

# Leer los datos de cada hoja
perfiles1 <- read_excel(ruta_excel, sheet = "Almacenamiento_Rango_1")
perfiles2 <- read_excel(ruta_excel, sheet = "Almacenamiento_Rango_2")
perfiles3 <- read_excel(ruta_excel, sheet = "Almacenamiento_Rango_3")

# Verificar si hay NA o Inf en los datos
verificar_datos_validos <- function(df) {
  return(!any(is.na(df)) && all(is.finite(df)))
}

# Normalizar los datos por fila (proporciones)
normalizar_datos <- function(df) {
  return(as.data.frame(t(apply(df[, -1], 1, function(x) x / sum(x, na.rm = TRUE)))))
}

perfiles1_normalizado <- cbind(perfiles1[1], normalizar_datos(perfiles1))
perfiles2_normalizado <- cbind(perfiles2[1], normalizar_datos(perfiles2))
perfiles3_normalizado <- cbind(perfiles3[1], normalizar_datos(perfiles3))

# Función para calcular la distancia de Jaffe (ajustada)
distancia_jaffe <- function(p1, p2) {
  return(1 - (sum(p1 * p2) / sqrt(sum(p1^2) * sum(p2^2))))
}

# Función para calcular la distancia euclidiana
distancia_euclidiana <- function(p1, p2) {
  return(sqrt(sum((p1 - p2)^2)))
}

# Función para calcular la distancia de mínimo complemento
distancia_minimo_complemento <- function(p1, p2) {
  suma_pmin <- sum(pmin(p1, p2))
  return(1 - suma_pmin)
}

# Función para calcular distancias entre dos empresas específicas en una hoja
calcular_distancias_especificas <- function(perfiles, empresa1, empresa2) {
  nombres_empresas <- perfiles$Empresa
  fila1 <- which(nombres_empresas == empresa1)
  fila2 <- which(nombres_empresas == empresa2)
  
  if(length(fila1) == 0 || length(fila2) == 0) {
    stop("Una o ambas empresas no se encuentran en los datos.")
  }
  
  p1 <- as.numeric(perfiles[fila1, -1])
  p2 <- as.numeric(perfiles[fila2, -1])
  
  dist_jaffe <- distancia_jaffe(p1, p2)
  dist_euclidiana <- distancia_euclidiana(p1, p2)
  dist_minimo_complemento <- distancia_minimo_complemento(p1, p2)
  
  return(list(jaffe = dist_jaffe, euclidiana = dist_euclidiana, minimo_complemento = dist_minimo_complemento))
}

# Calcular las distancias entre las dos empresas específicas para cada perfil
empresa1 <- "NATIONAL UNIVERSITY CORPORATION THE UNIVERSITY OF TOKYO"
empresa2 <- "TOHOKU UNIVERSITY"

distancias1 <- calcular_distancias_especificas(perfiles1_normalizado, empresa1, empresa2)
distancias2 <- calcular_distancias_especificas(perfiles2_normalizado, empresa1, empresa2)
distancias3 <- calcular_distancias_especificas(perfiles3_normalizado, empresa1, empresa2)

# Mostrar los resultados
print(paste("Distancias en almacenamiento_rango_1:"))
print(distancias1)
print(paste("Distancias en almacenamiento_rango_2:"))
print(distancias2)
print(paste("Distancias en almacenamiento_rango_3:"))
print(distancias3)

# Guardar los resultados en archivos Excel
resultados <- data.frame(
  Perfil = c("almacenamiento_rango_1", "almacenamiento_rango_2", "almacenamiento_rango_3"),
  Jaffe = c(distancias1$jaffe, distancias2$jaffe, distancias3$jaffe),
  Euclidiana = c(distancias1$euclidiana, distancias2$euclidiana, distancias3$euclidiana),
  Minimo_Complemento = c(distancias1$minimo_complemento, distancias2$minimo_complemento, distancias3$minimo_complemento)
)

write_xlsx(list("Resultados" = resultados), "DT_235.3.xlsx")

