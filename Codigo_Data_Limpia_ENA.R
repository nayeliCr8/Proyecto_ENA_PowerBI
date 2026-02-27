# Cargar librerías
library(readxl)
library(dplyr)
library(openxlsx) 
library(mice)

data <- read_excel("E:/UNIDAD D/FINESI - SEMESTRES/MAESTRÍA/II CICLO/BUSINESS INTELIGENT Y ANALISIS DE DATOS/SESIÓN 5/DATOS_EXCEL/2014-2019/Cap100_2014_2019_Prelimpio.xlsx")

# Ver porcentaje de valores faltantes por columna
na_percent <- sapply(data, function(x) mean(is.na(x)) * 100)
sort(na_percent, decreasing = TRUE)

# ---- Eliminar columnas con más del 90% de NA ----
data_limpia <- data %>%
  select(-c(P103_EE, ESTRATO, P101A, TipoMuestra))

# ---- Convertir variables categóricas a factor ----
data_limpia <- data_limpia %>%
  mutate(
    RESFIN = factor(RESFIN),       # 1 = Completa, 2 = Incompleta
    REGION = factor(REGION),       # 1 = Costa, 2 = Sierra, 3 = Selva
    DOMINIO = factor(DOMINIO),     # 1-7
    CODIGO = factor(CODIGO),       # 1-2
    P15 = factor(P15),             # 1-4
    P410A = factor(P410A),         # 1-2
    P418 = factor(P418)            # 1-2
  )

# ---- Configurar métodos de imputación ----
# Predictor matrix y métodos
ini <- mice(data_limpia, maxit = 0)
meth <- ini$method

# Definir métodos por tipo de variable
meth["RESFIN"] <- "logreg"
meth["REGION"] <- "polyreg"
meth["DOMINIO"] <- "polyreg"
meth["CODIGO"] <- "logreg"
meth["P15"] <- "polyreg"
meth["P410A"] <- "logreg"
meth["P418"] <- "logreg"

# Las demás (numéricas) quedan como "pmm"

# ---- Imputar ----
imp <- mice(data_limpia, m = 1, method = meth, maxit = 5, seed = 123)

# Extraer el dataset imputado
data_imputada <- complete(imp)

# Verificar NAs
colSums(is.na(data_imputada))

# Exportar a un Excel nuevo
write.xlsx(data_imputada,
           file = "E:/UNIDAD D/FINESI - SEMESTRES/MAESTRÍA/II CICLO/BUSINESS INTELIGENT Y ANALISIS DE DATOS/SESIÓN 5/DATOS_EXCEL/2014-2019/datos_limpios.xlsx")

cat("Archivo exportado como 'datos_limpios.xlsx'\n")
