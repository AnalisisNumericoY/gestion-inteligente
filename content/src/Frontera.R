# 0) Cargar datos y preparar columnas PYP
suppressPackageStartupMessages({
  library(data.table)
  library(openxlsx)
  library(dplyr)
  library(lightgbm)
})

input <- "./input"
output <- "./output"
src <- "./src"

# Cargar utilidades
source(here::here("src", "frontera_utils.R"), encoding = "UTF-8")

# Cargas tus objetos ya preparados
load(file.path(output,"modelo_30_oct_2025.RData"))
load(file.path(output,"08_data_maestro_sin_outliers.RData"))

data_maestro_sin_outliers <- as.data.table(data_maestro_sin_outliers)
actividades <- read.xlsx(file.path(input,"catalogo_actividades.xlsx"))
actividades <- as.data.table(actividades)

# Lectura Excel----
perfil_usuario <- read.xlsx(file.path(input,"Input_frontera.xlsx"))
perfil_usuario <- as.data.table(perfil_usuario)

# limpiar lo que viene del Excel
perfil_usuario[, `:=`(
  Sector_Econom = factor(trimws(toupper(Sector_Econom)),
                         levels = levels(data_maestro_sin_outliers$Sector_Econom)),
  Tamano_Emp    = factor(trimws(Tamano_Emp),
                         levels = levels(data_maestro_sin_outliers$Tamano_Emp)),
  Activ_Econ    = factor(trimws(Activ_Econ),
                         levels = levels(data_maestro_sin_outliers$Activ_Econ)),
  Sucursal      = factor(trimws(Sucursal),
                         levels = levels(data_maestro_sin_outliers$Sucursal)),
  Num_Empleados = as.numeric(trimws(Num_Empleados)),
  tasa_deseada  = as.numeric(trimws(tasa_deseada))
)]

# Asignación directa a variables de usuario
sector_usr    <- perfil_usuario$Sector_Econom
tamano_usr    <- perfil_usuario$Tamano_Emp
sucursal_usr  <- perfil_usuario$Sucursal 
activ_usr     <- perfil_usuario$Activ_Econ
num_emp_usr   <- perfil_usuario$Num_Empleados
tasa_deseada  <- perfil_usuario$tasa_deseada

# creación de variable target del modelo
conteo_target <- (tasa_deseada * num_emp_usr) / 100

# Detectar columnas 
pyp_cols <- grep("^PYP_AR.*_prop$", names(data_maestro_sin_outliers), value = TRUE)
cat_cols <- c("Sector_Econom", "Tamano_Emp", "Sucursal","Activ_Econ")
extra_vars <- c("Num_Empleados","tasa_100_mod")
x_cols <- c(pyp_cols, cat_cols, extra_vars)

# 3) Calcular actividades permitidas vs quemadas con fallbacks
sets_pyp <- get_pyp_activity_sets(
  df         = data_maestro_sin_outliers,
  sector     = sector_usr,
  tamano     = tamano_usr,
  sucursal   = sucursal_usr,
  activ_econ = activ_usr,
  pyp_cols   = pyp_cols
)

# 4) Inspeccionar resultados
nivel_usado <- attr(sets_pyp, "nivel_usado")
cat("Nivel usado para este perfil:", nivel_usado, "\n")

cat("Actividades PYP a tener en cuenta:\n")
invisible(lapply(seq_along(sets_pyp$allowed), function(i) {
  cat(sprintf("%2d. %s\n", i, sets_pyp$allowed[i]))
}))

# 5) Histórico usado
hist_usado <- get_historial_por_nivel(
  df         = data_maestro_sin_outliers,
  sector     = sector_usr,
  tamano     = tamano_usr,
  activ_econ = activ_usr,
  sucursal   = sucursal_usr,
  nivel_usado = nivel_usado
)

# 6) Promedios PYP y normalización (como en el original)
pyp_means <- hist_usado[, lapply(.SD, mean, na.rm = TRUE), .SDcols = sets_pyp$allowed]
sum(pyp_means)
pyp_promedios <- as.list(pyp_means[1, ])

# 7) Corre la frontera para el perfil
resultados <- list(
  correr_frontera_para_perfil(
    perfil_row                = perfil_usuario,
    actividades               = actividades,
    data_maestro_sin_outliers = data_maestro_sin_outliers,
    pyp_cols                  = pyp_cols,
    x_cols                    = x_cols,
    model                     = model
  )
)

# 8) Salidas
exportar_recomendacion_pyp (resultados,nivel_usado,idx   = 1L,file  = "Recomendacion_PYP.xlsx",sheet = "Recomendacion_PYP")
