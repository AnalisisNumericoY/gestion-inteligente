set.seed(as.integer(format(Sys.Date(), "%Y%m%d")))


# ---- Funciones auxiliares de Frontera ----

.compute_pyp_sets <- function(df_hist, pyp_cols, tol = 1e-6) {
  if (!nrow(df_hist)) {
    return(NULL)
  }
  zero_all <- vapply(
    df_hist[, ..pyp_cols],
    function(x) all(is.na(x) | abs(x) <= tol),
    logical(1)
  )
  forced_zero <- names(zero_all)[zero_all]
  allowed     <- setdiff(pyp_cols, forced_zero)
  if (!length(allowed)) {
    return(NULL)
  }
  list(allowed = allowed, forced_zero = forced_zero)
}

get_pyp_activity_sets <- function(df,
                                  sector,
                                  tamano,
                                  activ_econ,
                                  pyp_cols,
                                  sucursal = NULL) 
{
  if (!is.null(sucursal)) {
    df0 <- df[
      Sector_Econom == sector &
        Tamano_Emp  == tamano &
        Activ_Econ  == activ_econ &
        Sucursal    == sucursal
    ]
    sets0 <- .compute_pyp_sets(df0, pyp_cols = pyp_cols)
    if (!is.null(sets0)) {
      attr(sets0, "nivel_usado") <- "sector+tamano+Activ_Econ+Sucursal"
      return(sets0)
    }
  }
  
  df1 <- df[
    Sector_Econom == sector &
      Tamano_Emp  == tamano &
      Activ_Econ  == activ_econ
  ]
  sets1 <- .compute_pyp_sets(df1, pyp_cols = pyp_cols)
  if (!is.null(sets1)) {
    attr(sets1, "nivel_usado") <- "sector+tamano+Activ_Econ"
    return(sets1)
  }
  
  df2 <- df[
    Sector_Econom == sector &
      Activ_Econ  == activ_econ
  ]
  sets2 <- .compute_pyp_sets(df2, pyp_cols = pyp_cols)
  if (!is.null(sets2)) {
    attr(sets2, "nivel_usado") <- "sector+Activ_Econ"
    return(sets2)
  }
  
  df3 <- df[
    Sector_Econom == sector
  ]
  sets3 <- .compute_pyp_sets(df3, pyp_cols = pyp_cols)
  if (!is.null(sets3)) {
    attr(sets3, "nivel_usado") <- "solo sector"
    return(sets3)
  }
  
  stop("No se encontró ningún histórico con PYP no-cero para ese sector (ni en niveles agregados).")
}

get_historial_por_nivel <- function(df,
                                    sector,
                                    tamano,
                                    activ_econ,
                                    nivel_usado,
                                    sucursal = NULL) 
{
  if (nivel_usado == "sector+tamano+Activ_Econ+Sucursal") {
    df[
      Sector_Econom == sector &
        Tamano_Emp  == tamano &
        Activ_Econ  == activ_econ &
        Sucursal    == sucursal
    ]
  } else if (nivel_usado == "sector+tamano+Activ_Econ") {
    df[
      Sector_Econom == sector &
        Tamano_Emp  == tamano &
        Activ_Econ  == activ_econ
    ]
  } else if (nivel_usado == "sector+Activ_Econ") {
    df[
      Sector_Econom == sector &
        Activ_Econ  == activ_econ
    ]
  } else if (nivel_usado == "solo sector") {
    df[
      Sector_Econom == sector
    ]
  } else {
    stop("nivel_usado desconocido: ", nivel_usado)
  }
}


elegir_empresa_base_por_tasa <- function(df_hist, tasa_deseada) {
  df_valid <- df_hist[!is.na(tasa_100_mod)]
  if (!nrow(df_valid)) {
    stop("No hay registros con tasa_100_mod válida en el histórico para este perfil.")
  }
  df_valid[, dist := abs(tasa_100_mod - tasa_deseada)]
  idx <- which.min(df_valid$dist)
  base_row <- copy(df_valid[idx])
  base_row[, tasa_100_mod := tasa_deseada]
  base_row[, dist := NULL]
  base_row[]
}

crear_intervalos_pyp <- function(pyp_base_norm,
                                 rel_width = 0.30,
                                 min_width = 2,
                                 max_abs   = 100) {
  stopifnot(is.numeric(pyp_base_norm))
  m <- pyp_base_norm
  ancho <- pmax(rel_width * m, min_width)
  lower <- pmax(0,   m - ancho)
  upper <- pmin(max_abs, m + ancho)
  data.table::data.table(
    actividad = names(m),
    mean      = as.numeric(m),
    min       = as.numeric(lower),
    max       = as.numeric(upper)
  )
}

generar_plan_candidato <- function(intervalos_pyp, step = NULL) {
  n <- nrow(intervalos_pyp)
  vals <- runif(n, min = intervalos_pyp$min, max = intervalos_pyp$max)
  if (!is.null(step) && step > 0) {
    vals <- round(vals / step) * step
  }
  if (sum(vals, na.rm = TRUE) <= 0) {
    idx <- which.max(intervalos_pyp$mean)
    vals[idx] <- max(step %||% 1, 1)
  }
  vals <- (100 * vals) / sum(vals, na.rm = TRUE)
  setNames(vals, intervalos_pyp$actividad)
}

`%||%` <- function(a, b) if (!is.null(a)) a else b

verificar_intervalo_simetrico <- function(valor, centro, amplitud, expand = 0) {
  amplitud_adj <- amplitud * (1 + expand)
  lower <- centro - amplitud_adj
  upper <- centro + amplitud_adj
  dentro <- valor >= lower & valor <= upper
  list(
    valor = valor,
    centro = centro,
    amplitud_original = amplitud,
    amplitud_ajustada = amplitud_adj,
    limite_inferior = lower,
    limite_superior = upper,
    dentro = dentro
  )
}

optimizar_planes_pyp_objetivo <- function(base_row,
                                          sets_pyp,
                                          pyp_base_norm,
                                          model,
                                          x_cols,
                                          conteo_target,
                                          n_iter = 5000L,
                                          step   = NULL,
                                          guardar_historial = FALSE) {
  stopifnot(nrow(base_row) == 1)
  stopifnot(is.finite(conteo_target), conteo_target > 0)
  allowed     <- sets_pyp$allowed
  forced_zero <- sets_pyp$forced_zero
  rel_width_vec  <- c(0.15, 0.20)
  target_tol_vec <- c(0.10, 0.30)
  encontrado <- FALSE
  if (guardar_historial) {
    hist_planes <- data.table::data.table(
      iter_global = integer(0),
      rel_width   = numeric(0),
      tol_rel     = numeric(0),
      pred        = numeric(0),
      diff_target = numeric(0)
    )
  } else {
    hist_planes <- NULL
  }
  iter_global <- 0L
  for (rel_w in rel_width_vec) {
    intervalos_pyp <- crear_intervalos_pyp(
      pyp_base_norm,
      rel_width = rel_w,
      min_width = 2,
      max_abs   = 100
    )[actividad %in% allowed]
    if (!nrow(intervalos_pyp)) next
    for (tol_rel in target_tol_vec) {
      for (i in seq_len(n_iter)) {
        iter_global <- iter_global + 1L
        print(i)
        plan_i <- generar_plan_candidato(intervalos_pyp, step = step)
        row_i <- data.table::copy(base_row)
        if (length(forced_zero)) {
          row_i[, (forced_zero) := 0]
        }
        row_i[, (names(plan_i)) := as.list(plan_i)]
        X  <- data.matrix(row_i[, ..x_cols])
        mu <- as.numeric(predict(model, X))
        diff_mu <- abs(mu - conteo_target)
        if (!is.null(hist_planes)) {
          hist_planes <- rbind(
            hist_planes,
            data.table::data.table(
              iter_global = iter_global,
              rel_width   = rel_w,
              tol_rel     = tol_rel,
              pred        = mu,
              diff_target = diff_mu
            ),
            use.names = TRUE
          )
        }
        res_int <- verificar_intervalo_simetrico(
          valor    = mu,
          centro   = conteo_target,
          amplitud = conteo_target * tol_rel,
          expand   = 0
        )
        if (res_int$dentro) {
          encontrado <- TRUE
          tasa_pred <- (mu * 100) / base_row$Num_Empleados
          tasa_objetivo <- (conteo_target * 100) / base_row$Num_Empleados
          return(list(
            best_pred_tasa = tasa_pred,
            tasa_objetivo  = tasa_objetivo,
            best_plan      = plan_i,
            diff_tasa      = abs(tasa_pred - tasa_objetivo),
            meta           = list(rel_width = rel_w, tol_rel = tol_rel),
            historial_pred = hist_planes
          ))
        }
      }
    }
  }
  stop(
    sprintf(
      "No se encontró ningún plan PYP cuyo conteo predicho caiga dentro de las tolerancias (±10%%, luego ±30%%) alrededor del conteo objetivo = %.3f. 
Revise la tasa objetivo o las condiciones de búsqueda.",
      conteo_target
    )
  )
}

mapear_plan_actividades <- function(best_plan, actividades) {
  dt <- data.table(
    pyp_col    = names(best_plan),
    porcentaje = as.numeric(best_plan)
  )[porcentaje > 0]
  dt[, codigo_actividad := sub("^PYP_(AR[0-9]{4})_prop$", "\\1", pyp_col)]
  dt <- merge(
    dt,
    actividades,
    by = "codigo_actividad",
    all.x = TRUE
  )
  dt[order(-porcentaje),
     .(
       codigo_actividad,
       ACTIVIDAD,
       porcentaje_recomendado = round(porcentaje, 2)
     )]
}

correr_frontera_para_perfil <- function(perfil_row,
                                        actividades,
                                        data_maestro_sin_outliers,
                                        pyp_cols,
                                        x_cols,
                                        model) {
  sets_pyp <- get_pyp_activity_sets(
    df         = data_maestro_sin_outliers,
    sector     = sector_usr,
    tamano     = tamano_usr,
    activ_econ = activ_usr,
    sucursal   = sucursal_usr,
    pyp_cols   = pyp_cols
  )
  nivel_usado <- attr(sets_pyp, "nivel_usado")
  hist_usado <- get_historial_por_nivel(
    df          = data_maestro_sin_outliers,
    sector      = sector_usr,
    tamano      = tamano_usr,
    activ_econ  = activ_usr,
    sucursal    = sucursal_usr,
    nivel_usado = nivel_usado
  )
  pyp_means <- hist_usado[, lapply(.SD, mean, na.rm = TRUE),
                          .SDcols = sets_pyp$allowed]
  pyp_vec      <- unlist(pyp_means[1, ], use.names = TRUE)
  suma_pyp     <- sum(pyp_vec, na.rm = TRUE)
  if (suma_pyp <= 0) {
    stop("La suma de los promedios de PYP es cero; no se puede normalizar.")
  }
  pyp_base_norm <- (100 * pyp_vec) / suma_pyp
  base_row <- elegir_empresa_base_por_tasa(
    df_hist      = hist_usado,
    tasa_deseada = tasa_deseada
  )
  res_opt <- optimizar_planes_pyp_objetivo(
    base_row       = base_row,
    sets_pyp       = sets_pyp,
    pyp_base_norm  = pyp_base_norm,
    model          = model,
    x_cols         = x_cols,
    conteo_target  = conteo_target,
    n_iter         = 10000,
    step           = 1,
    guardar_historial = TRUE
  )
  tabla_pyp <- mapear_plan_actividades(
    best_plan      = res_opt$best_plan,
    actividades = actividades
  )
  meta <- res_opt$meta
  resumen <- data.table(
    Sector_Econom    = sector_usr,
    Tamano_Emp       = tamano_usr,
    Activ_Econ       = as.character(perfil_row$Activ_Econ),
    Num_Empleados    = num_emp_usr,
    Sucursal         = sucursal_usr,
    tasa_objetivo    = res_opt$tasa_objetivo,
    tasa_predicha    = res_opt$best_pred_tasa,
    diferencia_tasa  = res_opt$diff_tasa,
    nivel_hist_usado = nivel_usado,
    rel_width        = meta$rel_width,
    tol_rel          = meta$tol_rel
  )
  list(
    resumen_perfil   = resumen,
    recomendacion_pyp = tabla_pyp,
    historial_pred    = res_opt$historial_pred
  )
}

exportar_recomendacion_pyp <- function(resultados,
                                       nivel_usado,
                                       idx   = 1L,
                                       file  = "Recomendacion_PYP.xlsx",
                                       sheet = "Recomendacion_PYP") {
  # Se asume que ya cargaste:
  # library(openxlsx)
  # library(data.table)
  # library(stringi)
  
  res_i <- resultados[[idx]]
  
  # Data de recomendación
  df <- as.data.table(res_i$recomendacion_pyp)
  
  # Extraer resumen del perfil
  tol_rel         <- round(100 * res_i$resumen_perfil$tol_rel, 2)
  diferencia_tasa <- res_i$resumen_perfil$diferencia_tasa

  # Texto del encabezado / footer (con Unicode y %)
  fmt_header <- paste0(
    "La estimaci\u00f3n se encontr\u00f3 con un error de estimaci\u00f3n del %.2f%%.\n",
    "La diferencia con la tasa deseada es de %.2f.\n",
    "Para dar la recomendaci\u00f3n se usaron las empresas que cumplen con los criterios ",
    "deseados de %s."
  )
  
  encabezado <- sprintf(
    fmt_header,
    tol_rel,
    diferencia_tasa,
    nivel_usado
  )
  
  # Crear workbook y hoja
  wb <- createWorkbook()
  addWorksheet(wb, sheet)
  
  # Estilos
  header_style <- createStyle(
    fgFill = "#1F4E78",
    halign = "center",
    valign = "center",
    textDecoration = "bold",
    fontColour = "white",
    border = "TopBottomLeftRight"
  )
  
  body_style <- createStyle(
    border = "TopBottomLeftRight"
  )
  
  footer_style <- createStyle(
    fgFill   = "#E6E6E6",
    halign   = "center",
    valign   = "center",
    wrapText = TRUE,
    border   = "TopBottomLeftRight"
  )
  
  # 1) Tabla desde la fila 4
  writeData(
    wb, sheet,
    x = df,
    startRow = 4, startCol = 1,
    headerStyle = header_style,
    borders = "surrounding"
  )
  
  addStyle(
    wb, sheet,
    style = body_style,
    rows  = 5:(nrow(df) + 4),
    cols  = 1:ncol(df),
    gridExpand = TRUE
  )
  
  setColWidths(
    wb, sheet,
    cols = 1:ncol(df),
    widths = "auto"
  )
  
  addFilter(
    wb, sheet,
    row  = 4,
    cols = 1:ncol(df)
  )
  
  # 2) Texto después de la tabla
  fila_header <- 4 + nrow(df) + 2  # una fila en blanco y luego el texto
  
  writeData(
    wb, sheet,
    x = encabezado,
    startRow = fila_header,
    startCol = 1
  )
  
  mergeCells(
    wb, sheet,
    cols = 1:ncol(df),
    rows = fila_header
  )
  
  addStyle(
    wb, sheet,
    style = footer_style,
    rows  = fila_header,
    cols  = 1:ncol(df),
    gridExpand = TRUE
  )
  
  setRowHeights(
    wb, sheet,
    rows = fila_header,
    heights = 45
  )
  
  saveWorkbook(wb, file, overwrite = TRUE)
  
  invisible(file)
}

