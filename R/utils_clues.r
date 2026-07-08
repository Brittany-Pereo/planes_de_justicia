# R/utils_clues.R
#' Funciones auxiliares para manejo de CLUES
#'
#' Funciones para obtener CLUES relacionadas y construir consultas SQL

#' Obtener CLUES relacionadas (SSA)
#'
#' @param clues_seleccionada Código CLUES seleccionado
#' @param clues_info Data frame con información de CLUES (de datos internos)
#' @return Vector con la(s) CLUES a consultar
#'
#' @export
# R/utils_clues.R

obtener_clues_relacionadas <- function(clues_seleccionada, clues_info) {
  return(clues_seleccionada)
}

construir_clausula_in <- function(clues_seleccionada, clues_info) {
  
  clues_seleccionada <- gsub("'", "''", clues_seleccionada)
  
  valores <- paste0("'", clues_seleccionada, "'", collapse = ", ")
  paste0("(", valores, ")")
}

construir_consulta_clues <- function(clues_seleccionada,
                                     clues_info,
                                     parquet_path,
                                     limite = NULL) {
  
  clausula_in <- construir_clausula_in(clues_seleccionada, clues_info)
  
  filtro_plan <- if (
    is.null(clues_seleccionada) ||
    length(clues_seleccionada) == 0 ||
    clues_seleccionada == "NACIONAL"
  ) {
    "WHERE fecha IS NOT NULL"
  } else {
    sprintf(
      "WHERE \"plan de justicia\" IN %s
       AND fecha IS NOT NULL",
      clausula_in
    )
  }
  
  consulta <- sprintf("
    SELECT
      fecha,
      SUM(CAST(consultas_totales AS INT)) AS consulta_total,
      SUM(CAST(consultas_generales AS INT)) AS consulta_general,
      SUM(CAST(consultas_de_especialidad AS INT)) AS consulta_especialidad,
      SUM(CAST(procedimientos_quirurgicos AS INT)) AS procedimientos_qx,
      SUM(CAST(egresos AS INT)) AS egresos
    FROM parquet_scan('%s')
    %s
    GROUP BY fecha
    ORDER BY fecha
  ", parquet_path, filtro_plan)
  
  if (!is.null(limite)) {
    consulta <- paste(consulta, sprintf("LIMIT %d", limite))
  }
  
  return(consulta)
}


construir_consulta_personas <- function(clues_seleccionada,
                                        clues_info,
                                        parquet_path) {
  
  if (length(clues_seleccionada) == 0 || is.null(clues_seleccionada)) {
    stop("clues_seleccionada no puede estar vacío")
  }
  
  consulta <- sprintf("
    SELECT
      anio_insert AS fecha,
      tipo_procedimiento,
      procedimientos,
      personas
    FROM parquet_scan('%s')
    WHERE id = '%s'
    ORDER BY tipo_procedimiento, fecha
  ", parquet_path, clues_seleccionada)
  
  return(consulta)
}

crear_excel <- function(CLUES, ampliado, resumen = NULL) {
  
  wb <- openxlsx::createWorkbook()
  
  info_sel <- clues_info |>
    dplyr::filter(.data$clues == CLUES | .data$id == CLUES) |>
    dplyr::slice(1)
  
  if (nrow(info_sel) == 0) {
    info_sel <- tibble::tibble(
      nombre = CLUES,
      entidad = ifelse(CLUES == "NACIONAL", "NACIONAL", NA_character_),
      nivel_atencion = NA_character_,
      categoria_gerencial = NA_character_,
      estatus_de_operacion = NA_character_
    )
  }
  
  tabla_info <- tibble::tibble(
    campo = c(
      "Selección",
      "Nombre",
      "Entidad",
      "Nivel de atención",
      "Categoría gerencial",
      "Estatus de operación",
      "Fecha de corte"
    ),
    valor = c(
      CLUES,
      info_sel$nombre,
      info_sel$entidad,
      info_sel$nivel_atencion,
      info_sel$categoria_gerencial,
      info_sel$estatus_de_operacion,
      format(Sys.Date(), "%d/%m/%Y")
    )
  )
  
  resumen_anual <- ampliado |>
    dplyr::mutate(
      anio = lubridate::year(as.Date(fecha))
    ) |>
    dplyr::group_by(anio) |>
    dplyr::summarise(
      dplyr::across(
        where(is.numeric),
        ~ sum(.x, na.rm = TRUE)
      ),
      .groups = "drop"
    ) |>
    dplyr::arrange(anio)
  
  openxlsx::addWorksheet(wb, "resumen")
  
  openxlsx::writeData(
    wb,
    sheet = "resumen",
    x = tabla_info,
    startCol = 1,
    startRow = 1
  )
  
  openxlsx::writeData(
    wb,
    sheet = "resumen",
    x = resumen_anual,
    startCol = 1,
    startRow = 10
  )
  
  openxlsx::addWorksheet(wb, "productividad detalle")
  
  openxlsx::writeData(
    wb,
    sheet = "productividad detalle",
    x = ampliado
  )
  
  return(wb)
}