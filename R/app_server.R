# R/app_server.R
#' @title Servidor de la aplicación
#' @description Define la lógica del servidor Shiny
#' @import shiny
#' @import duckdb
#' @import DBI
#' @import arrow
#' @import dplyr
#' @import grid
#' @import ggplot2
#' @import clock

#'
#' @noRd
app_server <- function(input, output, session) {
  
  # 1. Cargar datos internos
  load(file.path("R", "sysdata.rda"))
  
  # 2. Configurar conexión a DuckDB
  con <- reactive({
    tryCatch({
      DBI::dbConnect(
        duckdb::duckdb(),
        dbdir = ":memory:",
        read_only = FALSE
      )
    }, error = function(e) {
      showNotification(
        paste("Error al conectar a DuckDB:", e$message),
        type = "error"
      )
      NULL
    })
  })
  
  # 3. Módulo principal
  datos_consulta <- mod_clues_query_server(
    id = "consulta_clues",
    con = con,
    clues_info = clues_info,
    metas_clues = metas_clues
  )
  # 4. UI del módulo
  output$modulo_consulta <- renderUI({
    mod_clues_query_ui("consulta_clues")
  })
  
  # 5. Observador para datos cargados
  observe({
    datos <- datos_consulta()
    
    if (!is.null(datos$datos) && nrow(datos$datos) > 0) {
      cat("\n📊 Datos cargados en la aplicación:\n")
      cat("  - Selección:", datos$clues_seleccionada, "\n")
      cat("  - Total registros:", nrow(datos$datos), "\n")
      cat("  - Columnas:", paste(names(datos$datos), collapse = ", "), "\n")
    }
  })
  
  cat("\n🚀 Aplicación iniciada\n")
}