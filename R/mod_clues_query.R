# R/mod_clues_query.R

mod_clues_query_ui <- function(id) {
  ns <- NS(id)
  tagList(
    uiOutput(ns("selector_clues")),
    uiOutput(ns("estado_consulta")),
    
    fluidRow(
      column(6, plotOutput(ns("grafica_general"), height = "280px")),
      column(6, plotOutput(ns("grafica_especialidad"), height = "280px"))
    ),
    fluidRow(
      column(6, plotOutput(ns("grafica_qx"), height = "280px")),
      column(6, plotOutput(ns("grafica_egresos"), height = "280px"))
    ),
    
    fluidRow(
      column(
        6,
        downloadButton(
          ns("descargar_datos"),
          "Descargar datos",
          class = "btn-success btn-block"
        )
      ),
      column(
        6,
        downloadButton(
          ns("btn_crear_pptx"),
          "Descargar informe",
          icon = icon("file-powerpoint"),
          class = "btn-success btn-block"
        )
      )
    )
  )
}


mod_clues_query_server <- function(id, con, clues_info) {
  moduleServer(id, function(input, output, session) {
    
    valores <- reactiveValues(
      datos = NULL,
      clues_seleccionada = NULL,
      consulta_actual = NULL,
      cargando = FALSE,
      error = NULL
    )
    
    parquet_path <- reactive({
      path <- system.file(
        "app", "data", "Cubos_completos_2020_2025.parquet",
        package = "pptx"
      )
      
      if (path == "") {
        valores$error <- "No se encontró el archivo Parquet"
        return(NULL)
      }
      
      path
    })
    
    output$selector_clues <- renderUI({
      ns <- session$ns
      
      choices_etiquetas <- clues_info %>%
        dplyr::mutate(
          etiqueta = dplyr::if_else(
            is.na(nombre) | nombre == "",
            clues,
            paste0(clues, " - ", nombre)
          )
        ) %>%
        dplyr::arrange(entidad, nombre) %>%
        dplyr::pull(clues, name = etiqueta)
      
      choices_etiquetas <- c(
        "NACIONAL" = "NACIONAL",
        choices_etiquetas
      )
      
      shiny::selectizeInput(
        ns("clues_select"),
        "Selecciona un Plan de Justicia:",
        choices = choices_etiquetas,
        selected = "NACIONAL",
        options = list(
          placeholder = "Busca por CLUES o nombre de unidad",
          maxOptions = 1000
        )
      )
    })
    
    observeEvent(input$clues_select, {
      req(input$clues_select)
      req(parquet_path())
      req(con())
      
      valores$clues_seleccionada <- input$clues_select
      valores$cargando <- TRUE
      valores$error <- NULL
      
      clues_relacionadas <- obtener_clues_relacionadas(
        input$clues_select,
        clues_info
      )
      
      consulta <- tryCatch({
        construir_consulta_clues(
          clues_seleccionada = input$clues_select,
          clues_info = clues_info,
          parquet_path = parquet_path(),
          limite = 1000
        )
      }, error = function(e) {
        valores$error <- paste("Error al construir consulta:", e$message)
        NULL
      })
      
      if (!is.null(consulta)) {
        valores$consulta_actual <- consulta
        
        tryCatch({
          cat("\n📝 Consulta SQL ejecutada:\n")
          cat(consulta, "\n")
          
          resultados <- DBI::dbGetQuery(con(), consulta)
          
          if (nrow(resultados) > 0) {
            valores$datos <- resultados
            
            cat("✅ Consulta exitosa. Registros obtenidos:", nrow(resultados), "\n")
            cat("📊 Columnas:", paste(names(resultados), collapse = ", "), "\n")
          } else {
            valores$error <- paste(
              "No se encontraron datos para:",
              paste(clues_relacionadas, collapse = ", ")
            )
            valores$datos <- NULL
          }
          
        }, error = function(e) {
          valores$error <- paste("Error al ejecutar consulta:", e$message)
          valores$datos <- NULL
          cat("❌", valores$error, "\n")
        })
      }
      
      valores$cargando <- FALSE
    })
    
    output$estado_consulta <- renderUI({
      if (valores$cargando) {
        div(
          class = "alert alert-info",
          icon("spinner", class = "fa-spin"),
          " Ejecutando consulta en DuckDB..."
        )
      } else if (!is.null(valores$error)) {
        div(
          class = "alert alert-danger",
          icon("exclamation-triangle"),
          " ", valores$error
        )
      } else if (!is.null(valores$datos) && nrow(valores$datos) > 0) {
        div(
          class = "alert alert-success",
          icon("check-circle"),
          sprintf(
            " ✅ Consulta exitosa: %d registros encontrados para: %s",
            nrow(valores$datos),
            valores$clues_seleccionada
          )
        )
      }
    })
    
    excel_exportado <- reactive({
      req(input$clues_select)
      req(valores$datos)
      
      crear_excel(
        CLUES = input$clues_select,
        ampliado = valores$datos,
        resumen = NULL
      )
    })
    
    datos_anual_grafica <- reactive({
      req(valores$datos)
      
      valores$datos %>%
        dplyr::mutate(
          fecha = as.Date(fecha),
          anio = lubridate::year(fecha)
        ) %>%
        dplyr::filter(anio %in% c(2024, 2025, 2026)) %>%
        dplyr::group_by(anio) %>%
        dplyr::summarise(
          consulta_general_anual = sum(consulta_general, na.rm = TRUE),
          consulta_especialidad_anual = sum(consulta_especialidad, na.rm = TRUE),
          procedimientos_qx_anual = sum(procedimientos_qx, na.rm = TRUE),
          egresos_anual = sum(egresos, na.rm = TRUE),
          .groups = "drop"
        )
    })
    
    metas_filtrado_grafica <- reactive({
      req(input$clues_select)
      
      if (input$clues_select == "NACIONAL") {
        metas_clues
      } else {
        metas_clues %>%
          dplyr::filter(clues == input$clues_select)
      }
    })
    
    crear_grafica_clues <- function(df, variable_sel, titulo,
                                    datos_anual_grafica, metas_filtrado) {
      
      fecha_corte <- max(as.Date(df$fecha), na.rm = TRUE)
      mes_corte <- lubridate::month(fecha_corte)
      dia_corte <- lubridate::day(fecha_corte)
      
      col_anual <- dplyr::case_when(
        variable_sel == "consulta_general" ~ "consulta_general_anual",
        variable_sel == "consulta_especialidad" ~ "consulta_especialidad_anual",
        variable_sel == "procedimientos_qx" ~ "procedimientos_qx_anual",
        variable_sel == "egresos" ~ "egresos_anual",
        TRUE ~ NA_character_
      )
      
      df_avance <- df %>%
        dplyr::mutate(
          fecha = as.Date(fecha),
          anio = lubridate::year(fecha),
          fecha_corte_anio = lubridate::ymd(
            paste0(anio, "-", mes_corte, "-", dia_corte)
          )
        ) %>%
        dplyr::filter(anio %in% c(2024, 2025, 2026)) %>%
        dplyr::group_by(anio) %>%
        dplyr::summarise(
          avance = sum(.data[[variable_sel]][fecha <= fecha_corte_anio], na.rm = TRUE),
          .groups = "drop"
        )
      
      hay_2026 <- df %>%
        dplyr::mutate(anio = lubridate::year(as.Date(fecha))) %>%
        dplyr::filter(anio == 2026) %>%
        dplyr::summarise(
          hay = any(.data[[variable_sel]] > 0, na.rm = TRUE),
          .groups = "drop"
        ) %>%
        dplyr::pull(hay)
      
      if (length(hay_2026) == 0 || is.na(hay_2026)) hay_2026 <- FALSE
      
      df_total <- datos_anual_grafica %>%
        dplyr::mutate(anio = as.numeric(anio)) %>%
        dplyr::filter(anio %in% c(2024, 2025, 2026)) %>%
        dplyr::transmute(
          anio,
          total_anual = .data[[col_anual]]
        ) %>%
        dplyr::mutate(
          total_anual = dplyr::case_when(
            hay_2026 & anio == 2026 & variable_sel == "consulta_general" ~
              sum(metas_filtrado$meta_general_anual, na.rm = TRUE),
            hay_2026 & anio == 2026 & variable_sel == "consulta_especialidad" ~
              sum(metas_filtrado$meta_especialidad_anual, na.rm = TRUE),
            hay_2026 & anio == 2026 & variable_sel == "procedimientos_qx" ~
              sum(metas_filtrado$meta_cirugia_anual, na.rm = TRUE),
            hay_2026 & anio == 2026 & variable_sel == "egresos" ~
              sum(metas_filtrado$meta_egresos_anual, na.rm = TRUE),
            TRUE ~ total_anual
          )
        )
      
      df_plot <- df_avance %>%
        dplyr::left_join(df_total, by = "anio") %>%
        dplyr::mutate(
          pendiente = pmax(total_anual - avance, 0),
          anio = as.character(anio)
        ) %>%
        dplyr::select(anio, avance, pendiente, total_anual) %>%
        tidyr::pivot_longer(
          cols = c(avance, pendiente),
          names_to = "tipo",
          values_to = "valor"
        ) %>%
        dplyr::mutate(
          tipo = factor(
            tipo,
            levels = c("avance", "pendiente"),
            labels = c("Avance al corte", "Resto del año")
          ),
          color_barra = dplyr::case_when(
            anio == "2026" & tipo == "Resto del año" ~ "#B08D57",
            tipo == "Resto del año" ~ "#D9D2BE",
            TRUE ~ "#1E5B4F"
          )
        )
      
      etiquetas <- df_plot %>%
        dplyr::group_by(anio) %>%
        dplyr::summarise(
          total_anual = sum(valor, na.rm = TRUE),
          .groups = "drop"
        )
      
      etiquetas_valores <- df_avance %>%
        dplyr::left_join(df_total, by = "anio") %>%
        dplyr::mutate(
          pendiente = pmax(total_anual - avance, 0),
          pct_avance = avance / total_anual,
          anio = as.character(anio),
          etiqueta_pct = scales::percent(pct_avance, accuracy = 1),
          etiqueta_avance = scales::comma(avance)
        )
      
      ggplot2::ggplot(df_plot, ggplot2::aes(x = anio, y = valor, fill = color_barra)) +
        ggplot2::geom_col(
          width = 0.65,
          position = ggplot2::position_stack(reverse = TRUE)
        ) +
        ggplot2::geom_text(
          data = etiquetas,
          ggplot2::aes(
            x = anio,
            y = total_anual,
            label = scales::comma(total_anual)
          ),
          inherit.aes = FALSE,
          vjust = -0.4,
          fontface = "bold",
          size = 5
        ) +
        ggplot2::geom_text(
          data = etiquetas_valores,
          ggplot2::aes(
            x = anio,
            y = avance / 2,
            label = etiqueta_avance
          ),
          inherit.aes = FALSE,
          color = "white",
          fontface = "bold",
          size = 5
        ) +
        ggplot2::geom_text(
          data = etiquetas_valores,
          ggplot2::aes(
            x = anio,
            y = avance + (pendiente * 0.1),
            label = etiqueta_pct
          ),
          inherit.aes = FALSE,
          color = "black",
          fontface = "bold",
          size = 5
        ) +
        ggplot2::scale_fill_identity(
          guide = "legend",
          breaks = c("#D9D2BE", "#1E5B4F", "#B08D57"),
          labels = c("Resto del año", "Avance al corte", "Meta")
        ) +
        ggplot2::scale_y_continuous(
          labels = scales::comma,
          expand = ggplot2::expansion(mult = c(0, 0.18))
        ) +
        ggplot2::labs(title = titulo, x = NULL, y = NULL, fill = NULL) +
        ggplot2::theme_minimal(base_family = "Noto Sans") +
        ggplot2::theme(
          plot.title = ggplot2::element_text(
            hjust = 0.5,
            face = "bold",
            size = 18,
            color = "#6B7280"
          ),
          axis.text.x = ggplot2::element_text(
            size = 13,
            face = "bold",
            color = "#6B7280"
          ),
          axis.text.y = ggplot2::element_text(
            size = 11,
            color = "#6B7280"
          ),
          legend.position = "bottom",
          legend.text = ggplot2::element_text(
            size = 14,
            face = "bold"
          ),
          panel.grid.major.x = ggplot2::element_blank(),
          panel.grid.minor = ggplot2::element_blank()
        )
    }
    
    output$grafica_general <- renderPlot({
      req(valores$datos, datos_anual_grafica(), metas_filtrado_grafica())
      
      crear_grafica_clues(
        valores$datos,
        "consulta_general",
        "Consulta general",
        datos_anual_grafica(),
        metas_filtrado_grafica()
      )
    })
    
    output$grafica_especialidad <- renderPlot({
      req(valores$datos, datos_anual_grafica(), metas_filtrado_grafica())
      
      crear_grafica_clues(
        valores$datos,
        "consulta_especialidad",
        "Consulta de especialidad",
        datos_anual_grafica(),
        metas_filtrado_grafica()
      )
    })
    
    output$grafica_qx <- renderPlot({
      req(valores$datos, datos_anual_grafica(), metas_filtrado_grafica())
      
      crear_grafica_clues(
        valores$datos,
        "procedimientos_qx",
        "Procedimientos quirúrgicos",
        datos_anual_grafica(),
        metas_filtrado_grafica()
      )
    })
    
    output$grafica_egresos <- renderPlot({
      req(valores$datos, datos_anual_grafica(), metas_filtrado_grafica())
      
      crear_grafica_clues(
        valores$datos,
        "egresos",
        "Egresos",
        datos_anual_grafica(),
        metas_filtrado_grafica()
      )
    })
    
    output$descargar_datos <- downloadHandler(
      filename = function() {
        paste0(
          "datos_clues_",
          valores$clues_seleccionada,
          "_",
          Sys.Date(),
          ".xlsx"
        )
      },
      content = function(file) {
        req(valores$datos)
        openxlsx::saveWorkbook(excel_exportado(), file, overwrite = TRUE)
      }
    )
    
    output$btn_crear_pptx <- downloadHandler(
      filename = function() {
        paste0(
          "datos_clues_",
          valores$clues_seleccionada,
          "_",
          Sys.Date(),
          ".pptx"
        )
      },
      contentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      content = function(file) {
        req(valores$datos)
        req(valores$clues_seleccionada)
        
        showNotification(
          "Generando informe en PowerPoint...",
          type = "default",
          duration = 3,
          session = session
        )
        
        presentacion <- crear_reporte_productividad(
          codigo_clues = valores$clues_seleccionada,
          clues_info = clues_info,
          metas = metas_clues,
          historicos = valores$datos,
          procedimientos_personas = NULL,
          ruta_master = system.file(
            "app", "data", "master_presentacion.pptx",
            package = "pptx"
          )
        )
        
        archivo_tmp <- tempfile(fileext = ".pptx")
        print(presentacion, target = archivo_tmp)
        file.copy(archivo_tmp, file, overwrite = TRUE)
        
        showNotification(
          "¡Informe generado exitosamente!",
          type = "default",
          duration = 5,
          session = session
        )
      }
    )
    
    return(reactive({
      list(
        datos = valores$datos,
        resumen = NULL,
        clues_seleccionada = valores$clues_seleccionada,
        consulta = valores$consulta_actual
      )
    }))
  })
}