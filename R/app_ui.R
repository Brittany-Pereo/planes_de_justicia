# R/app_ui.R

app_ui <- function(request) {
  tagList(
    golem_add_external_resources(),
    
    shinydashboard::dashboardPage(
      skin = "blue",
      
      shinydashboard::dashboardHeader(
        title = tags$div(
          icon("hospital"),
          tags$img(src = "www/logo_gob_mx.png", height = "30px")
        ),
        titleWidth = 350
      ),
      
      shinydashboard::dashboardSidebar(
        width = 350,
        shinydashboard::sidebarMenu(
          shinydashboard::menuItem(
            "Consulta",
            tabName = "consulta",
            icon = icon("search")
          ),
          shinydashboard::menuItem(
            "Ayuda",
            tabName = "ayuda",
            icon = icon("question-circle")
          )
        ),
        hr(),
        div(
          style = "padding: 15px;",
          p(icon("database"), " Información 2020 a la fecha"),
          p(icon("code-branch"), " Productividad por plan de justicia"),
          hr(),
          p(
            icon("info-circle"),
            " Incluye información 2020 a la fecha",
            style = "font-size: 12px; color: #9CA3AF;"
          )
        )
      ),
      
      shinydashboard::dashboardBody(
        tags$head(
          tags$style(HTML("
            
            /* Fondo general limpio */
            .content-wrapper, .right-side {
              background-color: #F8F7F5 !important;
            }
            
            /* Barra superior guinda */
            .skin-blue .main-header .logo,
            .skin-blue .main-header .navbar {
              background-color: #611232 !important;
            }
            
            .skin-blue .main-header .logo:hover {
              background-color: #4E0D28 !important;
            }
            
            /* Sidebar gris oscuro */
            .skin-blue .main-sidebar {
              background-color: #222D32 !important;
            }
            
            .skin-blue .sidebar a {
              color: #E5E7EB !important;
            }
            
            .skin-blue .sidebar-menu > li.active > a,
            .skin-blue .sidebar-menu > li:hover > a {
              background-color: #374151 !important;
              border-left-color: #611232 !important;
              color: white !important;
            }
            
            /* Separadores sidebar */
            .main-sidebar hr {
              border-top: 1px solid #611232 !important;
            }
            
            /* Boxes blancos con acento guinda */
            .box {
              border-radius: 6px !important;
              box-shadow: 0 1px 4px rgba(0,0,0,0.12) !important;
              border-top: 3px solid #611232 !important;
            }
            
            .box.box-primary {
              border-top-color: #611232 !important;
            }
            
            .box.box-primary > .box-header {
              background-color: white !important;
              color: #611232 !important;
              border-bottom: 1px solid #E5E7EB !important;
            }
            
            .box.box-primary > .box-header .box-title {
              font-weight: 700 !important;
            }
            
            .box.box-info {
              border-top-color: #611232 !important;
            }
            
            .box.box-info > .box-header {
              background-color: white !important;
              color: #611232 !important;
              border-bottom: 1px solid #E5E7EB !important;
            }
            
            /* Botones dorados elegantes */
            .btn-success {
              background-color: #BC955C !important;
              border-color: #BC955C !important;
              color: white !important;
              font-weight: 700 !important;
              border-radius: 4px !important;
            }
            
            .btn-success:hover,
            .btn-success:active,
            .btn-success:focus {
              background-color: #A57F2C !important;
              border-color: #A57F2C !important;
              color: white !important;
            }
            
            /* Alertas */
            .alert {
              padding: 10px;
              border-radius: 6px;
              margin-bottom: 15px;
            }
            
            .alert-info {
              background-color: #F4F0EA !important;
              border-color: #BC955C !important;
              color: #611232 !important;
            }
            
            .alert-success {
              background-color: #ECFDF5 !important;
              border-color: #1E5B4F !important;
              color: #1E5B4F !important;
            }
            
            .alert-danger {
              background-color: #F2DEDE !important;
              border-color: #EBCCD1 !important;
              color: #A94442 !important;
            }
            
            /* Inputs */
            .selectize-input.focus {
              border-color: #BC955C !important;
              box-shadow: 0 0 4px rgba(188,149,92,0.6) !important;
            }
            
            /* Texto general */
            label {
              color: #374151 !important;
              font-weight: 700 !important;
            }
            
          "))
        ),
        
        shinydashboard::tabItems(
          shinydashboard::tabItem(
            tabName = "consulta",
            fluidRow(
              shinydashboard::box(
                title = "Productividad por Plan de Justicia",
                status = "primary",
                solidHeader = TRUE,
                width = 12,
                uiOutput("modulo_consulta")
              )
            )
          ),
          
          shinydashboard::tabItem(
            tabName = "ayuda",
            fluidRow(
              shinydashboard::box(
                title = "Instrucciones de uso",
                status = "info",
                width = 12,
                h4("¿Cómo usar esta aplicación?"),
                tags$ol(
                  tags$li("Selecciona un plan de justicia del buscador"),
                  tags$li("El sistema automáticamente buscará la información disponible"),
                  tags$li("Los resultados se mostrarán en las gráficas"),
                  tags$li("Puedes descargar los datos en formato Excel"),
                  tags$li("También puedes descargar el informe en PowerPoint")
                ),
                h4("¿Qué incluye?"),
                p("Información de productividad por plan de justicia."),
                h4("Formato de datos"),
                p("Los datos se pueden descargar en un archivo Excel.")
              )
            )
          )
        )
      )
    )
  )
}