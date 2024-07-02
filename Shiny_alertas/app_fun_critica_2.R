
library(shiny)
library(shinyauthr)
library(flexdashboard)
library(shinydashboard)
library(shinyjs)
library(DT)
library(RSQLite)
library(pool)
library(uuid)
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(scales)
library(plotly)
library(stringr)
library(tidyverse)
library(zoo)
library(xts)
library(lubridate)
library(TSstudio)
library(tseries)
library(forecast)


source('funciones/generales.R')


source('funciones/alertas.R')
source('funciones/fun_tematica.R')
source('funciones/contribucion.R')


pool <- dbPool(RSQLite::SQLite(), dbname = "prueba4")

# comando para resetear la base de datos
#unlink("prueba4")
#poolClose(pool)

if(dbExistsTable(pool,"responses_df")){
}else{
  #names(responses_df)

  dbWriteTable(pool,
               "responses_df",
               desf,
               overwrite = FALSE,
               append = TRUE)
}

str(dbReadTable(pool, "responses_df"))

# dataframe that holds usernames, passwords and other user data
user_base <- tibble::tibble(
  user = c("Natalia","Julian","Camila"),
  password = sapply(c("nata0110","juli0921","cami0404"), sodium::password_store),
  permissions = c("admin"),
  name = c("Natalia Arteaga","Julian Espinosa","Camila Acosta")
)

ui <- fluidPage(
  useShinyjs(),
  div(id="total", class="main-container",
      style = "background-image: url(img1.png);
          height: 100vh;
          width: 100vw;
          background-position: center;
          background-size: cover;
          position: relative;
          display: flex;
          flex-direction: column;
          justify-content: center;
          align-items: flex-end;",
      div(class = "login-container", shinyauthr::loginUI(id = "login",title="Aplicativo EMMET",
                                                         user_title="Usuario",pass_title="Contraseña",
                                                         login_title = "Iniciar sesión"))),

  div(class = "pull-right", shinyauthr::logoutUI(id = "logout")),
  uiOutput("dashboard")
)

server <- function(input, output, session) {

  credentials <- shinyauthr::loginServer(
    id = "login",
    data = user_base,
    user_col = user,
    pwd_col = password,
    sodium_hashed = TRUE,
    log_out = reactive(logout_init())
  )


  # Logout to hide
  logout_init <- shinyauthr::logoutServer(
    id = "logout",
    active = reactive(credentials()$user_auth)
  )

  observeEvent(req(credentials()$user_auth), {
    # Hide background image when authenticated
    shinyjs::hide("total",anim=T,asis=T)
  })

  observeEvent(req(logout_init()), {
    # Hide background image when authenticated
    shinyjs::show("total",anim=T,asis=T)
  })

  output$dashboard <- renderUI({
    # Mostrar solo cuando el usuario está autenticado
    req(credentials()$user_auth)
    dashboardPage(
      dashboardHeader(
        tags$li(class = "dropdown",
                tags$style(".main-header {max-height: 100px}"),
                tags$style(".main-header .logo {height: 100px}")),
        title = tags$a(href='https://www.dane.gov.co',tags$img(src='logo.png', height = '50px', width ='150px')),
        titleWidth = 180),
      dashboardSidebar(
        tags$style(".left-side, .main-sidebar {background-color: #BF1450;
                                                 margin-top: 100px}"),
        sidebarMenu(
          menuItem("Identificación alertas", tabName = "subitem3", icon = icon("dashboard")),
          menuItem("Consolidado alertas ", tabName = "subitem4", icon = icon("dashboard")),
          menuItem("Agregado variables", tabName = "subitem1", icon = icon("dashboard")),
          menuItem("Comparación", tabName = "subitem2", icon = icon("dashboard")),
          menuItem("Critica", tabName = "subitem5", icon = icon("dashboard"))

        ),width=180),
      dashboardBody(
        tags$head(tags$style(".sidebar-menu li a {
            color: white;
          }")),
        textOutput("usuario"),
        tabItems(
          tabItem("subitem1",
                  fluidRow(
                    box(
                      title = "Filtros", width = 12, solidHeader = TRUE, status = "info",
                      column(6,
                             selectInput("anio", "Año:", choices = c(unique(tematica$ANIO)[-1])),
                             selectInput("mes", "Mes:", choices = c(unique(tematica$MES))),
                             selectInput("dominio", "Actividad ecónomica:", choices = c("Todos",unique(tematica$DESCRIPCIONDOMINIOEMMET39)))),
                      column(6,
                             selectInput("depto", "Departamento:", choices =  c("Todos",unique(tematica$INCLUSION_NOMBRE_DEPTO))),
                             selectInput("metro", "Área metropolitana:", choices =  c("Todos",unique(tematica$AREA_METROPOLITANA))),
                             selectInput("ciudad", "Ciudad:", choices =  c("Todos",unique(tematica$CIUDAD))))))
                  ,
                  fluidRow(
                    navbarPage("",
                               tabPanel("Mes - Año anterior",
                                        plotlyOutput("barra1"),
                                        plotOutput("serie1")),
                               tabPanel("Año corrido",
                                        plotlyOutput("barra2"),
                                        plotOutput("serie2")),
                               tabPanel("Año acumulado",
                                        plotlyOutput("barra3"),
                                        plotOutput("serie3")),
                               tabPanel("precovid",
                                        plotlyOutput("barra4"),
                                        plotOutput("serie4")),
                    )
                  )
          ),
          tabItem("subitem2",
                  fluidRow(
                    box(
                      title = "Filtros", width = 12, solidHeader = TRUE, status = "info",
                      column(4,
                             selectInput("desa", "Comparación:", choices =  c("CIIU","Dpto","Área metropolitana","Ciudad")),
                             selectInput("anio2", "Año:", choices = c(unique(tematica$ANIO)[-1])),
                             selectInput("mes2", "Mes:", choices = c(unique(tematica$MES)))),
                      column(4,
                             selectInput("dominio2", "Actividad ecónomica:", choices = c("Todos",unique(tematica$DESCRIPCIONDOMINIOEMMET39))),
                             selectInput("depto2", "Departamento:", choices =  c("Todos",unique(tematica$INCLUSION_NOMBRE_DEPTO)))),
                      column(4,
                             selectInput("metro2", "Área metropolitana:", choices =  c("Todos",unique(tematica$AREA_METROPOLITANA))),
                             selectInput("ciudad2", "Ciudad:", choices =  c("Todos",unique(tematica$CIUDAD)))))),
                  fluidRow(
                    navbarPage("",
                               tabPanel("Mes - Año anterior",
                                        plotlyOutput("comparacionCIIU_1")),
                               tabPanel("Año corrido",
                                        plotlyOutput("comparacionCIIU_2")),
                               tabPanel("Año acumulado",
                                        plotlyOutput("comparacionCIIU_3")),
                               tabPanel("precovid",
                                        plotlyOutput("comparacionCIIU_4")),
                    )
                  )
          ),
          tabItem("subitem3",
                  fluidRow(
                    box(
                      title = "Filtros", width = 12, solidHeader = TRUE, status = "info",
                      column(4,
                             #selectInput("nombre3", "Nombre Establecimiento:", choices =  c(unique(base_panel$NOMBRE_ESTAB))),
                             selectInput("id_num_crit", "Id NumOrd:", choices = c(unique(base_panel$ID_NUMORD))),
                      ),
                      column(4,
                             selectInput("anio_crit", "AÑO:", choices = c("Todos",unique(desf$ANIO)),selected = max(c(unique(desf$ANIO))))

                      ),
                      column(4,
                             selectInput("mes_crit", "MES:", choices = c("Todos",unique(desf$MES))))
                    )
                  ),
                  fluidRow(
                    column(width=12,
                           #infoBox(tags$span("Año", style="text-transform: none;"),textOutput("anio_imp"),icon=tags$i(class = "fas fa-thumbs-up", style="font-size: 12px; color: white"),color="yellow",value=tags$p(style = "font-size: 10px;")),
                           #infoBox(tags$span("Id_numord", style="text-transform: none;"),textOutput("id_numord_imp"),icon=tags$i(class = "fingerprint", style="font-size: 12px; color: white"),color="yellow",value=tags$p(style = "font-size: 10px;")),
                           infoBox(tags$span("Dominio EMMET", style="text-transform: none;"),textOutput("dominio_imp"),icon=icon("id-card", style="font-size: 80%; color: white"),color="yellow",value=tags$p(style = "font-size: 10px;")),
                           infoBox(tags$span("Departamento", style="text-transform: none;"),textOutput("departamento_imp"),icon=icon("map-pin", style="font-size: 80%; color: white"),color="yellow",value=tags$p(style = "font-size: 10px;")),
                           infoBox(tags$span("Municipio", style="text-transform: none;"),textOutput("municipio_imp"),icon=icon("map-pin", style="font-size: 80%; color: white"),color="yellow",value=tags$p(style = "font-size: 10px;")),
                           infoBox(tags$span("Contribucion", style="text-transform: none;"),textOutput("contribucion_imp"),icon=icon("map-pin", style="font-size: 80%; color: white"),color="yellow",value=tags$p(style = "font-size: 10px;")),
                    )
                  ),
                  #dataTableOutput("tabla"),
                  fluidRow(actionButton("edit_button", "Editar", icon("edit")),
                           actionButton("export_button", "Exportar", icon("copy"))
                  ),

                  #br(),
                  fluidRow(width="100%",
                           #box(title = "Critica",
                           dataTableOutput("responses_table", width = "100%")#)
                  ),
                  selectInput("variable", label = "Variable:",
                              choices = colnames(base_variables_encuesta)[8:length(colnames(base_variables_encuesta))], selected = "Empleados_permanentes"),
                  fluidRow(
                    box(
                      title="Gráfica de la serie",
                      plotlyOutput("boxplot"),width=12,)),
                  fluidRow(
                    box(
                      title="Datos",
                      dataTableOutput("tablaserie"),
                      width=12, status="primary", solidHeader=TRUE))
          ),
          tabItem("subitem4",
                  fluidRow(
                    box(
                      title = "Filtros", width = 12, solidHeader = TRUE, status = "info",
                      column(12,
                             selectInput("capitulo_filtro", "Seleccione el capítulo", choices = c("Capitulo 2", "Capitulo 3")),
                             selectInput("tipo", "Seleccione el tipo de imputación", choices = c("imputacion_deuda", "imputacion_caso_especial")),
                             ))),
                  dataTableOutput("tabla2")

          )#,
          # tabItem("subitem5",
          #         fluidRow(
          #           box(
          #             title = "Filtros", width = 12, solidHeader = TRUE, status = "info",
          #             column(4,
          #                    selectInput("id_num", "ID_NUNMORD:", choices = c("Todos",unique(desf$ID_NUMORD)))
          #
          #                    ),
          #             column(4,
          #                    selectInput("anio_crit", "AÑO:", choices = c("Todos",unique(desf$ANIO)),selected = max(c(unique(desf$ANIO))))
          #
          #             ),
          #             column(4,
          #                    selectInput("mes_crit", "MES:", choices = c("Todos",unique(desf$MES)))
          #
          #             )
          #             )
          #           ),
          #         fluidRow(actionButton("edit_button", "Editar", icon("edit")),
          #                  actionButton("export_button", "Exportar", icon("copy"))
          #                  ),
          #
          #         #br(),
          #         fluidRow(width="100%",
          #                  #box(title = "Critica",
          #                  dataTableOutput("responses_table", width = "100%")#)
          #         ),
          #         #dataTableOutput("responses_table")
          #
          #         )
        )))

  })


  output$usuario <- renderText({
    # use req to only render results when credentials()$user_auth is TRUE
    req(credentials()$user_auth)
    paste0("Bienvenid@", credentials()$info$user)
  })



  dataInput <- reactive({
    datos2 <- tematica
    datos2$ANIO2 <- as.numeric(ifelse(datos2$MES%in%c((as.numeric(input$mes)+1):12),datos2$ANIO+1,datos2$ANIO))

    if(input$dominio!="Todos"){
      datos2 <- datos2 %>% filter(DESCRIPCIONDOMINIOEMMET39==input$dominio)
    }

    if(input$depto!="Todos"){
      datos2 <- datos2 %>% filter(INCLUSION_NOMBRE_DEPTO==input$depto)
    }

    if(input$metro!="Todos"){
      datos2 <- datos2 %>% filter(AREA_METROPOLITANA==input$metro)
    }
    if(input$ciudad!="Todos"){
      datos2 <- datos2 %>% filter(CIUDAD==input$ciudad)
    }
    return(datos2)
  })

  output$barra1 <- renderPlotly({

    variaciones<-estructura(periodo = 1,base=dataInput(),mes=input$mes,anio=input$anio)
    variaciones<-estructura_barra(variaciones,base=dataInput(),mes=input$mes,anio=input$anio)

    barras(data=variaciones)

  })
  output$serie1<-renderPlot({

    variaciones_totales<-estructura(periodo = 1,base=dataInput(),mes=input$mes,anio=input$anio)
    variaciones_totales<-estructura_serie(variaciones_totales,base=dataInput(),mes=input$mes,anio=input$anio)

    series(periodo = 1, data = variaciones_totales)


  })
  output$barra2 <- renderPlotly({

    variaciones<-estructura(periodo = 2,base=dataInput(),mes=input$mes,anio=input$anio)
    variaciones<-estructura_barra(variaciones,base=dataInput(),mes=input$mes,anio=input$anio)

    barras(data=variaciones)


  })
  output$serie2<-renderPlot({

    variaciones_totales<-estructura(periodo = 2,base=dataInput(),mes=input$mes,anio=input$anio)
    variaciones_totales<-estructura_serie(data = variaciones_totales,base=dataInput(),mes=input$mes,anio=input$anio)

    series(periodo = 2, data = variaciones_totales)


  })
  output$barra3 <- renderPlotly({

    variaciones<-estructura(periodo = 3,base=dataInput(),mes=input$mes,anio=input$anio)
    variaciones<-estructura_barra(variaciones,base=dataInput(),mes=input$mes,anio=input$anio)

    barras(data=variaciones)


  })
  output$serie3<-renderPlot({


    variaciones_totales<-estructura(periodo = 3,base=dataInput(),mes=input$mes,anio=input$anio)
    variaciones_totales<-estructura_serie(variaciones_totales,base=dataInput(),mes=input$mes,anio=input$anio)

    series(periodo = 3, data = variaciones_totales)

  })
  output$barra4 <- renderPlotly({

    variaciones<-estructura(periodo = 4,base=dataInput(),mes=input$mes,anio=input$anio)
    variaciones1<-estructura_barra(variaciones,base=dataInput(),mes=input$mes,anio=input$anio)

    barras(data=variaciones1)

  })
  output$serie4<-renderPlot({

    variaciones_totales<-estructura(periodo = 4,base=dataInput(),mes=input$mes,anio=input$anio)
    variaciones_totales<-estructura_serie(variaciones_totales,base=dataInput(),mes=input$mes,anio=input$anio)

    series(periodo = 4, data = variaciones_totales)

  })

  dataCompa <- reactive({
    datos <- tematica
    datos$ANIO2 <- as.numeric(ifelse(datos$MES%in%c((as.numeric(input$mes)+1):12),datos$ANIO+1,datos$ANIO))

    if(input$dominio2!="Todos"){
      datos <- datos %>% filter(DESCRIPCIONDOMINIOEMMET39==input$dominio2)
    }

    if(input$depto2!="Todos"){
      datos <- datos %>% filter(INCLUSION_NOMBRE_DEPTO==input$depto2)
    }

    if(input$metro2!="Todos"){
      datos <- datos %>% filter(AREA_METROPOLITANA==input$metro2)
    }
    if(input$ciudad2!="Todos"){
      datos <- datos %>% filter(CIUDAD==input$ciudad2)
    }

    datos$desa_cat <- ifelse(rep(input$desa=="CIIU",dim(datos)[1]),datos$DESCRIPCIONDOMINIOEMMET39,
                             ifelse(rep(input$desa=="Dpto",dim(datos)[1]),datos$INCLUSION_NOMBRE_DEPTO,
                                    ifelse(rep(input$desa=="Área metropolitana",dim(datos)[1]),as.character(datos$AREA_METROPOLITANA),
                                           ifelse(rep(input$desa=="Ciudad",dim(datos)[1]),datos$CIUDAD,NA))))

    datos$desa_num <- ifelse(rep(input$desa=="CIIU",dim(datos)[1]),datos$DOMINIO_39,
                             ifelse(rep(input$desa=="Dpto",dim(datos)[1]),datos$INCLUSION_NOMBRE_DEPTO,
                                    ifelse(rep(input$desa=="Área metropolitana",dim(datos)[1]),as.character(datos$AREA_METROPOLITANA),
                                           ifelse(rep(input$desa=="Ciudad",dim(datos)[1]),datos$CIUDAD,NA))))


    return(datos)
  })

  output$comparacionCIIU_1 <- renderPlotly({

    tabla1<-tabla_CIIU(periodo = 1,base=dataCompa(),anio2=input$anio2,mes2=input$mes2)

    contribucion(tabla1,desa=input$desa)



  })
  output$comparacionCIIU_2 <- renderPlotly({

    tabla1<-tabla_CIIU(periodo = 2,base=dataCompa(),anio2=input$anio2,mes2=input$mes2)
    contribucion(tabla1,desa=input$desa)


  })
  output$comparacionCIIU_3 <-renderPlotly({

    tabla1<-tabla_CIIU(periodo = 3,base=dataCompa(),anio2=input$anio2,mes2=input$mes2)
    contribucion(tabla1,desa=input$desa)

  })

  output$comparacionCIIU_4 <- renderPlotly({
    tabla1<-tabla_CIIU(periodo = 4,base=dataCompa(),anio2=input$anio2,mes2=input$mes2)
    contribucion(tabla1,desa=input$desa)

  })


  dataEstab <- reactive({
    datos <- desf %>%  filter(ID_NUMORD%in%input$id_num_crit)
    return(datos)
  })

  base_variables <- reactive({
    datos <- base_variables_encuesta %>% filter(ID_NUMORD%in%input$id_num_crit)
    return(datos)
  })

  contribucion_ranking <- reactive({
    #Se halla la contribucion total
    contribucion_total <- tematica %>%
      group_by(ANIO,MES,DOMINIO_39,ID_DEPARTAMENTO)%>%
      summarise(produccion_total = sum(PRODUCCIONREALPOND),
                ventas_total=sum(VENTASREALESPOND),
                personal_total=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))

    #Se halla la contribución específica
    contribucion <- tematica %>%
      mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
      group_by(ANIO,MES,ID_NUMORD,DOMINIO_39,ID_DEPARTAMENTO) %>%
      summarise(prod = sum(PRODUCCIONREALPOND),vent=sum(VENTASREALESPOND),per=sum(PERSONAL))


    contribucion_final<-contribucion %>%
      left_join(contribucion_total, by=c("DOMINIO_39"="DOMINIO_39","MES"="MES","ANIO"="ANIO", "ID_DEPARTAMENTO"="ID_DEPARTAMENTO")) %>%
      mutate(produccion=produccion_total/prod,
             ventas=ventas_total/vent,
             personal=personal_total/per) %>%
      arrange(produccion, ventas,personal) %>%
      group_by(ANIO,MES,ID_DEPARTAMENTO,DOMINIO_39) %>%
      mutate(Ranking=c(1:n())) %>%
      select(ANIO,MES,ID_NUMORD,DOMINIO_39,produccion,ventas,personal,Ranking)


    contribucion_mostrar<- contribucion_final %>%
      filter(ANIO==input$anio_crit & MES==input$mes_crit & ID_NUMORD==input$id_num_crit)

    contribucion_rank<-contribucion_mostrar[,length(contribucion_mostrar)]

    return(contribucion_rank)
  })


  # Critica -----------------------------------------------------------------


  dataCrit <- reactive({
    #input$id_num
    input$submit_edit
    datos3 <- dbReadTable(pool,"responses_df")


    if(input$id_num_crit!="Todos"){
      datos3 <- datos3 %>% filter(ID_NUMORD==input$id_num_crit)
    }
    if(input$anio_crit!="Todos"){
      datos3 <- datos3 %>% filter(ANIO==input$anio_crit)
    }
    if(input$mes_crit!="Todos"){
      datos3 <- datos3 %>% filter(MES==input$mes_crit)
    }

    return(datos3)
  })


  entry_form <- function(button_id){

    showModal(
      modalDialog(
        div(id=("entry_form"),
            tags$head(tags$style(".modal-dialog{ width:400px}")), #Modify the width of the dialog
            tags$head(tags$style(HTML(".shiny-split-layout > div {overflow: visible}"))), #Necessary to show the input options
            fluidPage(
              fluidRow(
                splitLayout(
                  cellWidths = c("250px", "100px"),
                  cellArgs = list(style = "vertical-align: top"),
                ),
                #sliderInput("age", "Age", 0, 100, 1, ticks = TRUE, width = "354px"),
                textAreaInput("Caso_imputacion", "Caso_imputacion", placeholder = "", height = 100, width = "354px"),
                textAreaInput("OBSERVACIONES", "OBSERVACIONES", placeholder = "", height = 100, width = "354px"),
                textAreaInput(paste0(meses_c[mes]), paste0(meses_c[mes]), placeholder = "", height = 100, width = "354px"),
                #helpText(labelMandatory(""), paste("Mandatory field.")),
                actionButton(button_id, "Submit")
              ),
              easyClose = TRUE
            )
        )
      )
    )
  }


  observeEvent(input$edit_button, priority = 20,{

    SQL_df <- dbReadTable(pool, "responses_df")

    showModal(
      if(length(input$responses_table_rows_selected) > 1 ){
        modalDialog(
          title = "Warning",
          paste("Please select only one row." ),easyClose = TRUE)
      } else if(length(input$responses_table_rows_selected) < 1){
        modalDialog(
          title = "Warning",
          paste("Please select a row." ),easyClose = TRUE)
      })

    if(length(input$responses_table_rows_selected) == 1 ){

      entry_form("submit_edit")
      updateTextAreaInput(session, "Caso_imputacion", value = "continua_critica")
      updateTextAreaInput(session, "OBSERVACIONES", value = SQL_df[input$responses_table_rows_selected, "OBSERVACIONES"])
      updateTextAreaInput(session, paste0(meses_c[mes]), value = SQL_df[input$responses_table_rows_selected, paste0(meses_c[mes])])
      #updateTextAreaInput(session, "EDITADO", value = SQL_df[input$responses_table_rows_selected, "EDITADO"])

    }

  })

  observeEvent(input$submit_edit, priority = 20, {

    datos3 <- dbReadTable(pool, "responses_df")


    if(input$id_num_crit!="Todos"){
      datos3 <- datos3 %>% filter(ID_NUMORD==input$id_num_crit)
    }
    if(input$anio_crit!="Todos"){
      datos3 <- datos3 %>% filter(ANIO==input$anio_crit)
    }
    if(input$mes_crit!="Todos"){
      datos3 <- datos3 %>% filter(MES==input$mes_crit)
    }

    row_selection <- datos3[input$responses_table_row_last_clicked, "LLAVE"]
    row_selection2 <- datos3[input$responses_table_row_last_clicked, "Variables"]
    #row_selection<-"123455465"
    dbExecute(pool, sprintf('UPDATE "responses_df"
                        SET "OBSERVACIONES" = ?,
                            "Caso_imputacion" = ?,
                            `%s` = ?,
                            "EDITADO" = ?
                        WHERE "LLAVE" = ? AND "Variables" = ?',
                            meses_c[mes]),
              param = list(
                input$OBSERVACIONES,
                input$Caso_imputacion,
                input[[meses_c[mes]]], # Acceder al valor de la columna directamente usando [[]]
                credentials()$info$user,
                row_selection,
                row_selection2
              ))
    removeModal()

    #, "EDITADO" = credentials()$info$user
    #   credentials()$info$user


  })
  output$responses_table <- DT::renderDataTable({

    table <- dataCrit() %>% select("LLAVE","ID_NUMORD","Variables",meses_c[mes-1],
                                   meses_c[mes],"variacion","Caso_imputacion",
                                   "OBSERVACIONES","EDITADO")#,"NOMBRE_ESTAB","DOMINIOEMMET39"
    #names(table) <- c("ID","Date", "Name", "Sex", "Comment", "Age")
    table <- datatable(table,
                       selection = "single",
                       rownames = FALSE,
                       options = list(searching = FALSE, lengthChange = FALSE)
    )
  })

  observeEvent(input$export_button, priority = 20,{

    SQL_df <- dbReadTable(pool, "responses_df")
    b<-choose.dir()
    #dir <- tclvalue(tk_chooseDirectory(title = "Seleccionar directorio", message = "Por favor seleccione un directorio"))
    #write.csv(responses_df,"EMMET_Critica.csv")
    setwd(b)
    write.csv(SQL_df,"EMMET_Critica.csv")

  })



  # output$tabla <- renderDataTable({
  #
  #   tabla <- dataEstab() %>%
  #     select(NOMBRE_ESTAB,Variables,Octubre,Noviembre,variacion,Caso_imputacion)
  #     #mutate(Observaciones = map(rowname, .f = ~textinput_gt(.x,"_textinput", label = paste("My text label", .x))))
  #   colnames(tabla)[c(3,4)] <- c("Octubre 2022","Noviembre 2022")
  #   tabla
  # },options = list(pageLength = 50))

  output$tablaserie <- renderDataTable({
    tabla <- base_variables() %>%
      select(ANIO,MES,NOMBRE_ESTAB,input$variable)
    tabla
  },options = list(pageLength = 100))

  output$boxplot <- renderPlotly({
    tabla <- base_variables()
    smexter_a=ts(tabla[,input$variable],start=c(2018,1),frequency=12)
    ts_plot(smexter_a,
            title = paste0("Serie ",as.character(input$variable)," para el establecimiento ",as.character(input$nombre3)),
            Ytitle = "Valor",
            Xtitle = "Año",
            Xgrid = TRUE,
            Ygrid = TRUE)

  })
  output$anio_imp <- renderText({
    dataEstab()$ANIO[1]

  })
  output$id_numord_imp <- renderText({
    dataEstab()$ID_NUMORD[1]
  })
  output$dominio_imp <- renderText({
    dataEstab()$DOMINIOEMMET39[1]
  })
  output$departamento_imp <- renderText({
    dataEstab()$NOMBREDEPARTAMENTO[1]
  })
  output$municipio_imp <- renderText({
    dataEstab()$NOMBREMUNICIPIO[1]
  })
  output$contribucion_imp<-renderText({
    paste0("Ranking: ",contribucion_ranking())
  })


  # shinyInput = function(FUN, len, id, ...) {
  #   #validate(need(character(len)>0,message=paste("")))
  #   inputs = character(len)
  #   for (i in seq_len(len)) {
  #     inputs[i] = as.character(FUN(paste0(id, i), label = NULL, ...))
  #   }
  #   inputs
  # }

  output$tabla2<- renderDataTable({


    # Filtrar la tabla segC:n la selección del usuario
    if (input$capitulo_filtro == "Capitulo 3") {
      # Si se selecciona "Capítulo 2", filtrar la tabla segC:n la lista "capitulo_2"
      tabla <- desf %>%select(NOMBRE_ESTAB,Variables,meses_c[mes-1],meses_c[mes],variacion,Caso_imputacion) %>%
        filter((Variables %in% capitulo3) & Caso_imputacion==input$tipo)

      #pool <- dbPool(RSQLite::SQLite(), dbname = "Cap3")

    } else {
      # Si se selecciona "Capitulo 3", filtrar la tabla segC:n la lista "capitulo_3"
      tabla <- desf %>%select(NOMBRE_ESTAB,Variables,meses_c[mes-1],meses_c[mes],variacion,Caso_imputacion) %>%
        filter(!(Variables %in% capitulo3) & Caso_imputacion==input$tipo)

      #pool <- dbPool(RSQLite::SQLite(), dbname = "Cap2")
    }
    # Devolver la tabla filtrada
    colnames(tabla)[c(3,4)] <- c(paste0(meses_c[mes-1]," ",anio),paste0(meses_c[mes]," ",anio))
    tabla
  },options = list(pageLength = 50))



  # Critica -----------------------------------------------------------------
  #
  #
  #   dataCrit <- reactive({
  #     #input$id_num
  #     input$submit_edit
  #     datos3 <- dbReadTable(pool,"responses_df")
  #
  #
  #     if(input$id_num!="Todos"){
  #       datos3 <- datos3 %>% filter(ID_NUMORD==input$id_num)
  #     }
  #     if(input$anio_crit!="Todos"){
  #       datos3 <- datos3 %>% filter(ANIO==input$anio_crit)
  #     }
  #     if(input$mes_crit!="Todos"){
  #       datos3 <- datos3 %>% filter(MES==input$mes_crit)
  #     }
  #
  #     return(datos3)
  #   })
  #
  #
  #
  #   # dataCrit_0 <- reactive({
  #   #
  #   #   input$submit_edit
  #   #   dbReadTable(pool, "responses_df")
  #   #
  #   #
  #   # })
  #
  #   entry_form <- function(button_id){
  #
  #     showModal(
  #       modalDialog(
  #         div(id=("entry_form"),
  #             tags$head(tags$style(".modal-dialog{ width:400px}")), #Modify the width of the dialog
  #             tags$head(tags$style(HTML(".shiny-split-layout > div {overflow: visible}"))), #Necessary to show the input options
  #             fluidPage(
  #               fluidRow(
  #                 splitLayout(
  #                   cellWidths = c("250px", "100px"),
  #                   cellArgs = list(style = "vertical-align: top"),
  #                   #textInput("name", labelMandatory("Name"), placeholder = ""),
  #                   #selectInput("sex", labelMandatory("Sex"), multiple = FALSE, choices = c("", "M", "F"))
  #                 ),
  #                 #sliderInput("age", "Age", 0, 100, 1, ticks = TRUE, width = "354px"),
  #                 textAreaInput("Caso_imputacion", "Caso_imputacion", placeholder = "", height = 100, width = "354px"),
  #                 textAreaInput("OBSERVACIONES", "OBSERVACIONES", placeholder = "", height = 100, width = "354px"),
  #                 textAreaInput("variacion", "variacion", placeholder = "", height = 100, width = "354px"),
  #                 #helpText(labelMandatory(""), paste("Mandatory field.")),
  #                 actionButton(button_id, "Submit")
  #               ),
  #               easyClose = TRUE
  #             )
  #         )
  #       )
  #     )
  #   }
  #
  #
  #   observeEvent(input$edit_button, priority = 20,{
  #
  #     SQL_df <- dbReadTable(pool, "responses_df")
  #
  #     showModal(
  #       if(length(input$responses_table_rows_selected) > 1 ){
  #         modalDialog(
  #           title = "Warning",
  #           paste("Please select only one row." ),easyClose = TRUE)
  #       } else if(length(input$responses_table_rows_selected) < 1){
  #         modalDialog(
  #           title = "Warning",
  #           paste("Please select a row." ),easyClose = TRUE)
  #       })
  #
  #     if(length(input$responses_table_rows_selected) == 1 ){
  #
  #       entry_form("submit_edit")
  #       updateTextAreaInput(session, "Caso_imputacion", value = SQL_df[input$responses_table_rows_selected, "Caso_imputacion"])
  #       updateTextAreaInput(session, "OBSERVACIONES", value = SQL_df[input$responses_table_rows_selected, "OBSERVACIONES"])
  #       updateTextAreaInput(session, "variacion", value = SQL_df[input$responses_table_rows_selected, "variacion"])
  #       #updateTextAreaInput(session, "EDITADO", value = SQL_df[input$responses_table_rows_selected, "EDITADO"])
  #
  #     }
  #
  #   })
  #
  #   observeEvent(input$submit_edit, priority = 20, {
  #
  #     datos3 <- dbReadTable(pool, "responses_df")
  #
  #
  #     if(input$id_num!="Todos"){
  #       datos3 <- datos3 %>% filter(ID_NUMORD==input$id_num)
  #     }
  #     if(input$anio_crit!="Todos"){
  #       datos3 <- datos3 %>% filter(ANIO==input$anio_crit)
  #     }
  #     if(input$mes_crit!="Todos"){
  #       datos3 <- datos3 %>% filter(MES==input$mes_crit)
  #     }
  #
  #     row_selection <- datos3[input$responses_table_row_last_clicked, "LLAVE"]
  #     row_selection2 <- datos3[input$responses_table_row_last_clicked, "Variables"]
  #     #row_selection<-"123455465"
  #     dbExecute(pool, sprintf('UPDATE "responses_df" SET "OBSERVACIONES" = ? , "Caso_imputacion" = ?, "variacion" = ?, "EDITADO" = ?
  #                             WHERE "LLAVE" = ("%s") AND "Variables" =  ("%s")' , row_selection, row_selection2 ),
  #               param = list(
  #                 input$OBSERVACIONES,
  #                 input$Caso_imputacion,
  #                 input$variacion,
  #                 credentials()$info$user))
  #     removeModal()
  #
  #     #, "EDITADO" = credentials()$info$user
  #     #   credentials()$info$user
  #
  #
  #   })
  #   output$responses_table <- DT::renderDataTable({
  #
  #     table <- dataCrit() %>% select("LLAVE","ID_NUMORD","Variables","Octubre",
  #                                    "Noviembre","variacion","Caso_imputacion",
  #                                    "OBSERVACIONES","EDITADO")#,"NOMBRE_ESTAB","DOMINIOEMMET39"
  #     #names(table) <- c("ID","Date", "Name", "Sex", "Comment", "Age")
  #     table <- datatable(table,
  #                        rownames = FALSE,
  #                        options = list(searching = FALSE, lengthChange = FALSE)
  #     )
  #   })
  #
  #   observeEvent(input$export_button, priority = 20,{
  #
  #     SQL_df <- dbReadTable(pool, "responses_df")
  #     b<-choose.dir()
  #     #dir <- tclvalue(tk_chooseDirectory(title = "Seleccionar directorio", message = "Por favor seleccione un directorio"))
  #     #write.csv(responses_df,"EMMET_Critica.csv")
  #     setwd(b)
  #     write.csv(SQL_df,"EMMET_Critica.csv")
  #
  #   })

}



shinyApp(ui = ui, server = server)

