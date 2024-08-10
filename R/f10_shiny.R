#' Shiny alertas
#'
#' @description Está función abre el navegador para abrir la aplicación de alertas, en la cual se puede loguear
#' y se encontraran 4 ventanas
#'
#' 1. identificación de alertas: la cual se puede filtrar por ID_NUMORD, mes y año; se encontrará una tabla con
#' los valores para las variables de capítulo 2 y 3 para los filtros aplicados, grafico de la serie para la variable
#' que se seleccione y finalmente los datos para el establecimiento en esa variable desde 2018 hasta la fecha.
#'
#' 2. Consolidado de alertas: la cual se puede filtrar por capitulo a la que pertenece la variable y el tipo de
#' imputación (imputación deuda o imputación caso especial) en el año y mes de interés, tiene el registro de todos
#' los establecimientos y variables en los que se identificó un posible caso de imputación.
#'
#' 3. Agregado variables: la cual se puede filtrar por año, mes, actividad económica, departamento, área metropolitana
#' y ciudad, se observan 2 gráficos, el primero un gráfico de barras en el cual se muestra la variación de la contribución según los filtros
#' y en el segundo grafico se observa la serie de tiempo desde 2019, estos gráficos se pueden observar por mes-año anterior, año corrido, año acumulado y
#' precovid
#'
#' 4. Comparación: la cual se puede filtrar por comparación entre CIUU,DPTO,ciudad y área metropolitana, Adicionalmente se puede filtrar por año, mes, actividad económica
#' departamento, área metropolitana y ciudad, en la cual se observa un gráfico de barras para producción, para ventas
#' y para personal, donde la elongación de la barra se refiere a la variación y el valor es referente a la contribución, nuevamente estos
#' gráficos se pueden observar por mes-año anterior, año corrido, año acumulado y
#' precovid

#'
#' @param mes Definir el mes a ejecutar, ej: 11
#' @param anio Definir el año a ejecutar, ej: 2022
#' @param directorio definir el directorio donde se encuentran ubicado los datos de entrada
#'
#'
#'
#' @examples f10_shiny(directorio"Documents/DANE/Procesos DIMPE /PilotoEMMET",mes = 11  ,anio =2022)





f10_shiny <- function(directorio, mes, anio) {
  library(rmarkdown)
  library(readxl)
  library(lubridate)
  library(shiny)
  source("https://raw.githubusercontent.com/sub-dane/EMMET/main/R/utils.R")
  assign("directorio", directorio, envir = .GlobalEnv)
  assign("mes", mes, envir = .GlobalEnv)
  assign("anio", anio, envir = .GlobalEnv)

  url_boletin <- "https://github.com/NataliArteaga/DANE.EMMET/raw/main/Shiny_alertas.zip"

  # Descargar y descomprimir la carpeta shiny
  archivo_zip <- file.path(directorio, "Shiny_alertas.zip")
  download.file(url_boletin, destfile = archivo_zip)
  unzip(archivo_zip, exdir = file.path(directorio))

  ruta_app_shiny <- file.path(directorio,"Shiny_alertas/app_fun_critica_2.R")

  ruta_archivo_excel<-paste0(directorio,"/results/S4_tematica/EMMET_PANEL_tematica_",meses[mes],anio,".csv")

  #identificar si el archivo de tematica existe
  if (!file.exists(ruta_archivo_excel)) {
    # Si el archivo no está presente, ejecuta la función f4_imputacion y f5_tematica
    f3_imputacion(directorio,mes,anio)
    f4_tematica(directorio,mes,anio)
  } else {
  }

#función que corre la app
  shiny::runApp(
    appDir = ruta_app_shiny,
    launch.browser = TRUE  # Esto abrirá automáticamente el navegador para mostrar la aplicación
  )

  }
