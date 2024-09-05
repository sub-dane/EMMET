#' Función inicial
#'
#' Funcion que instala las librerias necesarias para todo el proceso y crea las carpetas en donde se guardaran los archivos.
#'
#' @param directorio definir el directorio donde se crearan las carpetas.
#'
#' @details Se revisa si las funciones están instaladas en el entorno, en caso de
#' que no estén instaladas se procederán a instalar, luego se procede a revisar
#' si están creadas las carpetas donde se guardarán los archivos creados con la librería
#' en total se crearán siete carpetas, para cada una de las funciones que genera algún archivo de salida.
#'
#' 1 S1_integracion:\code{\link{f1_integracion}}
#'
#' 2 S2_estandarizacion:\code{\link{f2_estandarizacion}}
#'
#' 3 S3_identificacion_alertas:\code{\link{f3_identificacion_alertas}}
#'
#' 4 S4_imputacion:\code{\link{f4_imputacion}}
#'
#' 5 S5_tematica:\code{\link{f5_tematica}}
#'
#' 6 S6_anexos:\code{\link{f6_anacional}} y \code{\link{f7_aterritorial}}
#'
#' 7 S7_boletin:\code{\link{f8_boletin}}
#'
#' @examples f0_inicial(directorio="Documents/DANE/Procesos DIMPE /PilotoEMMET")

f0_inicial<-function(directorio,anio,mes){

  #instalar todas las librerias necesarias para el proceso

  # Lista de librerías que deseas instalar o cargar
  librerias <- c("tidyverse", "ggplot2", "dplyr","readr","readxl","stringr","tidyr","lubridate","forecast",
                 "tsoutliers","data.table","VIM","scales","kableExtra","formattable","htmltools","webshot",
                 "openxlsx","seasonal","rmarkdown","roxygen2","plotly","gt","purrr","knitr","tinytex",
                 "webshot2","installr","shiny","shinyauthr","flexdashboard","shinydashboard","shinyjs",
                 "DT","RSQLite","pool","uuid","zoo","xts","TSstudio","tseries","flextable","writexl")

  # Verificar si las librerías están instaladas
  librerias_faltantes <- librerias[!sapply(librerias, requireNamespace, quietly = TRUE)]

  # Instalar librerías faltantes
  if (length(librerias_faltantes) > 0) {
    install.packages(librerias_faltantes)
  }

  # Cargar todas las librerías
  lapply(librerias, require, character.only = TRUE)



  #crear la función que revisa si la carpeta existe, de lo contrario la crea
  crearCarpeta <- function(ruta) {

    if (!dir.exists(ruta)) {
      dir.create(ruta)
      mensaje <- paste("Se ha creado la carpeta:", ruta)
      print(mensaje)
    } else {
      mensaje <- paste("La carpeta", ruta, "ya existe.")
      print(mensaje)
    }
  }
  #crear la carpeta results
  ruta=paste0(directorio,"/results")
  crearCarpeta(ruta)

  #crear la carpeta de integracion
  ruta=paste0(directorio,"/results/S1_integracion")
  crearCarpeta(ruta)

  #crear la carpeta data
  ruta=paste0(directorio,"/data")
  crearCarpeta(ruta)
  
  #crear la carpeta del año
  ruta=paste0(directorio,"/data/",anio)
  crearCarpeta(ruta)
  
  #crear la carpeta del mes
  ruta=paste0(directorio,"/data/",anio,"/",meses[mes])
  crearCarpeta(ruta)

  #crear la carpeta de alertas
  ruta=paste0(directorio,"/results/S2_identificacion_alertas")
  crearCarpeta(ruta)

  #crear la carpeta de imputacion
  ruta=paste0(directorio,"/results/S3_imputacion")
  crearCarpeta(ruta)


  #crear la carpeta de tematica
  ruta=paste0(directorio,"/results/S4_tematica")
  crearCarpeta(ruta)


  #crear la carpeta anexos
  ruta=paste0(directorio,"/results/S5_anexos")
  crearCarpeta(ruta)

  #crear la carpeta boletin
  ruta=paste0(directorio,"/results/S6_boletin")
  crearCarpeta(ruta)

  para_boletin <- data.frame(parametro = c("IC_prod","IC_ven","IC_empl","TNR","TI_prod","TI_ven","TI_empl","Anio_grafico"),
                             valores = c(98.1,98.1,98.5,2,2.6,2.5,2.7,2018))
  variables_iniciales <- list(Var_inicial=c("NORDEST","ANIO","MES","NOVEDAD","NOMBRE_ESTABLECIMIENTO",
                                            "IMP","ID_ESTADO","NORDEMP","NIT","DEPARTAMENTO","CLASE_CIIU4",
                                            "DOMINIO_39","DOMINIO39_DESCRIP","NOM_CODSEDE","NPERS_EP",
                                            "SUELD_EP","AJU_SUELD_EP","FECHA_INI_EP","FECHA_FIN_EP",
                                            "NPERS_ET","SUELD_ET","AJU_SUELD_ET","FECHA_INI_ET",
                                            "FECHA_FIN_ET","NPERS_ETA","SUELD_ETA","AJU_SUELD_ETA",
                                            "FECHA_INI_ETA","FECHA_FIN_ETA","NPERS_APREA","SUELD_APREA",
                                            "AJU_SUELD_APREA","FECHA_INI_APREA","FECHA_FIN_APREA",
                                            "TOTPERS_ADM","TOTSUELD_ADM","NPERS_OP","SUELD_OP",
                                            "AJU_SUELD_OP","FECHA_INI_OP","FECHA_FIN_OP","NPERS_OT",
                                            "SUELD_OT","AJU_SUELD_OT","FECHA_INI_OT","FECHA_FIN_OT",
                                            "NPERS_OTA","SUELD_OTA","AJU_SUELD_OTA","FECHA_INI_OTA",
                                            "FECHA_FIN_OTA","NPERS_APREO","SUELD_APREO","AJU_SUELD_APREO",
                                            "FECHA_INI_APREO","FECHA_FIN_APREO","TOTPERS_PRO",
                                            "TOTSUELD_PRO","TOTPERS","HORAS_ORDI","AJU_HORAS_ORDI",
                                            "HORAS_EXT","AJU_HORAS_EXT","FECHA_INI_HORAS",
                                            "FECHA_FIN_HORAS","PRODUCCION","AJU_PRODUCCION","VENTASIN",
                                            "AJU_VENTASIN","VENTASEX","AJU_VENTASEX","FECHA_INI_PYV",
                                            "FECHA_FIN_PYV","TOTAL_VENTAS","AJU_TOTAL_VENTAS",
                                            "EXISTENCIAS","ID_MUNICIPIO","NOMBREMPIO"))
  Var_inicial <- list(variables_iniciales)

  #Descargar archivo necesario
  
  url_excel <- "https://github.com/sub-dane/EMMET/raw/main/festivos.zip"
  
  # Definir el nombre y la ubicación del archivo de Excel descargado
  archivo_excel <- file.path(directorio, "data/festivos.zip")
  
  # Descargar el archivo de Excel desde GitHub
  download.file(url_excel, destfile = archivo_excel)
  unzip(archivo_excel, exdir = file.path(directorio, "data/festivos"))  
  library(openxlsx)

  wb <- createWorkbook()
  
  # Añadir la primera hoja con datos
  addWorksheet(wb, "Parametros")
  writeData(wb,sheet = "Parametros",para_boletin)

  
  # Añadir la segunda hoja con datos
  addWorksheet(wb, "Vector")
  writeData(wb, sheet = "Vector", list(variables_iniciales))  
  
  # Escribir el vector en la segunda hoja del archivo Excel
  saveWorkbook(wb,file =  paste0(directorio,"/results/S6_boletin/parametros_boletin.xlsx"), overwrite = TRUE)
  print(paste0("Se creo el archivo parametros_boletin.xlsx en ",directorio,"/results/S6_boletin/"))
}
