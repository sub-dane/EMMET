#'  Integración
#'
#'  Con esta función se integraran los archivos de entrada que son necesarios en el proceso EMMET, con el fin
#'  de trabajar con una base llamada "base panel.csv". Adicionalmente, en las variables numeticas, los datos faltantes se cambian por cero.
#'  A estas mismas variables se es establece un formato numerico para la porterior manipulación.
#'
#'
#' @param mes Definir el mes a ejecutar, ej: 11
#' @param anio Definir el año a ejecutar, ej: 2022
#' @param directorio definir el directorio donde se encuentran ubicado los datos de entrada
#'
#' @return CSV file
#' @export
#'
#' @details En esta función se leen las siguientes bases:
#'
#' 1 Base logística: Es la base con todos los datos desde enero de 2018 hasta el mes
#'  indicado en el parámetro.
#'
#' 2 Base paramétrica: Base con las características de cada establecimiento.
#'
#'
#' El procedimiento para unificar las bases es el siguiente:
#'
#' 1 Se cambia el formato de la base logistica de xlsx a csv renombrandola como base panel.
#'
#' 2 Se agregan las columnas de la base paramétrica a la base panel por medio del ID_NUMORD.
#'
#'
#'
#'
#' @examples f1_integracion(directorio="Documents/DANE/Procesos DIMPE /PilotoEMMET",
#'                        mes=11,anio=2022)


f1_integracion <- function(directorio,
                        mes,
                        anio){
  error_dire <- "El valor de 'directorio' debe ser proporcionado.Ej: \"Documents/DANE/Procesos DIMPE/PilotoEMMET\"."

  if(missingArg(directorio)) stop(error_dire,call.=FALSE)

  if (!is.character(directorio)) {
    error_message <- sprintf("El valor de 'directorio' debe ser un caracter: %s no es válido. Ej:\"Documents/DANE/Procesos DIMPE /PilotoEMMET\".", directorio)
    stop(error_message,call. = FALSE)
  }
  error_anio<-"El valor de 'anio' debe ser numerico.Ej: 2022"
  if(missingArg(anio)) stop(error_anio,call.=FALSE)

  if (!is.numeric(anio)) {
    error_message <- sprintf("El valor de 'anio' debe ser numerico: %s no es válido.Ej: 2022", anio)
    stop(error_message,call. = FALSE)
  }
 error_mes<-"El valor de 'mes' debe ser numerico.Ej: 11"
  if(missingArg(mes)) stop(error_mes,call.=FALSE)

  if (!is.numeric(mes)) {
    error_message <- sprintf("El valor de 'mes' debe ser numerico: %s no es válido.Ej: 11", mes)
    stop(error_message,call. = FALSE)
  }

  # librerias ---------------------------------------------------------------

 library(readr)
 library(readxl)
 library(stringr)
 library(tidyr)
 library(dplyr)
 library(data.table)
 library(writexl)
 source("https://raw.githubusercontent.com/NataliArteaga/DANE.EMMET/main/R/utils.R")




 # Cargar bases necesarias y aplicar funcion de formato de nombres

 #Cambio: se agrea , sheet = "LOGISTICA"
 base_logistica          <- read.xlsx(paste0(directorio,"/data/",anio,"/",meses[mes],"/EMMET_PANEL_imputada_",meses[mes],anio,".xlsx"))
 colnames(base_logistica) <- colnames_format(base_logistica)

 #cambia en las columnas que contienen la palabra "OBSE" cualquier caracter que no sea alfanumerico por un espacio
 base_logistica           <-  base_logistica %>%
   mutate_at(vars(contains("OBSER")),~str_replace_all(.,pattern="[^[:alnum:]]",replacement=" "))

 base_logistica           <-  base_logistica %>%
   mutate_at(vars("DOMINIO39_DESCRIP"),~str_replace_all(.,pattern="[^[:alnum:]]",replacement=" "))

 base_logistica           <-  base_logistica %>%
   mutate_at(vars("NOMBRE_ESTABLECIMIENTO"),~str_replace_all(.,pattern="[^[:alnum:]]",replacement=" "))

 base_logistica           <-  base_logistica %>%
   mutate_at(vars("DEPARTAMENTO"),~str_replace_all(.,pattern="[^[:alnum:]]",replacement=" "))

 parametro <- read.xlsx(paste0(directorio,"/results/S6_boletin/parametros_boletin.xlsx"), sheet = "Vector")

 base_logistica           <-  base_logistica %>%
   select(parametro$Var_inicial)

 write.csv(base_logistica,paste0(directorio,"/data/",anio,"/",meses[mes],"/EMMET_PANEL_imputada_",meses[mes],anio,".csv"),row.names=F)

 base_logistica          <-  read_csv(paste0(directorio,"/data/",anio,"/",meses[mes],"/EMMET_PANEL_imputada_",meses[mes],anio,".csv"))
 colnames(base_logistica) <- colnames_format(base_logistica)


 #Cargar bases insumo

 base_parametrica           <- read.xlsx(paste0(directorio,"/data/",anio,"/",meses[mes],"EMMET_parametrica_historico.xlsx"))
 colnames(base_parametrica) <- colnames_format(base_parametrica)

 # Concatenar base Logistica con base Original ----------------------------------------------------

 #base_panel <- rbind.data.frame(base_logistica,base_original)
 base_panel <- base_logistica

 # Concatenar con Base Parametrica ---------------------------------------------
 base_panel <- base_panel %>%
   left_join(base_parametrica %>% select(!c(NOMBRE_ESTABLECIMIENTO,NOVEDAD,CLASE_CIIU4,
                                            DOMINIO_39,DOMINIO39_DESCRIP,ID_MUNICIPIO,
                                            NOMBREMPIO)),
             by=c("NORDEST"="NORDEST","ANIO"="ANIO","MES"="MES"))

 # Estandarizacion nombres Departamento y Municipio ------------------------------------------------
 base_panel=base_panel %>%
   mutate_at(vars("DOMINIO39_DESCRIP"),~str_replace_all(.,pattern="[^[:alnum:]]",replacement=" "))

 # Estandarizar Variables Numericas ----------------------------------------

 base_panel<-base_panel %>%
   mutate(
     AJU_VENTASIN=ifelse(is.na(AJU_VENTASIN),0,AJU_VENTASIN),
     AJU_VENTASEX=ifelse(is.na(AJU_VENTASEX),0,AJU_VENTASEX),
     NPERS_EP=ifelse(is.na(NPERS_EP),0,NPERS_EP),
     NPERS_ET=ifelse(is.na(NPERS_ET),0,NPERS_ET),
     NPERS_ETA=ifelse(is.na(NPERS_ETA),0,NPERS_ETA),
     NPERS_APREA=ifelse(is.na(NPERS_APREA),0,NPERS_APREA),
     NPERS_OP=ifelse(is.na(NPERS_OP),0,NPERS_OP),
     NPERS_OT=ifelse(is.na(NPERS_OT),0,NPERS_OT),
     NPERS_OTA=ifelse(is.na(NPERS_OTA),0,NPERS_OTA),
     NPERS_APREO=ifelse(is.na(NPERS_APREO),0,NPERS_APREO),
     TOTPERS=ifelse(is.na(TOTPERS),0,TOTPERS),
     TOTPERS_ADM=ifelse(is.na(TOTPERS_ADM),0,TOTPERS_ADM),
     TOTSUELD_ADM=ifelse(is.na(TOTSUELD_ADM),0,TOTSUELD_ADM),
     TOTSUELD_PRO=ifelse(is.na(TOTSUELD_PRO),0,TOTSUELD_PRO),
     AJU_SUELD_EP=ifelse(is.na(AJU_SUELD_EP),0,AJU_SUELD_EP),
     AJU_SUELD_ET=ifelse(is.na(AJU_SUELD_ET),0,AJU_SUELD_ET),
     AJU_SUELD_ETA=ifelse(is.na(AJU_SUELD_ETA),0,AJU_SUELD_ETA),
     AJU_SUELD_APREA=ifelse(is.na(AJU_SUELD_APREA),0,AJU_SUELD_APREA),
     AJU_SUELD_OP=ifelse(is.na(AJU_SUELD_OP),0,AJU_SUELD_OP),
     AJU_SUELD_OT=ifelse(is.na(AJU_SUELD_OT),0,AJU_SUELD_OT),
     AJU_SUELD_OTA=ifelse(is.na(AJU_SUELD_OTA),0,AJU_SUELD_OTA),
     AJU_SUELD_APREO=ifelse(is.na(AJU_SUELD_APREO),0,AJU_SUELD_APREO),
     AJU_HORAS_ORDI=ifelse(is.na(AJU_HORAS_ORDI),0,AJU_HORAS_ORDI),
     AJU_HORAS_EXT=ifelse(is.na(AJU_HORAS_EXT),0,AJU_HORAS_EXT),
     AJU_PRODUCCION=ifelse(is.na(AJU_PRODUCCION),0,AJU_PRODUCCION),
     EXISTENCIAS=ifelse(is.na(EXISTENCIAS),0,EXISTENCIAS),
     AJU_TOTAL_VENTAS=ifelse(is.na(AJU_TOTAL_VENTAS),0,AJU_TOTAL_VENTAS)
   )

 # Exportar bases de Datos integradas -------------------------------------------------

 write.csv(base_panel,paste0(directorio,"/results/S1_integracion/EMMET_PANEL_trabajo_original_",meses[mes],anio,".csv"),row.names=F,fileEncoding = "latin1")

}
