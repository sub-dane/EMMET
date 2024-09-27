#' Tematica
#'
#' @description
#'  Esta funcion construye la base tematica. Esta base expone los datos
#'  procesados, de acuerdo a la metodología de la operación, en ese sentido
#'  presenta los datos reales a partir de los nominales, sumado a ellos presenta
#'  la información ponderada y agrega variables de la identificacion de los
#'  dominios por los cuales se publca. Con esto, se puede presentar la base
#'  tematica como la base final de cuadros.
#'  Para construir esta base, se usa como insumo la base Panel con los casos
#'  de imputacion aplicados, y la base deflactor que contiene la información para presentar los valores
#'  reales según IPP_PyC,IPP_EXP,IPC. Finalmente exporta un archivo de extension .xlsx .
#'
#'  @param mes Definir las tres primeras letras del mes a ejecutar, ej: 11
#'  @param anio Definir el año a ejecutar, ej: 2022
#'  @param directorio definir el directorio donde se encuentran ubicado los datos de entrada
#'
#'
#'
#'
#' @return CSV file
#' @export
#'
#' @examples f4_tematica(directorio="Documents/DANE/Procesos DIMPE /PilotoEMMET",
#'                        mes=11,anio=2022)
#'
#'
#' @details
#'  Esta funcion crea las variables de la base tematica de la siguiente manera:
#'
#'  VentasReales: Suma de, la división entre ventas en el interior e IPP_PYC y,
#'     la división entre ventas en el exterior e IPP_PYC
#'
#'  TotalPersonas: Suma de las variables que contabilizan el número de
#'     trabajadores en las diferentes categorías.
#'
#'  TotalSueldosNominal: Suma de las variables de sueldos nominales
#'
#'  TotalSueldosReales: División entre TotalSueldosNominal e IPC
#'
#'  SueldosPermanentesNominal: Suma de las variables de sueldos nominales del
#'     personal contratado por la categoria permanente
#'
#'  SueldosPermanentesReales: División entre SueldosPermanentesNominal e IPC
#'
#'  SueldosTemporalesNominal: Suma de las variables de sueldos nominales del
#'     personal contratado por la categoria temporales
#'
#'  SueldosTemporalesReales: División entre SueldosTemporalesNominal e IPC
#'
#'  SueldosAdmonNominal: Suma de las variables de sueldo nominal del
#'     personal administrativo nominales
#'
#'  SueldosAdmonReal: División entre SueldosAdmonNominal e IPC
#'
#'  SueldosAdmonPermReal: División entre AJU_II_PA_PP_SUELD_EP e IPC
#'
#'  SueldosAdmonTempNomin: Suma de las variables de sueldo del
#'     personal administrativos temporales
#'
#'  SueldosAdmonTempReal: División entre SueldosAdmonTempNomin e IPC
#'
#'  SueldosProducNominal: Suma de las variables del personal de
#'     producción nominal
#'
#'  SueldosProducReal: División entre SueldosProducNominal e IPC
#'
#'  SueldosProducPermReal: División entre AJU_II_PP_PP_SUELD_OP e IPC.
#'
#'  SueldosProducTempNomin: Suma de las variables del personal de
#'     producción temporal
#'
#'  SueldosProducTempReal: División entre SueldosProducTempNomin e IPC.
#'
#'  TotalHoras: Suma de varaibles de horas extras y ordinarias
#'
#'  TotalEmpleoPermanente: Suma de varaibles del personal de
#'     administración y producción
#'
#'  TotalEmpleoTemporal: Suma de varaibles del personal temporal
#'
#'  TotalEmpleoAdmon: Suma de varaibles del personal administrativo
#'
#'  TotalEmpleoProduc: Suma de varaibles del personal de producción
#'
#'  EmpleoProducTempor: Suma de varaibles del personal de producción
#'     temporal
#'
#'  TOTAL_VENTAS: Suma de varaibles en el interior y exterior de país
#'
#'  DEFLACTOR_NAL: División entre TOTAL_VENTAS y VentasReales
#'




f4_tematica <- function(directorio,mes,anio){
  # Cargar librerC-as --------------------------------------------------------
  library(readxl)
  library(dplyr)
  library(ggplot2)
  library(tidyr)
  library(scales)
  library(kableExtra)
  library(lubridate)
  library(formattable)
  library("htmltools")
  library("webshot")
  library(openxlsx)
  library(seasonal)
  library(stringr)
  source("https://raw.githubusercontent.com/sub-dane/EMMET/main/R/utils.R")



  # Cargar bases y variables ------------------------------------------------

  base_panel2<-read.csv(paste0(directorio,"/results/S3_imputacion/EMMET_PANEL_imputada_",meses[mes],anio,".csv"),fileEncoding = "latin1")
  base_deflactor             <- read_excel(paste0(directorio,"/data/",anio,"/",meses[mes],"/DEFLACTOR_",meses[mes],anio,".xlsx"))
  colnames(base_deflactor)   <- colnames_format(base_deflactor)


  # Concatenar con Base Deflactor ----------------------------------------------------
  base_panel2 <- base_panel2 %>%
    left_join(base_deflactor,by=c("CLASE_CIIU4"="CIIU4","ANIO"="ANO","MES"="MES"))


  # Modificar tipo de variables ---------------------------------------------



  base_panel2 <- as.data.frame(base_panel2)
  base_panel2$MES=as.numeric(base_panel2$MES)
  base_panel2$ANIO=as.numeric(base_panel2$ANIO)
  # variablesinte=c("NPERS_EP","AJU_SUELD_EP","NPERS_ET","AJU_SUELD_ET",
  #                 "NPERS_ETA","AJU_SUELD_ETA","NPERS_APREA","AJU_SUELD_APREA",
  #                 "NPERS_OP","AJU_SUELD_OP","NPERS_OT","AJU_SUELD_OT",
  #                 "NPERS_OTA","AJU_SUELD_OTA","NPERS_APREO","AJU_SUELD_APREO",
  #                 "AJU_HORAS_ORDI","AJU_HORAS_EXT",
  #                 "AJU_PRODUCCION","AJU_VENTASIN","AJU_VENTASEX","EXISTENCIAS")

  for (i in variablesinte) {

    base_panel2[,i] <- as.numeric(base_panel2[,i])
    base_panel2[,i] <- ifelse(is.na(base_panel2[,i]),0,base_panel2[,i])
  }



  base_panel2$NPERS_EP<-as.numeric(base_panel2$NPERS_EP)
  base_panel2$AJU_VENTASEX<-as.numeric(base_panel2$AJU_VENTASEX)
  base_panel2$AJU_HORAS_EXT<-as.numeric(base_panel2$AJU_HORAS_EXT)
  base_panel2$PONDERADOR<-gsub(",",".",base_panel2$PONDERADOR)
  base_panel2$PONDERADOR<-as.numeric(base_panel2$PONDERADOR)






  # Funcion -----------------------------------------------------------------


  #Creacion de todas las variables
  base_tematica<-function(base){
    base_tematica<-base %>%
      select(ANIO,MES,NOVEDAD,ID_ESTADO,NORDEST,DEPARTAMENTO,NPERS_EP,
             AJU_SUELD_EP,NPERS_ET,AJU_SUELD_ET,
             NPERS_ETA,AJU_SUELD_ETA,NPERS_APREA,
             AJU_SUELD_APREA,NPERS_OP,AJU_SUELD_OP,
             NPERS_OT,AJU_SUELD_OT,NPERS_OTA,
             AJU_SUELD_OTA,NPERS_APREO,AJU_SUELD_APREO,TOTPERS,
             AJU_HORAS_ORDI,AJU_HORAS_EXT,TOTAL_HORAS,
             AJU_PRODUCCION,AJU_VENTASIN,AJU_VENTASEX,
             TOTAL_VENTAS,ID_MUNICIPIO,ORDEN_AREA,
             AREA_METROPOLITANA,ORDEN_CIUDAD,CIUDAD,ORDEN_DEPTO,
             INCLUSION_NOMBRE_DEPTO,NOMBRE_ESTABLECIMIENTO,CLASE_CIIU4,DIVISION,EMMET_CLASE,
             ORDENDOMINDEPTO,AGREG_DOMINIO_REG,DOMINIO_43,DOMINIO_39,
             DOMINIO39_DESCRIP,DOM_PONDERADOR_CIUDADES,PONDERADOR,
             DEFLACTOR,IPP_PYC,IPP_EXP,IPC,DEPARTAMENTO) %>%
      mutate(VentasReales=round((AJU_VENTASIN/IPP_PYC)+(AJU_VENTASEX/IPP_EXP),8),
             TotalSueldosNominal=round((AJU_SUELD_EP+AJU_SUELD_ET+
                                          AJU_SUELD_ETA+AJU_SUELD_APREA+
                                          AJU_SUELD_OP+AJU_SUELD_OT+
                                          AJU_SUELD_OTA+AJU_SUELD_APREO),8),
             TotalSueldosReales=round((TotalSueldosNominal/IPC),8),
             SueldosPermanentesNominal=round((AJU_SUELD_EP+AJU_SUELD_OP),8),
             SueldosPermanentesReales= round((SueldosPermanentesNominal/IPC),8),
             SueldosTemporalesNominal= round((AJU_SUELD_ET+AJU_SUELD_ETA+
                                                AJU_SUELD_APREA+AJU_SUELD_OT+
                                                AJU_SUELD_OTA+AJU_SUELD_APREO),8),
             SueldosTemporalesReales= round((SueldosTemporalesNominal/IPC),8),
             SueldosAdmonNominal= round((AJU_SUELD_EP+AJU_SUELD_ET+
                                           AJU_SUELD_ETA+AJU_SUELD_APREA),8),
             SueldosAdmonReal= round((SueldosAdmonNominal/IPC),8),
             SueldosAdmonPermReal= round((AJU_SUELD_EP/IPC),8),
             SueldosAdmonTempNomin= round((AJU_SUELD_ET+AJU_SUELD_ETA+
                                             AJU_SUELD_APREA),8),
             SueldosAdmonTempReal= round((SueldosAdmonTempNomin/IPC),8),
             SueldosProducNominal=round((AJU_SUELD_OP+AJU_SUELD_OT+
                                           AJU_SUELD_OTA+AJU_SUELD_APREO),8),
             SueldosProducReal = round((SueldosProducNominal/IPC),8),
             SueldosProducPermReal= round((AJU_SUELD_OP/IPC),8),
             SueldosProducTempNomin=round((AJU_SUELD_OT+AJU_SUELD_OTA+
                                             AJU_SUELD_APREO),8),
             SueldosProducTempReal = round((SueldosProducTempNomin/IPC),8),
             TotalHoras= round((AJU_HORAS_ORDI+AJU_HORAS_EXT),8),
             TotalEmpleoPermanente= round((NPERS_EP+NPERS_OP),8),
             TotalEmpleoTemporal= round((NPERS_ET+NPERS_ETA+NPERS_APREA+
                                           NPERS_OT+NPERS_OTA+NPERS_APREO),8),
             TotalEmpleoAdmon= round((NPERS_EP+NPERS_ET+NPERS_ETA+
                                        NPERS_APREA),8),
             EmpleoAdmonPerman=round((NPERS_EP),8),
             EmpleoAdmonTempor=round((NPERS_ET+NPERS_ETA+NPERS_APREA),8),
             TotalEmpleoProduc= round((NPERS_OP+NPERS_OT+NPERS_OTA+
                                         NPERS_APREO),8),
             EmpleoProducPerman=round((NPERS_OP),8),
             EmpleoProducTempor= round((NPERS_OT+NPERS_OTA+NPERS_APREO),8),
             TOTAL_VENTAS=round((AJU_VENTASIN+AJU_VENTASEX),8),
             DEFLACTOR_NAL=(TOTAL_VENTAS/VentasReales))
    base_tematica<- base_tematica %>%
      mutate(deflactor_nal=case_when(
        is.na(DEFLACTOR_NAL)==TRUE ~ IPP_PYC,
        (VentasReales==0 & TOTAL_VENTAS==0) ~ IPP_PYC,
        TRUE ~ DEFLACTOR_NAL
      ))

    base_tematica<-base_tematica %>%
      mutate(ProduccionReal= round((AJU_PRODUCCION/deflactor_nal),10),
             ProduccionNomPond= round((AJU_PRODUCCION*PONDERADOR),10),
             ProduccionRealPond= round((ProduccionReal*PONDERADOR),10),
             VentasNominPond= round(((AJU_VENTASIN+AJU_VENTASEX)*PONDERADOR),10),
             VentasRealesPond= round((VentasReales*PONDERADOR),10))
    base_tematica=subset(base_tematica,select=-DEFLACTOR_NAL)


    return(base_tematica)
  }



  base_tematica<-base_tematica(base_panel2)


  write.csv(base_tematica,paste0(directorio,"/results/S4_tematica/EMMET_PANEL_tematica_",meses[mes],anio,".csv"),row.names=F,fileEncoding ="latin1")

}
