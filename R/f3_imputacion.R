#' Imputación outliers
#'
#' Función para realizar la imputación por deuda y por casos especiales.
#'
#' @param mes Definir las tres primeras letras del mes a ejecutar, ej: 11
#' @param anio Definir el año a ejecutar, ej: 2022
#' @param directorio definir el directorio donde se encuentran ubicado los datos de entrada
#' @param avance Denifir que porcentaje de la base va a ser cargada para el proceso de imputación, por defecto el valor
#' esta en 100.
#'
#' @details La metodología de imputación es la siguiente:
#'
#' 1.	Para las variables de capítulo 2 la metodología a usar es la imputación por el mes anterior.
#'
#' 2.	Para las variables de capítulo 3 la metodología a usar para imputación por casos especiales
#'  es el KNN combinado con variación mes anterior. Para la imputación por deuda en variables pertenecientes
#'  a capítulo 3 se realiza un proceso similar.
#'
#' El KNN combinado consta de los siguientes pasos:
#'
#' 1. Realizar un primer KNN imputando los individuos atípicos en las variables de interés (KNN1).
#'
#' 2. Calcular la variación con respecto al mes anterior.
#'
#' 3. Realizar un segundo KNN para la variación con respecto al mes anterior de los
#' establecimientos del mismo dominio (KNN2).
#'
#' 4. Calcular el valor final de la siguiente manera:
#' \deqn{ \hat{y} = KNN_1*(1+KNN_2)}
#' Donde \eqn{ \hat{y}} es el valor resultante con el que se imputara la variable de interés.
#'
#' Para el caso de imputación deuda en capítulo 3 se sigue la siguiente formula:
#' \deqn{ \hat{y} = MA*(1+KNN_2)}
#' Esto es que el valor resultante con el que se imputara la variable de interés es igual al valor
#' del mes anterior, multiplicado por 1 más la variación con respecto al mes anterior de los establecimientos
#' del mismo dominio.
#'
#' Para conocer más sobre la imputación por el método k-Nearest Neighbors ver \code{\link[VIM:kNN]{KNN_VIM}}.
#'
#' @return CSV file
#' @export
#'
#' @examples f3_imputacion(directorio="Documents/DANE/Procesos DIMPE /PilotoEMMET",
#'                        mes=11,anio=2022,avance=100)


# Función identificación e imputación de outliers ------------------------------------------



f3_imputacion <- function(directorio,mes,anio,avance=100) {
  
  
  # librerias ---------------------------------------------------------------
  
  
  library(dplyr)
  library(readxl)
  library(readr)
  library(data.table)
  library(VIM)
  library(forecast)
  library(openxlsx)
  library(imputeTS)
  source("https://raw.githubusercontent.com/NataliArteaga/DANE.EMMET/main/R/utils.R")
  
  
  #cargar base estandarizada
  datos <- read.csv(paste0(directorio,"/results/S1_integracion/EMMET_PANEL_trabajo_original_",meses[mes],anio,".csv"),fileEncoding = "latin1")
  
  #crear una copia de la base de datos
  datoscom=datos
  #crear un data frame con las variables de interes
  datos=datos %>% select(ANIO,MES,NOVEDAD,DEPARTAMENTO,NOMBREMPIO,
                         NORDEST,NOMBRE_ESTABLECIMIENTO,DOMINIO_39,CLASE_CIIU4,
                         NPERS_EP,AJU_SUELD_EP,NPERS_ET,
                         AJU_SUELD_ET,NPERS_ETA,
                         AJU_SUELD_ETA,NPERS_APREA,AJU_SUELD_APREA,
                         NPERS_OP,AJU_SUELD_OP,NPERS_OT,
                         AJU_SUELD_OT,NPERS_OTA,
                         AJU_SUELD_OTA,NPERS_APREO,AJU_SUELD_APREO,
                         AJU_HORAS_ORDI,AJU_HORAS_EXT,
                         AJU_PRODUCCION,AJU_VENTASIN,
                         AJU_VENTASEX,EXISTENCIAS)
  #convertir la base en data frame y convertir variables de año y mes en numéricas
  datos<- as.data.frame(datos)
  datos$MES=as.numeric(datos$MES)
  datos$ANIO=as.numeric(datos$ANIO)
  for (i in variablesinte) {
    datos[,i] <- as.numeric(datos[,i])
    datos[,i] <- ifelse(is.na(datos[,i]),0,datos[,i])
  }
  #Dejar la base sin los datos del mes actual
  datos <- filter(datos, !(ANIO == anio & MES == mes))
  #cargar la base de alertas
  if(avance==100){
    wowimp=fread(paste0(directorio,"/results/S2_identificacion_alertas/EMMET_PANEL_alertas_",meses[mes],anio,".csv"),encoding = "Latin-1")
    
  }else{
    wowimp=fread(paste0(directorio,"/results/S2_identificacion_alertas/EMMET_PANEL_alertas_",meses[mes],anio,"_",avance,".csv"),encoding = "Latin-1")
    
  }
  wowimp=as.data.frame(wowimp)
  for (i in variablesinte) {
    wowimp[,i] <- as.numeric(wowimp[,i])
    wowimp[,i] <- ifelse(is.na(wowimp[,i]),0,wowimp[,i])
  }
  # Convertir los datos que son casos de imputación en NA
  for (i in variablesinte) {
    wowimp[!grepl("continua", tolower(wowimp[, paste0(i, "_caso_de_imputacion")])),i]<- NA
  }
  
  wowimp=wowimp %>% select(ANIO,MES,NOVEDAD,DEPARTAMENTO,NOMBREMPIO,
                           NORDEST,NOMBRE_ESTABLECIMIENTO,DOMINIO_39,CLASE_CIIU4,
                           NPERS_EP,AJU_SUELD_EP,NPERS_ET,
                           AJU_SUELD_ET,NPERS_ETA,
                           AJU_SUELD_ETA,NPERS_APREA,AJU_SUELD_APREA,
                           NPERS_OP,AJU_SUELD_OP,NPERS_OT,
                           AJU_SUELD_OT,NPERS_OTA,
                           AJU_SUELD_OTA,NPERS_APREO,AJU_SUELD_APREO,
                           AJU_HORAS_ORDI,AJU_HORAS_EXT,
                           AJU_PRODUCCION,AJU_VENTASIN,
                           AJU_VENTASEX,EXISTENCIAS)
  #Combinar la base sin el mes actual con la base de alertas
  base_imputar=rbind(datos,wowimp)
  
  
  # imputacion ---------------------------------------------------------------------
  
  vec=c(0,0,0,0,1,1,0)
  
  
  #base_imputar=rbind(datos,wowimp)
  base_imputar2=base_imputar %>% filter((MES>=mes & ANIO==anio-2) |(ANIO==anio-1)|(MES<=mes & ANIO==anio))
  
  set.seed(11)
  for (i in cap3) {
    ids=wowimp[which(!(complete.cases(wowimp[,paste0(i)]))),"NORDEST"]
    for (j in ids) {
      
      mmm <- base_imputar2 %>% filter(NORDEST==j) %>%
        as.data.frame()
      
      model_p <- auto.arima(as.ts(mmm[,i]))
      pru_mod=(model_p$arma==vec)
      if(sum(pru_mod)<7){
        model <- model_p
        pred <- forecast(model, h = 1)$mean
      }else{
        kalman <- na_kalman((mmm[,i]))
        pred <-tail(kalman,1)
      }
      
      
      
      base_imputar2[base_imputar2$NORDEST==j & base_imputar2$ANIO==anio & base_imputar2$MES==mes,paste0(i)]=pred[1]
      
    }
    
  }
  
  
  arima_2anios=base_imputar2 %>% select(ANIO,MES,NOVEDAD,DEPARTAMENTO,NOMBREMPIO,
                                        NORDEST,NOMBRE_ESTABLECIMIENTO,DOMINIO_39,CLASE_CIIU4,
                                        NPERS_EP,AJU_SUELD_EP,NPERS_ET,
                                        AJU_SUELD_ET,NPERS_ETA,
                                        AJU_SUELD_ETA,NPERS_APREA,AJU_SUELD_APREA,
                                        NPERS_OP,AJU_SUELD_OP,NPERS_OT,
                                        AJU_SUELD_OT,NPERS_OTA,
                                        AJU_SUELD_OTA,NPERS_APREO,AJU_SUELD_APREO,
                                        AJU_HORAS_ORDI,AJU_HORAS_EXT,
                                        AJU_PRODUCCION,AJU_VENTASIN,
                                        AJU_VENTASEX,EXISTENCIAS) %>% filter(MES==mes & ANIO==anio) %>%
    arrange(NORDEST)
  
  
  
  for (i in cap3) {
    arima_2anios[,i]<- ifelse(arima_2anios[,i] < 0, 0, arima_2anios[,i])
  }
  
  #crear una base con los valores del mes anterior
  tra <- base_imputar  %>% group_by(NORDEST) %>%
    mutate_at(c(variablesinte),.funs=list(rezago1=~lag(.)))#,
  
  #convierte el NORDEST en factor
  tra$NORDEST <- as.factor(tra$NORDEST)
  
  mes_ant <- tra  %>% filter(MES==mes & ANIO==anio)
  
  mes_ant=as.data.frame(mes_ant)
  
  
  for (j in cap2) {
    #identificar las filas que tienen NA
    missing_values <- is.na(mes_ant[, j])
    
    #imputar las variables del capítulo 2 por el valor del mes anterior
    mes_ant[missing_values, j] <- round(mes_ant[missing_values, paste0(j,"_rezago1")])
    
  }
  
  #filtrar solo por el mes de interés y dejar las variables importantes
  mes_ant=mes_ant%>% filter(MES==mes & ANIO==anio) %>%
    arrange(NORDEST)
  
  mes_ant=mes_ant %>% select(ANIO,MES,NOVEDAD,DEPARTAMENTO,NOMBREMPIO,
                             NORDEST,NOMBRE_ESTABLECIMIENTO,DOMINIO_39,CLASE_CIIU4,
                             NPERS_EP,AJU_SUELD_EP,NPERS_ET,
                             AJU_SUELD_ET,NPERS_ETA,
                             AJU_SUELD_ETA,NPERS_APREA,AJU_SUELD_APREA,
                             NPERS_OP,AJU_SUELD_OP,NPERS_OT,
                             AJU_SUELD_OT,NPERS_OTA,
                             AJU_SUELD_OTA,NPERS_APREO,AJU_SUELD_APREO,
                             AJU_HORAS_ORDI,AJU_HORAS_EXT,
                             AJU_PRODUCCION,AJU_VENTASIN,
                             AJU_VENTASEX,EXISTENCIAS,ends_with("_rezago1"))
  
  
  #quitar las columnas del capítulo 3 para imputarlos por el valor obtenido con los KNN
  mes_ant=arrange(mes_ant,mes_ant$NORDEST)
  mes_ant <- mes_ant[, !(names(mes_ant) %in% cap3)]
  mes_ant <- cbind(mes_ant, arima_2anios[, cap3])
  
  for (i in cap3) {
    mes_ant[,i]=ifelse(mes_ant[,paste0(i,"_rezago1")]==0,0,mes_ant[,i])
  }
  
  #crear una base con el mes de interes y ordenar por el NORDEST
  novtem=filter(datoscom, (ANIO == anio & MES == mes))
  novtem=arrange(novtem,novtem$NORDEST)
  
  #convertir algunas variables en numericas
  novtem$NPERS_EP<-as.numeric(novtem$NPERS_EP)
  novtem$AJU_VENTASEX<-as.numeric(novtem$AJU_VENTASEX)
  novtem$AJU_HORAS_EXT<-as.numeric(novtem$AJU_HORAS_EXT)
  novtem=as.data.frame(novtem)
  
  #Modificar la base copia para que no tenga el mes de interes
  datoscom <- filter(datoscom, !(ANIO == anio & MES == mes))
  datoscom<- as.data.frame(datoscom)
  datoscom$MES=as.numeric(datoscom$MES)
  datoscom$ANIO=as.numeric(datoscom$ANIO)
  for (i in variablesinte) {
    datoscom[,i] <- as.numeric(datoscom[,i])
    datoscom[,i] <- ifelse(is.na(datoscom[,i]),0,datoscom[,i])
  }
  
  #validar si hay 0 personas entonces el sueldo debe ser 0
  for (i in 1:8) {
    mes_ant[,paste0(personas[i],"_prueba")]=ifelse((mes_ant[,personas[i]]==0 &  mes_ant[,sueldos[i]]!=0) ,1,0)
    mes_ant[,paste0(sueldos[i],"_prueba")]=ifelse((mes_ant[,personas[i]]!=0 &  (mes_ant[,sueldos[i]]/mes_ant[,personas[i]])<1000) ,1,0)
    
  }
  
  #arreglar las columnas que no cumplen ocn la regla
  for (i in 1:8) {
    
    mes_ant[mes_ant[,paste0(personas[i],"_prueba")]==1,paste0(sueldos[i])]<- 0
    
  }
  
  mes_ant[,"horasordi_prueba"]=ifelse((mes_ant$AJU_HORAS_ORDI/(24*(mes_ant$NPERS_OP+mes_ant$NPERS_OT+mes_ant$NPERS_OTA+mes_ant$NPERS_APREO))
  )<7 | (mes_ant$AJU_HORAS_ORDI/(24*(mes_ant$NPERS_OP+mes_ant$NPERS_OT+mes_ant$NPERS_OTA+mes_ant$NPERS_APREO))
  )>9,1,0)
  
  mes_ant[,"horasextras_prueba"]=ifelse(mes_ant$AJU_HORAS_EXT>mes_ant$AJU_HORAS_ORDI,1,0)
  
  mes_ant[,"existencias_prueba"]=ifelse((mes_ant[,"AJU_PRODUCCION"]>(mes_ant[,"AJU_VENTASEX"]+mes_ant[,"AJU_VENTASIN"]) & mes_ant[,"EXISTENCIAS"]<mes_ant[,"EXISTENCIAS_rezago1"])|(mes_ant[,"AJU_PRODUCCION"]<(mes_ant[,"AJU_VENTASEX"]+mes_ant[,"AJU_VENTASIN"]) & mes_ant[,"EXISTENCIAS"]>mes_ant[,"EXISTENCIAS_rezago1"]),1,0)
  
  #pegar las columnas imputadas en el data frame del mes actual
  for (i in variablesinte){
    novtem[,i]<- mes_ant[,i]
    
  }
  
  
  #combinar las base original estandarizada con el mes imputado
  
  imputa=rbind(datoscom,novtem)
  imputa=as.data.frame(imputa)
  for (i in variablesinte) {
    imputa[,i] <- as.numeric(imputa[,i])
  }
  imputa <- imputa %>% mutate_if(is.integer, as.numeric)
  
  imputa <- imputa %>%
    mutate(
      TOTAL_VENTAS=(AJU_VENTASIN+AJU_VENTASEX),
      TOTAL_HORAS= (AJU_HORAS_ORDI+AJU_HORAS_EXT),
      TOT_PERS=(NPERS_EP+NPERS_ET+NPERS_ETA+
                  NPERS_APREA+NPERS_OP+NPERS_OT+
                  NPERS_OTA+NPERS_APREO)
    )
  
  # Exportar la base imputada -------------------------------
  write.csv(mes_ant,paste0(directorio,"/results/S3_imputacion/EMMET_reglas_consistencia_",meses[mes],anio,".csv"),row.names=F,fileEncoding = "latin1")
  
  write.csv(imputa,paste0(directorio,"/results/S3_imputacion/EMMET_PANEL_imputada_",meses[mes],anio,".csv"),row.names=F,fileEncoding = "latin1")
}

