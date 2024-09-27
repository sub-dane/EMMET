#' Cuadros desagregaion regiones
#'
#' @description Esta funcion crea un archivo de excel que contiene las variaciones
#' y contribuciones anueales desagregadas por cada región. Se crea
#' una hoja por cada desagregación en donde se consigna las variaciones y contribuciones de las variables
#' de produccion, ventas, emppleados, sueldos y horas.
#' 
#' 
#'  
#' @param mes Definir el mes a ejecutar, ej: 11
#' @param anio Definir el año a ejecutar, ej: 2022
#' @param directorio definir el directorio donde se encuentran ubicado los datos de entrada
#'
#' @return CSV file
#' @export
#'
#' @examples f8_cregiones(directorio="Documents/DANE/Procesos DIMPE /PilotoEMMET",
#'                        mes=11,anio=2022)
#'
#' @details Esta funcion se hace con la finalidad de crear cuadros informativos 
#' con datos desagregada según el interés, para facilitar el análisis del comportamiento
#' de cada región.
#' 
#'  


f8_cregiones<-function(directorio,anio,mes){
  # Librerias ---------------------------------------------------------------
  
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
  library(DT)
  library(data.table)
  library(stringr)
  
  source("https://raw.githubusercontent.com/sub-dane/EMMET/main/R/utils.R")
  
  # Cargar bases y variables ------------------------------------------------
  
  meses <- c("ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic")

  
  data<-read.csv(paste0(directorio,"/results/S4_tematica/EMMET_PANEL_tematica_",meses[mes],anio,".csv"),fileEncoding = "latin1")
  colnames(data) <- colnames_format(data)
  
  num_cols <- sapply(data, is.numeric)
  data[num_cols] <- lapply(data[num_cols], function(x) { x[is.na(x)] <- 0; x })
  
  data <- data %>% 
    group_by(ANIO,MES,ID_NUMORD,DOMINIO_39) %>% 
    mutate(PERSONAL=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL))
  
  nomb_estab <- data %>% 
    filter((ANIO==anio & MES==mes)) %>% 
    select(ID_NUMORD,NOMBRE_ESTAB,CLASE_CIIU4,NOVEDAD) 
  
  
  if(mes==1){
    mes_inte <- 12
    anio_inte <- anio-1
  }else{
    mes_inte <- mes-1
    anio_inte <- anio
  }
  
  # Crear un nuevo libro de Excel -------------------------------------------
  
  wb <- openxlsx::createWorkbook()
  Salida<-paste0(directorio,"/results/S5_anexos/cuadros_territoriales_",meses[mes],"_",anio,".xlsx")
  
  # Define el formato de número con un decimal para la columna de números
  num_formato <- createStyle(numFmt = "0.0")
  #Formato de negrilla
  bold_style <- createStyle(textDecoration = "bold")
  # Totales dept ------------------------------------------------------------
  
  
  #Funcion
  
  Tot_depto <- function(var_inte){
    agreg_inte <- data %>%
      select(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,{{var_inte}})
    
    
    prod_cont <- agreg_inte %>%
      group_by(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO) %>%
      summarise(PRODREAL=sum({{var_inte}}))
    
    
    prod_cont <- prod_cont %>%
      pivot_wider(names_from = ANIO, values_from = PRODREAL)
    
    T_IND <- data %>%
      select(ANIO,MES,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(ORDEN_DEPTO=0,
                INCLUSION_NOMBRE_DEPTO="Total Industria",
                TOTAL=sum({{var_inte}}))
    
    
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    prod_cont <- rbind(prod_cont,T_IND)
    
    n=ncol(prod_cont)
    for(i in 4:ncol(prod_cont)){
      if(i==n){
        prod_cont[paste0("var_",names(prod_cont)[i])] <- prod_cont[,i]/prod_cont[,i-1]*100-100
      }else{
        prod_cont[paste0("var_",names(prod_cont)[i+1])] <- prod_cont[,i+1]/prod_cont[,i]*100-100
      }
    }
    
    prod_cont_ori <- prod_cont %>%
      select(MES,INCLUSION_NOMBRE_DEPTO,!matches("^var"))
    n=ncol(prod_cont_ori)
    
    T_IND <- agreg_inte %>%
      select(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(TOTAL=sum({{var_inte}}))
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    colnames(T_IND) <- paste0(names(T_IND),"_T")
    
    prod_cont_ori <- prod_cont_ori %>%
      left_join(T_IND, by=c("MES"="MES_T"))
    
    
    for(i in 4:n){
      if(i==n){
        a <- prod_cont_ori[,i]-prod_cont_ori[,i-1]
        b <- prod_cont_ori[,(i+(anio-2018))]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i])]  <- (a/b)*100
      }else{
        a <- prod_cont_ori[,i+1]-prod_cont_ori[,i]
        b <- prod_cont_ori[,(i+(anio-2018)+1)]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i+1])]  <- (a/b)*100
      }
    }
    
    
    prod_cont_ori <- prod_cont_ori %>%
      select(MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,matches("^cont"))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_longer(cols = 4:ncol(prod_cont_ori),
                   names_to = "CONT_ANIO", values_to = "CONTRIBUCION")
    
    prod_cont_ori <- prod_cont_ori %>%
      arrange(CONT_ANIO)
    
    prod_cont_ori["ANIO_MES"] <- paste0(prod_cont_ori$CONT_ANIO,"_",prod_cont_ori$MES)
    
    prod_cont_ori<- subset(prod_cont_ori, select = -c(CONT_ANIO,MES))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_wider(names_from = ANIO_MES, values_from = CONTRIBUCION)
    
    
    prod_cont_var <- prod_cont %>%
      select(MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,matches("^var"))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_longer(cols = 4:ncol(prod_cont_var),
                   names_to = "VAR_ANIO", values_to = "VARIACION")
    
    prod_cont_var <- prod_cont_var %>%
      arrange(VAR_ANIO)
    
    prod_cont_var["ANIO_MES"] <- paste0(prod_cont_var$VAR_ANIO,"_",prod_cont_var$MES)
    
    prod_cont_var<- subset(prod_cont_var, select = -c(VAR_ANIO,MES))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_wider(names_from = ANIO_MES, values_from = VARIACION)
    
    prod_cont_fin <- prod_cont_var %>%
      left_join(prod_cont_ori, by=c("ORDEN_DEPTO"="ORDEN_DEPTO",
                                    "INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
    
    
    prod_cont_fin <- prod_cont_fin[,c("ORDEN_DEPTO",
                                      "INCLUSION_NOMBRE_DEPTO",
                                      paste0("var_",anio_inte,"_",mes_inte),
                                      paste0("var_",anio,"_",mes),
                                      paste0("cont_",anio_inte,"_",mes_inte),
                                      paste0("cont_",anio,"_",mes))]
    
    prod_cont_fin <- prod_cont_fin %>% 
      arrange(ORDEN_DEPTO) %>% 
      rename(Orden=ORDEN_DEPTO,
             Dominio=INCLUSION_NOMBRE_DEPTO)
    
  }
  
  prod_depto <- Tot_depto(PRODUCCIONREALPOND)
  prod_depto["DIFERENCIA"] <- prod_depto[,ncol(prod_depto)]-prod_depto[,(ncol(prod_depto)-1)]
  colnames(prod_depto) <- paste0(names(prod_depto),"_prod")
  
  vent_depto <- Tot_depto(VENTASREALESPOND)
  vent_depto["DIFERENCIA"] <- vent_depto[,ncol(vent_depto)]-vent_depto[,(ncol(vent_depto)-1)]
  colnames(vent_depto) <- paste0(names(vent_depto),"_vent")
  
  emp_depto <- Tot_depto(PERSONAL)
  emp_depto["DIFERENCIA"] <- emp_depto[,ncol(emp_depto)]-emp_depto[,(ncol(emp_depto)-1)]
  colnames(emp_depto) <- paste0(names(emp_depto),"_emp")
  
  sueld_depto <- Tot_depto(TOTALSUELDOSREALES)
  sueld_depto["DIFERENCIA"] <- sueld_depto[,ncol(sueld_depto)]-sueld_depto[,(ncol(sueld_depto)-1)]
  colnames(sueld_depto) <- paste0(names(sueld_depto),"_sueld")
  
  horas_depto <- Tot_depto(TOTALHORAS)
  horas_depto["DIFERENCIA"] <- horas_depto[,ncol(horas_depto)]-horas_depto[,(ncol(horas_depto)-1)]
  colnames(horas_depto) <- paste0(names(horas_depto),"_horas")
  
  
  Tot_dep <- prod_depto %>% 
    left_join(vent_depto,by=c("Orden_prod"="Orden_vent",
                              "Dominio_prod"="Dominio_vent")) %>% 
    left_join(emp_depto,by=c("Orden_prod"="Orden_emp",
                             "Dominio_prod"="Dominio_emp")) %>% 
    left_join(sueld_depto,by=c("Orden_prod"="Orden_sueld",
                               "Dominio_prod"="Dominio_sueld")) %>% 
    left_join(horas_depto,by=c("Orden_prod"="Orden_horas",
                               "Dominio_prod"="Dominio_horas"))
  
  
  addWorksheet(wb, sheetName = "Totales dept")
  
  for (i in 3:ncol(Tot_dep)){
    # Aplica el formato a la columna de números (columna A en este caso)
    addStyle(wb, sheet = "Totales dept", style = num_formato, 
             rows = 2:(nrow(Tot_dep) + 1), cols = i)
  }
  openxlsx::addStyle(wb, sheet ="Totales dept", row = 1, col = 1:ncol(Tot_dep), style = bold_style)
  
  
  # Escribir la base de datos en la primera hoja del archivo Excel
  writeData(wb, sheet = "Totales dept", x = Tot_dep)
  
  
  # Totales domin -----------------------------------------------------------
  
  
  Tot_dom <- function(var_inte){
    agreg_inte <- data %>%
      select(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
             ORDENDOMINDEPTO,AGREG_DOMINIO_REG,{{var_inte}})
    
    
    prod_cont <- agreg_inte %>%
      group_by(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
               ORDENDOMINDEPTO,AGREG_DOMINIO_REG) %>% 
      summarise(PRODREAL=sum({{var_inte}}))
    
    
    prod_cont <- prod_cont %>%
      pivot_wider(names_from = ANIO, values_from = PRODREAL)
    
    T_IND_1 <- data %>%
      select(ANIO,MES,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(ORDEN_DEPTO=0,
                INCLUSION_NOMBRE_DEPTO="T_IND",
                ORDENDOMINDEPTO=0,
                AGREG_DOMINIO_REG="Total Industria",
                TOTAL=sum({{var_inte}}))
    
    
    T_IND_1$ANIO <- as.character(T_IND_1$ANIO)
    
    T_IND_1 <- T_IND_1 %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    prod_cont <- rbind(prod_cont,T_IND_1)
    
    n=ncol(prod_cont)
    for(i in 6:ncol(prod_cont)){
      if(i==n){
        prod_cont[paste0("var_",names(prod_cont)[i])] <- prod_cont[,i]/prod_cont[,i-1]*100-100
      }else{
        prod_cont[paste0("var_",names(prod_cont)[i+1])] <- prod_cont[,i+1]/prod_cont[,i]*100-100
      }
    }
    
    prod_cont_ori <- prod_cont %>%
      select(MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
             ORDENDOMINDEPTO,AGREG_DOMINIO_REG,!matches("^var"))
    n=ncol(prod_cont_ori)
    
    T_IND <- data %>%
      select(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
             ORDENDOMINDEPTO,AGREG_DOMINIO_REG,{{var_inte}}) %>%
      group_by(ANIO,MES,INCLUSION_NOMBRE_DEPTO) %>%
      summarise(TOTAL=sum({{var_inte}}))
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    T_IND <- rbind(T_IND,T_IND_1 %>% select(names(T_IND)))
    
    colnames(T_IND) <- paste0(names(T_IND),"_T")
    
    prod_cont_ori <- prod_cont_ori %>%
      left_join(T_IND, by=c("MES"="MES_T",
                            "INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO_T"))
    
    
    for(i in 6:n){
      if(i==n){
        a <- prod_cont_ori[,i]-prod_cont_ori[,i-1]
        b <- prod_cont_ori[,(i+(anio-2018))]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i])]  <- (a/b)*100
      }else{
        a <- prod_cont_ori[,i+1]-prod_cont_ori[,i]
        b <- prod_cont_ori[,(i+(anio-2018)+1)]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i+1])]  <- (a/b)*100
      }
    }
    
    
    prod_cont_ori <- prod_cont_ori %>%
      select(MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
             ORDENDOMINDEPTO,AGREG_DOMINIO_REG,matches("^cont"))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_longer(cols = 6:ncol(prod_cont_ori),
                   names_to = "CONT_ANIO", values_to = "CONTRIBUCION")
    
    prod_cont_ori <- prod_cont_ori %>%
      arrange(CONT_ANIO)
    
    prod_cont_ori["ANIO_MES"] <- paste0(prod_cont_ori$CONT_ANIO,"_",prod_cont_ori$MES)
    
    prod_cont_ori<- subset(prod_cont_ori, select = -c(CONT_ANIO,MES))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_wider(names_from = ANIO_MES, values_from = CONTRIBUCION)
    
    
    prod_cont_var <- prod_cont %>%
      select(MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
             ORDENDOMINDEPTO,AGREG_DOMINIO_REG,matches("^var"))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_longer(cols = 6:ncol(prod_cont_var),
                   names_to = "VAR_ANIO", values_to = "VARIACION")
    
    prod_cont_var <- prod_cont_var %>%
      arrange(VAR_ANIO)
    
    prod_cont_var["ANIO_MES"] <- paste0(prod_cont_var$VAR_ANIO,"_",prod_cont_var$MES)
    
    prod_cont_var<- subset(prod_cont_var, select = -c(VAR_ANIO,MES))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_wider(names_from = ANIO_MES, values_from = VARIACION)
    
    prod_cont_fin <- prod_cont_var %>%
      left_join(prod_cont_ori, by=c("ORDEN_DEPTO"="ORDEN_DEPTO",
                                    "INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO",
                                    "ORDENDOMINDEPTO"="ORDENDOMINDEPTO",
                                    "AGREG_DOMINIO_REG"="AGREG_DOMINIO_REG"))
    
    
    prod_cont_fin <- prod_cont_fin[,c("ORDEN_DEPTO","INCLUSION_NOMBRE_DEPTO",
                                      "ORDENDOMINDEPTO","AGREG_DOMINIO_REG",
                                      paste0("var_",anio_inte,"_",mes_inte),
                                      paste0("var_",anio,"_",mes),
                                      paste0("cont_",anio_inte,"_",mes_inte),
                                      paste0("cont_",anio,"_",mes))]
    
    
    prod_cont_fin <- prod_cont_fin %>% 
      arrange(ORDEN_DEPTO,ORDENDOMINDEPTO) %>% 
      rename(N_Depto=ORDEN_DEPTO,
             Departamento=INCLUSION_NOMBRE_DEPTO,
             N_Dominio=ORDENDOMINDEPTO,
             Dominio=AGREG_DOMINIO_REG)
  }
  
  
  prod_dom <- Tot_dom(PRODUCCIONREALPOND)
  prod_dom["DIFERENCIA"] <- prod_dom[,ncol(prod_dom)]-prod_dom[,(ncol(prod_dom)-1)]
  colnames(prod_dom) <- paste0(names(prod_dom),"_prod")
  
  vent_dom <- Tot_dom(VENTASREALESPOND)
  vent_dom["DIFERENCIA"] <- vent_dom[,ncol(vent_dom)]-vent_dom[,(ncol(vent_dom)-1)]
  colnames(vent_dom) <- paste0(names(vent_dom),"_vent")
  
  
  emp_dom <- Tot_dom(PERSONAL)
  emp_dom["DIFERENCIA"] <- emp_dom[,ncol(emp_dom)]-emp_dom[,(ncol(emp_dom)-1)]
  colnames(emp_dom) <- paste0(names(emp_dom),"_emp")
  
  sueld_dom <- Tot_dom(TOTALSUELDOSREALES)
  sueld_dom["DIFERENCIA"] <- sueld_dom[,ncol(sueld_dom)]-sueld_dom[,(ncol(sueld_dom)-1)]
  colnames(sueld_dom) <- paste0(names(sueld_dom),"_sueld")
  
  horas_dom <- Tot_dom(TOTALHORAS)
  horas_dom["DIFERENCIA"] <- horas_dom[,ncol(horas_dom)]-horas_dom[,(ncol(horas_dom)-1)]
  colnames(horas_dom) <- paste0(names(horas_dom),"_horas")
  
  
  
  Total_dom <- prod_dom %>% 
    left_join(vent_dom,by=c("N_Depto_prod"="N_Depto_vent",
                            "Departamento_prod"="Departamento_vent",
                            "N_Dominio_prod"="N_Dominio_vent",
                            "Dominio_prod"="Dominio_vent")) %>% 
    left_join(emp_dom,by=c("N_Depto_prod"="N_Depto_emp",
                           "Departamento_prod"="Departamento_emp",
                           "N_Dominio_prod"="N_Dominio_emp",
                           "Dominio_prod"="Dominio_emp")) %>% 
    left_join(sueld_dom,by=c("N_Depto_prod"="N_Depto_sueld",
                             "Departamento_prod"="Departamento_sueld",
                             "N_Dominio_prod"="N_Dominio_sueld",
                             "Dominio_prod"="Dominio_sueld")) %>% 
    left_join(horas_dom,by=c("N_Depto_prod"="N_Depto_horas",
                             "Departamento_prod"="Departamento_horas",
                             "N_Dominio_prod"="N_Dominio_horas",  
                             "Dominio_prod"="Dominio_horas"))
  
  addWorksheet(wb, sheetName = "Totales domin")
  
  for (i in 5:ncol(Total_dom)){
    # Aplica el formato a la columna de números (columna A en este caso)
    addStyle(wb, sheet = "Totales domin", style = num_formato, 
             rows = 2:(nrow(Total_dom) + 1), cols = i)
  }
  openxlsx::addStyle(wb, sheet ="Totales domin", row = 1, col = 1:ncol(Total_dom), style = bold_style)
  
  # Escribir la base de datos en la primera hoja del archivo Excel
  writeData(wb, sheet = "Totales domin", x = Total_dom)
  
  
  # Por estable -------------------------------------------------------------
  
  
  Estab <- function(var_inte){
    agreg_inte <- data %>%
      select(ANIO,MES,ID_NUMORD,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
             ORDENDOMINDEPTO,AGREG_DOMINIO_REG,{{var_inte}})
    
    
    prod_cont <- agreg_inte %>%
      group_by(ANIO,MES,ID_NUMORD,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
               ORDENDOMINDEPTO,AGREG_DOMINIO_REG) %>%
      summarise(PRODREAL=sum({{var_inte}}))
    
    
    prod_cont <- prod_cont %>%
      pivot_wider(names_from = ANIO, values_from = PRODREAL)
    
    prod_cont <- replace(prod_cont,is.na(prod_cont),0)
    
    T_IND_0 <- data %>%
      select(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
             ORDENDOMINDEPTO,AGREG_DOMINIO_REG,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(ID_NUMORD=0,
                ORDEN_DEPTO=0,
                INCLUSION_NOMBRE_DEPTO="T_IND",
                ORDENDOMINDEPTO=0,
                AGREG_DOMINIO_REG="Total_Industria",
                TOTAL=sum({{var_inte}}))
    
    
    T_IND_0$ANIO <- as.character(T_IND_0$ANIO)
    
    T_IND_0 <- T_IND_0 %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    prod_cont <- rbind(prod_cont,T_IND_0)
    
    n=ncol(prod_cont)
    for(i in 7:ncol(prod_cont)){
      if(i==n){
        prod_cont[paste0("var_",names(prod_cont)[i])] <- prod_cont[,i]/prod_cont[,i-1]*100-100
      }else{
        prod_cont[paste0("var_",names(prod_cont)[i+1])] <- prod_cont[,i+1]/prod_cont[,i]*100-100
      }
    }
    
    prod_cont_ori <- prod_cont %>%
      select(!matches("^var"))
    n=ncol(prod_cont_ori)
    
    T_IND <- data %>%
      select(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
             ORDENDOMINDEPTO,AGREG_DOMINIO_REG,{{var_inte}}) %>%
      group_by(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
               ORDENDOMINDEPTO,AGREG_DOMINIO_REG) %>%
      summarise(TOTAL=sum({{var_inte}}))
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    T_IND <- rbind(T_IND,(T_IND_0 %>% select(!ID_NUMORD)))
    
    colnames(T_IND) <- paste0(names(T_IND),"_T")
    
    prod_cont_ori <- prod_cont_ori %>%
      left_join(T_IND, by=c("MES"="MES_T",
                            "ORDEN_DEPTO"="ORDEN_DEPTO_T",
                            "ORDENDOMINDEPTO"="ORDENDOMINDEPTO_T",
                            "INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO_T",
                            "AGREG_DOMINIO_REG"="AGREG_DOMINIO_REG_T"))
    
    for(i in 7:n){
      if(i==n){
        a <- prod_cont_ori[,i]-prod_cont_ori[,i-1]
        b <- prod_cont_ori[,(i+(anio-2018))]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i])]  <- (a/b)*100
      }else{
        a <- prod_cont_ori[,i+1]-prod_cont_ori[,i]
        b <- prod_cont_ori[,(i+(anio-2018)+1)]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i+1])]  <- (a/b)*100
      }
    }
    
    
    prod_cont_ori <- prod_cont_ori %>%
      select(MES,ID_NUMORD,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
             ORDENDOMINDEPTO,AGREG_DOMINIO_REG,matches("^cont"))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_longer(cols = 7:ncol(prod_cont_ori),
                   names_to = "CONT_ANIO", values_to = "CONTRIBUCION")
    
    prod_cont_ori <- prod_cont_ori %>%
      arrange(CONT_ANIO)
    
    prod_cont_ori["ANIO_MES"] <- paste0(prod_cont_ori$CONT_ANIO,"_",prod_cont_ori$MES)
    
    prod_cont_ori<- subset(prod_cont_ori, select = -c(CONT_ANIO,MES))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_wider(names_from = ANIO_MES, values_from = CONTRIBUCION)
    
    
    prod_cont_var <- prod_cont %>%
      select(MES,ID_NUMORD,INCLUSION_NOMBRE_DEPTO,AGREG_DOMINIO_REG,matches("^var"))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_longer(cols = 7:ncol(prod_cont_var),
                   names_to = "VAR_ANIO", values_to = "VARIACION")
    
    prod_cont_var <- prod_cont_var %>%
      arrange(VAR_ANIO)
    
    prod_cont_var["ANIO_MES"] <- paste0(prod_cont_var$VAR_ANIO,"_",prod_cont_var$MES)
    
    prod_cont_var<- subset(prod_cont_var, select = -c(VAR_ANIO,MES))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_wider(names_from = ANIO_MES, values_from = VARIACION)
    
    prod_cont_fin <- prod_cont_var %>%
      left_join(prod_cont_ori, by=c("ID_NUMORD"="ID_NUMORD",
                                    "ORDEN_DEPTO"="ORDEN_DEPTO",
                                    "INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO",
                                    "ORDENDOMINDEPTO"="ORDENDOMINDEPTO",
                                    "AGREG_DOMINIO_REG"="AGREG_DOMINIO_REG"))
    
    
    prod_cont_fin <- prod_cont_fin[,c("ID_NUMORD",
                                      "ORDEN_DEPTO","INCLUSION_NOMBRE_DEPTO",
                                      "ORDENDOMINDEPTO","AGREG_DOMINIO_REG",
                                      paste0("var_",anio_inte,"_",mes_inte),
                                      paste0("var_",anio,"_",mes),
                                      paste0("cont_",anio_inte,"_",mes_inte),
                                      paste0("cont_",anio,"_",mes))]
    
    
    prod_cont_fin <- prod_cont_fin %>%
      arrange(ORDEN_DEPTO,ORDENDOMINDEPTO)
    
  }
  
  
  prod_estab <- Estab(PRODUCCIONREALPOND)
  prod_estab["DIFERENCIA"] <- prod_estab[,ncol(prod_estab)]-prod_estab[,(ncol(prod_estab)-1)]
  colnames(prod_estab) <- paste0(names(prod_estab),"_prod")
  
  vent_estab <- Estab(VENTASREALESPOND)
  vent_estab["DIFERENCIA"] <- vent_estab[,ncol(vent_estab)]-vent_estab[,(ncol(vent_estab)-1)]
  colnames(vent_estab) <- paste0(names(vent_estab),"_vent")
  
  
  emp_estab <- Estab(PERSONAL)
  emp_estab["DIFERENCIA"] <- emp_estab[,ncol(emp_estab)]-emp_estab[,(ncol(emp_estab)-1)]
  colnames(emp_estab) <- paste0(names(emp_estab),"_emp")
  
  sueld_estab <- Estab(TOTALSUELDOSREALES)
  sueld_estab["DIFERENCIA"] <- sueld_estab[,ncol(sueld_estab)]-sueld_estab[,(ncol(sueld_estab)-1)]
  colnames(sueld_estab) <- paste0(names(sueld_estab),"_sueld")
  
  horas_estab <- Estab(TOTALHORAS)
  horas_estab["DIFERENCIA"] <- horas_estab[,ncol(horas_estab)]-horas_estab[,(ncol(horas_estab)-1)]
  colnames(horas_estab) <- paste0(names(horas_estab),"_horas")
  
  
  Por_estable <- prod_estab %>% 
    left_join(vent_estab,by=c("ID_NUMORD_prod"="ID_NUMORD_vent",
                              "ORDEN_DEPTO_prod"="ORDEN_DEPTO_vent",
                              "INCLUSION_NOMBRE_DEPTO_prod"="INCLUSION_NOMBRE_DEPTO_vent",
                              "ORDENDOMINDEPTO_prod"="ORDENDOMINDEPTO_vent",
                              "AGREG_DOMINIO_REG_prod"="AGREG_DOMINIO_REG_vent")) %>% 
    left_join(emp_estab,by=c("ID_NUMORD_prod"="ID_NUMORD_emp",
                             "ORDEN_DEPTO_prod"="ORDEN_DEPTO_emp",
                             "INCLUSION_NOMBRE_DEPTO_prod"="INCLUSION_NOMBRE_DEPTO_emp",
                             "ORDENDOMINDEPTO_prod"="ORDENDOMINDEPTO_emp",
                             "AGREG_DOMINIO_REG_prod"="AGREG_DOMINIO_REG_emp")) %>% 
    left_join(sueld_estab,by=c("ID_NUMORD_prod"="ID_NUMORD_sueld",
                               "ORDEN_DEPTO_prod"="ORDEN_DEPTO_sueld",
                               "INCLUSION_NOMBRE_DEPTO_prod"="INCLUSION_NOMBRE_DEPTO_sueld",
                               "ORDENDOMINDEPTO_prod"="ORDENDOMINDEPTO_sueld",
                               "AGREG_DOMINIO_REG_prod"="AGREG_DOMINIO_REG_sueld")) %>% 
    left_join(horas_estab,by=c("ID_NUMORD_prod"="ID_NUMORD_horas",
                               "ORDEN_DEPTO_prod"="ORDEN_DEPTO_horas",
                               "INCLUSION_NOMBRE_DEPTO_prod"="INCLUSION_NOMBRE_DEPTO_horas",
                               "ORDENDOMINDEPTO_prod"="ORDENDOMINDEPTO_horas",
                               "AGREG_DOMINIO_REG_prod"="AGREG_DOMINIO_REG_horas"))
  
  n_por_est <- colnames(Por_estable)
  
  Por_estable <- Por_estable %>% 
    left_join(nomb_estab %>% select(!c(ANIO,MES)), by=c("ID_NUMORD_prod"="ID_NUMORD"))
  
  Por_estable <- Por_estable %>% 
    select(ID_NUMORD_prod,NOMBRE_ESTAB,NOVEDAD,CLASE_CIIU4,
           DOMINIO_39,c(n_por_est[2:length(n_por_est)])) %>% 
    arrange(ORDEN_DEPTO_prod,ORDENDOMINDEPTO_prod,ID_NUMORD_prod) %>% 
    rename(Norden=ID_NUMORD_prod,
           Nombre=NOMBRE_ESTAB,
           Nov=NOVEDAD,
           EMMET_Clae=CLASE_CIIU4,
           Dom_39=DOMINIO_39,
           N_Depto=ORDEN_DEPTO_prod,
           Departamento=INCLUSION_NOMBRE_DEPTO_prod,
           N_Dominio=ORDENDOMINDEPTO_prod,
           DominioDpto=AGREG_DOMINIO_REG_prod) 
  
  
  addWorksheet(wb, sheetName = "Por estable")
  
  for (i in 10:ncol(Por_estable)){
    # Aplica el formato a la columna de números (columna A en este caso)
    addStyle(wb, sheet = "Por estable", style = num_formato, 
             rows = 2:(nrow(Por_estable) + 1), cols = i)
  }
  openxlsx::addStyle(wb, sheet ="Por estable", row = 1, col = 1:ncol(Por_estable), style = bold_style)
  
  
  # Escribir la base de datos en la primera hoja del archivo Excel
  writeData(wb, sheet = "Por estable", x = Por_estable)
  
  
  # Areas met ---------------------------------------------------------------
  
  
  Area_met <- function(var_inte){
    agreg_inte <- data %>%
      select(ANIO,MES,ORDEN_AREA,AREA_METROPOLITANA,{{var_inte}})
    
    
    prod_cont <- agreg_inte %>%
      group_by(ANIO,MES,ORDEN_AREA,AREA_METROPOLITANA) %>%
      summarise(PRODREAL=sum({{var_inte}}))
    
    
    prod_cont <- prod_cont %>%
      pivot_wider(names_from = ANIO, values_from = PRODREAL)
    
    T_IND <- data %>%
      select(ANIO,MES,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(ORDEN_AREA=0,
                AREA_METROPOLITANA="Total Industria",
                TOTAL=sum({{var_inte}}))
    
    
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    prod_cont <- rbind(prod_cont,T_IND)
    
    n=ncol(prod_cont)
    for(i in 4:ncol(prod_cont)){
      if(i==n){
        prod_cont[paste0("var_",names(prod_cont)[i])] <- prod_cont[,i]/prod_cont[,i-1]*100-100
      }else{
        prod_cont[paste0("var_",names(prod_cont)[i+1])] <- prod_cont[,i+1]/prod_cont[,i]*100-100
      }
    }
    
    prod_cont_ori <- prod_cont %>%
      select(MES,ORDEN_AREA,AREA_METROPOLITANA,!matches("^var"))
    n=ncol(prod_cont_ori)
    
    T_IND <- agreg_inte %>%
      select(ANIO,MES,ORDEN_AREA,AREA_METROPOLITANA,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(TOTAL=sum({{var_inte}}))
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    colnames(T_IND) <- paste0(names(T_IND),"_T")
    
    prod_cont_ori <- prod_cont_ori %>%
      left_join(T_IND, by=c("MES"="MES_T"))
    
    
    for(i in 4:n){
      if(i==n){
        a <- prod_cont_ori[,i]-prod_cont_ori[,i-1]
        b <- prod_cont_ori[,(i+(anio-2018))]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i])]  <- (a/b)*100
      }else{
        a <- prod_cont_ori[,i+1]-prod_cont_ori[,i]
        b <- prod_cont_ori[,(i+(anio-2018)+1)]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i+1])]  <- (a/b)*100
      }
    }
    
    
    prod_cont_ori <- prod_cont_ori %>%
      select(MES,ORDEN_AREA,AREA_METROPOLITANA,matches("^cont"))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_longer(cols = 4:ncol(prod_cont_ori),
                   names_to = "CONT_ANIO", values_to = "CONTRIBUCION")
    
    prod_cont_ori <- prod_cont_ori %>%
      arrange(CONT_ANIO)
    
    prod_cont_ori["ANIO_MES"] <- paste0(prod_cont_ori$CONT_ANIO,"_",prod_cont_ori$MES)
    
    prod_cont_ori<- subset(prod_cont_ori, select = -c(CONT_ANIO,MES))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_wider(names_from = ANIO_MES, values_from = CONTRIBUCION)
    
    
    prod_cont_var <- prod_cont %>%
      select(MES,ORDEN_AREA,AREA_METROPOLITANA,matches("^var"))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_longer(cols = 4:ncol(prod_cont_var),
                   names_to = "VAR_ANIO", values_to = "VARIACION")
    
    prod_cont_var <- prod_cont_var %>%
      arrange(VAR_ANIO)
    
    prod_cont_var["ANIO_MES"] <- paste0(prod_cont_var$VAR_ANIO,"_",prod_cont_var$MES)
    
    prod_cont_var<- subset(prod_cont_var, select = -c(VAR_ANIO,MES))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_wider(names_from = ANIO_MES, values_from = VARIACION)
    
    prod_cont_fin <- prod_cont_var %>%
      left_join(prod_cont_ori, by=c("ORDEN_AREA"="ORDEN_AREA",
                                    "AREA_METROPOLITANA"="AREA_METROPOLITANA"))
    
    
    prod_cont_fin <- prod_cont_fin[,c("ORDEN_AREA","AREA_METROPOLITANA",
                                      paste0("var_",anio_inte,"_",mes_inte),
                                      paste0("var_",anio,"_",mes),
                                      paste0("cont_",anio_inte,"_",mes_inte),
                                      paste0("cont_",anio,"_",mes))]
    
    prod_cont_fin <- prod_cont_fin %>% 
      arrange(ORDEN_AREA) 
    
    
  }
  
  
  
  prod_area <- Area_met(PRODUCCIONREALPOND)
  prod_area["DIFERENCIA"] <- prod_area[,ncol(prod_area)]-prod_area[,(ncol(prod_area)-1)]
  colnames(prod_area) <- paste0(names(prod_area),"_prod")
  
  vent_area <- Area_met(VENTASREALESPOND)
  vent_area["DIFERENCIA"] <- vent_area[,ncol(vent_area)]-vent_area[,(ncol(vent_area)-1)]
  colnames(vent_area) <- paste0(names(vent_area),"_vent")
  
  
  emp_area <- Area_met(PERSONAL)
  emp_area["DIFERENCIA"] <- emp_area[,ncol(emp_area)]-emp_area[,(ncol(emp_area)-1)]
  colnames(emp_area) <- paste0(names(emp_area),"_emp")
  
  sueld_area <- Area_met(TOTALSUELDOSREALES)
  sueld_area["DIFERENCIA"] <- sueld_area[,ncol(sueld_area)]-sueld_area[,(ncol(sueld_area)-1)]
  colnames(sueld_area) <- paste0(names(sueld_area),"_sueld")
  
  horas_area <- Area_met(TOTALHORAS)
  horas_area["DIFERENCIA"] <- horas_area[,ncol(horas_area)]-horas_area[,(ncol(horas_area)-1)]
  colnames(horas_area) <- paste0(names(horas_area),"_horas")
  
  
  Area <- prod_area %>% 
    left_join(vent_area,by=c("ORDEN_AREA_prod"="ORDEN_AREA_vent",
                             "AREA_METROPOLITANA_prod"="AREA_METROPOLITANA_vent")) %>% 
    left_join(emp_area,by=c("ORDEN_AREA_prod"="ORDEN_AREA_emp",
                            "AREA_METROPOLITANA_prod"="AREA_METROPOLITANA_emp")) %>% 
    left_join(sueld_area,by=c("ORDEN_AREA_prod"="ORDEN_AREA_sueld",
                              "AREA_METROPOLITANA_prod"="AREA_METROPOLITANA_sueld")) %>% 
    left_join(horas_area,by=c("ORDEN_AREA_prod"="ORDEN_AREA_horas",
                              "AREA_METROPOLITANA_prod"="AREA_METROPOLITANA_horas"))
  
  Area <- Area %>% 
    rename(N_Dom=ORDEN_AREA_prod,
           Dominio=AREA_METROPOLITANA_prod)
  
  
  addWorksheet(wb, sheetName = "Area met")
  
  for (i in 3:ncol(Area)){
    # Aplica el formato a la columna de números (columna A en este caso)
    addStyle(wb, sheet = "Area met", style = num_formato, 
             rows = 2:(nrow(Area) + 1), cols = i)
  }
  openxlsx::addStyle(wb, sheet ="Area met", row = 1, col = 1:ncol(Area), style = bold_style)
  
  
  # Escribir la base de datos en la primera hoja del archivo Excel
  writeData(wb, sheet = "Area met", x = Area)
  
  
  # Ares est ----------------------------------------------------------------
  
  
  Area_est <- function(var_inte){
    agreg_inte <- data %>%
      select(ANIO,MES,ID_NUMORD,ORDEN_AREA,AREA_METROPOLITANA,{{var_inte}})
    
    
    prod_cont <- agreg_inte %>%
      group_by(ANIO,MES,ID_NUMORD,ORDEN_AREA,AREA_METROPOLITANA) %>%
      summarise(PRODREAL=sum({{var_inte}}))
    
    
    prod_cont <- prod_cont %>%
      pivot_wider(names_from = ANIO, values_from = PRODREAL)
    
    prod_cont <- replace(prod_cont,is.na(prod_cont),0)
    
    T_IND_0 <- data %>%
      select(ANIO,MES,ORDEN_AREA,AREA_METROPOLITANA,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(ID_NUMORD=0,
                ORDEN_AREA=0,
                AREA_METROPOLITANA="Total Industria",
                TOTAL=sum({{var_inte}}))
    
    
    T_IND_0$ANIO <- as.character(T_IND_0$ANIO)
    
    T_IND_0 <- T_IND_0 %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    prod_cont <- rbind(prod_cont,T_IND_0)
    
    n=ncol(prod_cont)
    for(i in 5:ncol(prod_cont)){
      if(i==n){
        prod_cont[paste0("var_",names(prod_cont)[i])] <- prod_cont[,i]/prod_cont[,i-1]*100-100
      }else{
        prod_cont[paste0("var_",names(prod_cont)[i+1])] <- prod_cont[,i+1]/prod_cont[,i]*100-100
      }
    }
    
    prod_cont_ori <- prod_cont %>%
      select(!matches("^var"))
    n=ncol(prod_cont_ori)
    
    T_IND <- data %>%
      select(ANIO,MES,ORDEN_AREA,AREA_METROPOLITANA,{{var_inte}}) %>%
      group_by(ANIO,MES,ORDEN_AREA,AREA_METROPOLITANA) %>%
      summarise(TOTAL=sum({{var_inte}}))
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    T_IND <- rbind(T_IND,(T_IND_0 %>% select(!ID_NUMORD)))
    
    colnames(T_IND) <- paste0(names(T_IND),"_T")
    
    prod_cont_ori <- prod_cont_ori %>%
      left_join(T_IND, by=c("MES"="MES_T",
                            "ORDEN_AREA"="ORDEN_AREA_T",
                            "AREA_METROPOLITANA"="AREA_METROPOLITANA_T"))
    
    
    for(i in 5:n){
      if(i==n){
        a <- prod_cont_ori[,i]-prod_cont_ori[,i-1]
        b <- prod_cont_ori[,(i+(anio-2018))]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i])]  <- (a/b)*100
      }else{
        a <- prod_cont_ori[,i+1]-prod_cont_ori[,i]
        b <- prod_cont_ori[,(i+(anio-2018)+1)]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i+1])]  <- (a/b)*100
      }
    }
    
    
    
    prod_cont_ori <- prod_cont_ori %>%
      select(MES,ID_NUMORD,ORDEN_AREA,AREA_METROPOLITANA,matches("^cont"))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_longer(cols = 5:ncol(prod_cont_ori),
                   names_to = "CONT_ANIO", values_to = "CONTRIBUCION")
    
    prod_cont_ori <- prod_cont_ori %>%
      arrange(CONT_ANIO)
    
    prod_cont_ori["ANIO_MES"] <- paste0(prod_cont_ori$CONT_ANIO,"_",prod_cont_ori$MES)
    
    prod_cont_ori<- subset(prod_cont_ori, select = -c(CONT_ANIO,MES))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_wider(names_from = ANIO_MES, values_from = CONTRIBUCION)
    
    
    prod_cont_var <- prod_cont %>%
      select(MES,ID_NUMORD,ORDEN_AREA,AREA_METROPOLITANA,matches("^var"))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_longer(cols = 5:ncol(prod_cont_var),
                   names_to = "VAR_ANIO", values_to = "VARIACION")
    
    prod_cont_var <- prod_cont_var %>%
      arrange(VAR_ANIO)
    
    prod_cont_var["ANIO_MES"] <- paste0(prod_cont_var$VAR_ANIO,"_",prod_cont_var$MES)
    
    prod_cont_var<- subset(prod_cont_var, select = -c(VAR_ANIO,MES))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_wider(names_from = ANIO_MES, values_from = VARIACION)
    
    prod_cont_fin <- prod_cont_var %>%
      left_join(prod_cont_ori, by=c("ID_NUMORD"="ID_NUMORD",
                                    "ORDEN_AREA"="ORDEN_AREA",
                                    "AREA_METROPOLITANA"="AREA_METROPOLITANA"))
    
    
    prod_cont_fin <- prod_cont_fin[,c("ID_NUMORD","ORDEN_AREA",
                                      "AREA_METROPOLITANA",
                                      paste0("var_",anio_inte,"_",mes_inte),
                                      paste0("var_",anio,"_",mes),
                                      paste0("cont_",anio_inte,"_",mes_inte),
                                      paste0("cont_",anio,"_",mes))]
    
    prod_cont_fin <- prod_cont_fin %>%
      arrange(ORDEN_AREA)
  }
  
  
  prod_areaest <- Area_est(PRODUCCIONREALPOND)
  prod_areaest["DIFERENCIA"] <- prod_areaest[,ncol(prod_areaest)]-prod_areaest[,(ncol(prod_areaest)-1)]
  colnames(prod_areaest) <- paste0(names(prod_areaest),"_prod")
  
  vent_areaest <- Area_est(VENTASREALESPOND)
  vent_areaest["DIFERENCIA"] <- vent_areaest[,ncol(vent_areaest)]-vent_areaest[,(ncol(vent_areaest)-1)]
  colnames(vent_areaest) <- paste0(names(vent_areaest),"_vent")
  
  
  emp_areaest <- Area_est(PERSONAL)
  emp_areaest["DIFERENCIA"] <- emp_areaest[,ncol(emp_areaest)]-emp_areaest[,(ncol(emp_areaest)-1)]
  colnames(emp_areaest) <- paste0(names(emp_areaest),"_emp")
  
  sueld_areaest <- Area_est(TOTALSUELDOSREALES)
  sueld_areaest["DIFERENCIA"] <- sueld_areaest[,ncol(sueld_areaest)]-sueld_areaest[,(ncol(sueld_areaest)-1)]
  colnames(sueld_areaest) <- paste0(names(sueld_areaest),"_sueld")
  
  horas_areaest <- Area_est(TOTALHORAS)
  horas_areaest["DIFERENCIA"] <- horas_areaest[,ncol(horas_areaest)]-horas_areaest[,(ncol(horas_areaest)-1)]
  colnames(horas_areaest) <- paste0(names(horas_areaest),"_horas")
  
  
  Area_estable <- prod_areaest %>% 
    left_join(vent_areaest,by=c("ID_NUMORD_prod"="ID_NUMORD_vent",
                                "ORDEN_AREA_prod"="ORDEN_AREA_vent",
                                "AREA_METROPOLITANA_prod"="AREA_METROPOLITANA_vent")) %>% 
    left_join(emp_areaest,by=c("ID_NUMORD_prod"="ID_NUMORD_emp",
                               "ORDEN_AREA_prod"="ORDEN_AREA_emp",
                               "AREA_METROPOLITANA_prod"="AREA_METROPOLITANA_emp")) %>% 
    left_join(sueld_areaest,by=c("ID_NUMORD_prod"="ID_NUMORD_sueld",
                                 "ORDEN_AREA_prod"="ORDEN_AREA_sueld",
                                 "AREA_METROPOLITANA_prod"="AREA_METROPOLITANA_sueld")) %>% 
    left_join(horas_areaest,by=c("ID_NUMORD_prod"="ID_NUMORD_horas",
                                 "ORDEN_AREA_prod"="ORDEN_AREA_horas",
                                 "AREA_METROPOLITANA_prod"="AREA_METROPOLITANA_horas"))
  
  
  n_area_est <- colnames(Area_estable)
  
  Area_estable <- Area_estable %>% 
    left_join(nomb_estab %>% select(!c(ANIO,MES)), by=c("ID_NUMORD_prod"="ID_NUMORD"))
  
  Area_estable <- Area_estable %>% 
    select(ID_NUMORD_prod,NOMBRE_ESTAB,NOVEDAD,CLASE_CIIU4,
           DOMINIO_39,c(n_area_est[2:length(n_area_est)])) %>%  
    rename(Norden=ID_NUMORD_prod,
           Nombre=NOMBRE_ESTAB,
           Nov=NOVEDAD,
           EMMET_Clae=CLASE_CIIU4,
           Dom_39=DOMINIO_39,
           N_Dom=ORDEN_AREA_prod,
           Area_Metropolitana=AREA_METROPOLITANA_prod)
  
  addWorksheet(wb, sheetName = "Areas est")
  
  for (i in 8:ncol(Area_estable)){
    # Aplica el formato a la columna de números (columna A en este caso)
    addStyle(wb, sheet = "Areas est", style = num_formato, 
             rows = 2:(nrow(Area_estable) + 1), cols = i)
  }
  openxlsx::addStyle(wb, sheet ="Areas est", row = 1, col = 1:ncol(Area_estable), style = bold_style)
  
  
  # Escribir la base de datos en la primera hoja del archivo Excel
  writeData(wb, sheet = "Areas est", x = Area_estable)
  
  
  # Ciudades ----------------------------------------------------------------
  
  
  Ciudad <- function(var_inte){
    
    agreg_inte <- data %>%
      select(ANIO,MES,ORDEN_CIUDAD,CIUDAD,{{var_inte}})
    
    
    prod_cont <- agreg_inte %>%
      group_by(ANIO,MES,ORDEN_CIUDAD,CIUDAD) %>%
      summarise(PRODREAL=sum({{var_inte}}))
    
    
    prod_cont <- prod_cont %>%
      pivot_wider(names_from = ANIO, values_from = PRODREAL)
    
    T_IND <- data %>%
      select(ANIO,MES,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(ORDEN_CIUDAD=0,
                CIUDAD="Total Industria",
                TOTAL=sum({{var_inte}}))
    
    
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    prod_cont <- rbind(prod_cont,T_IND)
    
    n=ncol(prod_cont)
    for(i in 4:ncol(prod_cont)){
      if(i==n){
        prod_cont[paste0("var_",names(prod_cont)[i])] <- prod_cont[,i]/prod_cont[,i-1]*100-100
      }else{
        prod_cont[paste0("var_",names(prod_cont)[i+1])] <- prod_cont[,i+1]/prod_cont[,i]*100-100
      }
    }
    
    prod_cont_ori <- prod_cont %>%
      select(MES,ORDEN_CIUDAD,CIUDAD,!matches("^var"))
    n=ncol(prod_cont_ori)
    
    T_IND <- agreg_inte %>%
      select(ANIO,MES,ORDEN_CIUDAD,CIUDAD,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(TOTAL=sum({{var_inte}}))
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    colnames(T_IND) <- paste0(names(T_IND),"_T")
    
    prod_cont_ori <- prod_cont_ori %>%
      left_join(T_IND, by=c("MES"="MES_T"))
    
    
    for(i in 4:n){
      if(i==n){
        a <- prod_cont_ori[,i]-prod_cont_ori[,i-1]
        b <- prod_cont_ori[,(i+(anio-2018))]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i])]  <- (a/b)*100
      }else{
        a <- prod_cont_ori[,i+1]-prod_cont_ori[,i]
        b <- prod_cont_ori[,(i+(anio-2018)+1)]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i+1])]  <- (a/b)*100
      }
    }
    
    
    prod_cont_ori <- prod_cont_ori %>%
      select(MES,ORDEN_CIUDAD,CIUDAD,matches("^cont"))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_longer(cols = 4:ncol(prod_cont_ori),
                   names_to = "CONT_ANIO", values_to = "CONTRIBUCION")
    
    prod_cont_ori <- prod_cont_ori %>%
      arrange(CONT_ANIO)
    
    prod_cont_ori["ANIO_MES"] <- paste0(prod_cont_ori$CONT_ANIO,"_",prod_cont_ori$MES)
    
    prod_cont_ori<- subset(prod_cont_ori, select = -c(CONT_ANIO,MES))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_wider(names_from = ANIO_MES, values_from = CONTRIBUCION)
    
    
    prod_cont_var <- prod_cont %>%
      select(MES,ORDEN_CIUDAD,CIUDAD,matches("^var"))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_longer(cols = 4:ncol(prod_cont_var),
                   names_to = "VAR_ANIO", values_to = "VARIACION")
    
    prod_cont_var <- prod_cont_var %>%
      arrange(VAR_ANIO)
    
    prod_cont_var["ANIO_MES"] <- paste0(prod_cont_var$VAR_ANIO,"_",prod_cont_var$MES)
    
    prod_cont_var<- subset(prod_cont_var, select = -c(VAR_ANIO,MES))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_wider(names_from = ANIO_MES, values_from = VARIACION)
    
    prod_cont_fin <- prod_cont_var %>%
      left_join(prod_cont_ori, by=c("ORDEN_CIUDAD"="ORDEN_CIUDAD",
                                    "CIUDAD"="CIUDAD"))
    
    
    prod_cont_fin <- prod_cont_fin[,c("ORDEN_CIUDAD","CIUDAD",
                                      paste0("var_",anio_inte,"_",mes_inte),
                                      paste0("var_",anio,"_",mes),
                                      paste0("cont_",anio_inte,"_",mes_inte),
                                      paste0("cont_",anio,"_",mes))]
    
    prod_cont_fin <- prod_cont_fin %>% 
      arrange(ORDEN_CIUDAD)
    
  }
  
  prod_ciud <- Ciudad(PRODUCCIONREALPOND)
  prod_ciud["DIFERENCIA"] <- prod_ciud[,ncol(prod_ciud)]-prod_ciud[,(ncol(prod_ciud)-1)]
  colnames(prod_ciud) <- paste0(names(prod_ciud),"_prod")
  
  vent_ciud <- Ciudad(VENTASREALESPOND)
  vent_ciud["DIFERENCIA"] <- vent_ciud[,ncol(vent_ciud)]-vent_ciud[,(ncol(vent_ciud)-1)]
  colnames(vent_ciud) <- paste0(names(vent_ciud),"_vent")
  
  emp_ciud <- Ciudad(PERSONAL)
  emp_ciud["DIFERENCIA"] <- emp_ciud[,ncol(emp_ciud)]-emp_ciud[,(ncol(emp_ciud)-1)]
  colnames(emp_ciud) <- paste0(names(emp_ciud),"_emp")
  
  sueld_ciud <- Ciudad(TOTALSUELDOSREALES)
  sueld_ciud["DIFERENCIA"] <- sueld_ciud[,ncol(sueld_ciud)]-sueld_ciud[,(ncol(sueld_ciud)-1)]
  colnames(sueld_ciud) <- paste0(names(sueld_ciud),"_sueld")
  
  horas_ciud <- Ciudad(TOTALHORAS)
  horas_ciud["DIFERENCIA"] <- horas_ciud[,ncol(horas_ciud)]-horas_ciud[,(ncol(horas_ciud)-1)]
  colnames(horas_ciud) <- paste0(names(horas_ciud),"_horas")
  
  Ciudes_ <- prod_ciud %>% 
    left_join(vent_ciud,by=c("ORDEN_CIUDAD_prod"="ORDEN_CIUDAD_vent",
                             "CIUDAD_prod"="CIUDAD_vent")) %>% 
    left_join(emp_ciud,by=c("ORDEN_CIUDAD_prod"="ORDEN_CIUDAD_emp",
                            "CIUDAD_prod"="CIUDAD_emp")) %>% 
    left_join(sueld_ciud,by=c("ORDEN_CIUDAD_prod"="ORDEN_CIUDAD_sueld",
                              "CIUDAD_prod"="CIUDAD_sueld")) %>% 
    left_join(horas_ciud,by=c("ORDEN_CIUDAD_prod"="ORDEN_CIUDAD_horas",
                              "CIUDAD_prod"="CIUDAD_horas"))
  
  Ciudes_ <- Ciudes_ %>% 
    rename(N_Dom=ORDEN_CIUDAD_prod,
           Dominio=CIUDAD_prod)
  
  
  addWorksheet(wb, sheetName = "Ciudades")
  
  for (i in 3:ncol(Ciudes_)){
    # Aplica el formato a la columna de números (columna A en este caso)
    addStyle(wb, sheet = "Ciudades", style = num_formato, 
             rows = 2:(nrow(Ciudes_) + 1), cols = i)
  }
  openxlsx::addStyle(wb, sheet ="Ciudades", row = 1, col = 1:ncol(Ciudes_), style = bold_style)
  
  
  # Escribir la base de datos en la primera hoja del archivo Excel
  writeData(wb, sheet = "Ciudades", x = Ciudes_)
  
  
  # Ciudades est ------------------------------------------------------------
  
  
  Ciud_est <- function(var_inte){
    agreg_inte <- data %>%
      select(ANIO,MES,ID_NUMORD,ORDEN_CIUDAD,CIUDAD,{{var_inte}})
    
    
    prod_cont <- agreg_inte %>%
      group_by(ANIO,MES,ID_NUMORD,ORDEN_CIUDAD,CIUDAD) %>%
      summarise(PRODREAL=sum({{var_inte}}))
    
    
    prod_cont <- prod_cont %>%
      pivot_wider(names_from = ANIO, values_from = PRODREAL)
    
    prod_cont <- replace(prod_cont,is.na(prod_cont),0)
    
    T_IND_0 <- data %>%
      select(ANIO,MES,ORDEN_CIUDAD,CIUDAD,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(ID_NUMORD=0,
                ORDEN_CIUDAD=0,
                CIUDAD="Total Industria",
                TOTAL=sum({{var_inte}}))
    
    
    T_IND_0$ANIO <- as.character(T_IND_0$ANIO)
    
    T_IND_0 <- T_IND_0 %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    prod_cont <- rbind(prod_cont,T_IND_0)
    
    n=ncol(prod_cont)
    for(i in 5:ncol(prod_cont)){
      if(i==n){
        prod_cont[paste0("var_",names(prod_cont)[i])] <- prod_cont[,i]/prod_cont[,i-1]*100-100
      }else{
        prod_cont[paste0("var_",names(prod_cont)[i+1])] <- prod_cont[,i+1]/prod_cont[,i]*100-100
      }
    }
    
    prod_cont_ori <- prod_cont %>%
      select(!matches("^var"))
    n=ncol(prod_cont_ori)
    
    T_IND <- data %>%
      select(ANIO,MES,ORDEN_CIUDAD,CIUDAD,{{var_inte}}) %>%
      group_by(ANIO,MES,ORDEN_CIUDAD,CIUDAD) %>%
      summarise(TOTAL=sum({{var_inte}}))
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    T_IND <- rbind(T_IND,T_IND_0 %>% select(!ID_NUMORD))
    
    
    colnames(T_IND) <- paste0(names(T_IND),"_T")
    
    prod_cont_ori <- prod_cont_ori %>%
      left_join(T_IND, by=c("MES"="MES_T",
                            "ORDEN_CIUDAD"="ORDEN_CIUDAD_T",
                            "CIUDAD"="CIUDAD_T"))
    
    
    for(i in 5:n){
      if(i==n){
        a <- prod_cont_ori[,i]-prod_cont_ori[,i-1]
        b <- prod_cont_ori[,(i+(anio-2018))]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i])]  <- (a/b)*100
      }else{
        a <- prod_cont_ori[,i+1]-prod_cont_ori[,i]
        b <- prod_cont_ori[,(i+(anio-2018)+1)]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i+1])]  <- (a/b)*100
      }
    }
    
    
    prod_cont_ori <- prod_cont_ori %>%
      select(MES,ID_NUMORD,ORDEN_CIUDAD,CIUDAD,matches("^cont"))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_longer(cols = 5:ncol(prod_cont_ori),
                   names_to = "CONT_ANIO", values_to = "CONTRIBUCION")
    
    prod_cont_ori <- prod_cont_ori %>%
      arrange(CONT_ANIO)
    
    prod_cont_ori["ANIO_MES"] <- paste0(prod_cont_ori$CONT_ANIO,"_",prod_cont_ori$MES)
    
    prod_cont_ori<- subset(prod_cont_ori, select = -c(CONT_ANIO,MES))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_wider(names_from = ANIO_MES, values_from = CONTRIBUCION)
    
    
    prod_cont_var <- prod_cont %>%
      select(MES,ID_NUMORD,ORDEN_CIUDAD,CIUDAD,matches("^var"))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_longer(cols = 5:ncol(prod_cont_var),
                   names_to = "VAR_ANIO", values_to = "VARIACION")
    
    prod_cont_var <- prod_cont_var %>%
      arrange(VAR_ANIO)
    
    prod_cont_var["ANIO_MES"] <- paste0(prod_cont_var$VAR_ANIO,"_",prod_cont_var$MES)
    
    prod_cont_var<- subset(prod_cont_var, select = -c(VAR_ANIO,MES))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_wider(names_from = ANIO_MES, values_from = VARIACION)
    
    prod_cont_fin <- prod_cont_var %>%
      left_join(prod_cont_ori, by=c("ID_NUMORD"="ID_NUMORD",
                                    "ORDEN_CIUDAD"="ORDEN_CIUDAD",
                                    "CIUDAD"="CIUDAD"))
    
    
    prod_cont_fin <- prod_cont_fin[,c("ID_NUMORD","ORDEN_CIUDAD","CIUDAD",
                                      paste0("var_",anio_inte,"_",mes_inte),
                                      paste0("var_",anio,"_",mes),
                                      paste0("cont_",anio_inte,"_",mes_inte),
                                      paste0("cont_",anio,"_",mes))]
    
  }
  
  
  prod_ciudest <- Ciud_est(PRODUCCIONREALPOND)
  prod_ciudest["DIFERENCIA"] <- prod_ciudest[,ncol(prod_ciudest)]-prod_ciudest[,(ncol(prod_ciudest)-1)]
  colnames(prod_ciudest) <- paste0(names(prod_ciudest),"_prod")
  
  vent_ciudest <- Ciud_est(VENTASREALESPOND)
  vent_ciudest["DIFERENCIA"] <- vent_ciudest[,ncol(vent_ciudest)]-vent_ciudest[,(ncol(vent_ciudest)-1)]
  colnames(vent_ciudest) <- paste0(names(vent_ciudest),"_vent")
  
  emp_ciudest <- Ciud_est(PERSONAL)
  emp_ciudest["DIFERENCIA"] <- emp_ciudest[,ncol(emp_ciudest)]-emp_ciudest[,(ncol(emp_ciudest)-1)]
  colnames(emp_ciudest) <- paste0(names(emp_ciudest),"_emp")
  
  sueld_ciudest <- Ciud_est(TOTALSUELDOSREALES)
  sueld_ciudest["DIFERENCIA"] <- sueld_ciudest[,ncol(sueld_ciudest)]-sueld_ciudest[,(ncol(sueld_ciudest)-1)]
  colnames(sueld_ciudest) <- paste0(names(sueld_ciudest),"_sueld")
  
  horas_ciudest <- Ciud_est(TOTALHORAS)
  horas_ciudest["DIFERENCIA"] <- horas_ciudest[,ncol(horas_ciudest)]-horas_ciudest[,(ncol(horas_ciudest)-1)]
  colnames(horas_ciudest) <- paste0(names(horas_ciudest),"_horas")
  
  Ciudes_estab <- prod_ciudest %>% 
    left_join(vent_ciudest,by=c("ID_NUMORD_prod"="ID_NUMORD_vent",
                                "ORDEN_CIUDAD_prod"="ORDEN_CIUDAD_vent",
                                "CIUDAD_prod"="CIUDAD_vent")) %>% 
    left_join(emp_ciudest,by=c("ID_NUMORD_prod"="ID_NUMORD_emp",
                               "ORDEN_CIUDAD_prod"="ORDEN_CIUDAD_emp",
                               "CIUDAD_prod"="CIUDAD_emp")) %>% 
    left_join(sueld_ciudest,by=c("ID_NUMORD_prod"="ID_NUMORD_sueld",
                                 "ORDEN_CIUDAD_prod"="ORDEN_CIUDAD_sueld",
                                 "CIUDAD_prod"="CIUDAD_sueld")) %>% 
    left_join(horas_ciudest,by=c("ID_NUMORD_prod"="ID_NUMORD_horas",
                                 "ORDEN_CIUDAD_prod"="ORDEN_CIUDAD_horas",
                                 "CIUDAD_prod"="CIUDAD_horas"))
  
  
  n_ciud_est <- colnames(Ciudes_estab)
  
  Ciudes_estab <- Ciudes_estab %>% 
    left_join(nomb_estab %>% select(!c(ANIO,MES)), by=c("ID_NUMORD_prod"="ID_NUMORD"))
  
  Ciudes_estab <- Ciudes_estab %>% 
    select(ID_NUMORD_prod,NOMBRE_ESTAB,NOVEDAD,CLASE_CIIU4,
           DOMINIO_39,c(n_ciud_est[2:length(n_ciud_est)])) %>%
    arrange(ORDEN_CIUDAD_prod,ID_NUMORD_prod) %>% 
    rename(Norden=ID_NUMORD_prod,
           Nombre=NOMBRE_ESTAB,
           Nov=NOVEDAD,
           EMMET_Clae=CLASE_CIIU4,
           Dom_39=DOMINIO_39,
           N_Dom=ORDEN_CIUDAD_prod,
           Ciudad=CIUDAD_prod)
  
  addWorksheet(wb, sheetName = "Ciudades est")
  
  
  for (i in 8:ncol(Ciudes_estab)){
    # Aplica el formato a la columna de números (columna A en este caso)
    addStyle(wb, sheet = "Ciudades est", style = num_formato, 
             rows = 2:(nrow(Ciudes_estab) + 1), cols = i)
  }
  openxlsx::addStyle(wb, sheet ="Ciudades est", row = 1, col = 1:ncol(Ciudes_estab), style = bold_style)
  
  
  # Escribir la base de datos en la primera hoja del archivo Excel
  writeData(wb, sheet = "Ciudades est", x = Ciudes_estab)
  
  
  # Depto est ---------------------------------------------------------------
  
  
  Depto_est <- function(var_inte){
    agreg_inte <- data %>%
      select(ANIO,MES,ID_NUMORD,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,
             {{var_inte}})
    
    
    prod_cont <- agreg_inte %>%
      group_by(ANIO,MES,ID_NUMORD,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO) %>%
      summarise(PRODREAL=sum({{var_inte}}))
    
    
    prod_cont <- prod_cont %>%
      pivot_wider(names_from = ANIO, values_from = PRODREAL)
    
    prod_cont <- replace(prod_cont,is.na(prod_cont),0)
    
    T_IND_0 <- data %>%
      select(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,{{var_inte}}) %>%
      group_by(ANIO,MES) %>%
      summarise(ID_NUMORD=0,
                ORDEN_DEPTO=0,
                INCLUSION_NOMBRE_DEPTO="Total Industria",
                TOTAL=sum({{var_inte}}))
    
    
    T_IND_0$ANIO <- as.character(T_IND_0$ANIO)
    
    T_IND_0 <- T_IND_0 %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    prod_cont <- rbind(prod_cont,T_IND_0)
    
    n=ncol(prod_cont)
    for(i in 5:ncol(prod_cont)){
      if(i==n){
        prod_cont[paste0("var_",names(prod_cont)[i])] <- prod_cont[,i]/prod_cont[,i-1]*100-100
      }else{
        prod_cont[paste0("var_",names(prod_cont)[i+1])] <- prod_cont[,i+1]/prod_cont[,i]*100-100
      }
    }
    
    prod_cont_ori <- prod_cont %>%
      select(!matches("^var"))
    n=ncol(prod_cont_ori)
    
    T_IND <- data %>%
      select(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,{{var_inte}}) %>%
      group_by(ANIO,MES,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO) %>%
      summarise(TOTAL=sum({{var_inte}}))
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>%
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    T_IND <- rbind(T_IND,T_IND_0 %>% select(!ID_NUMORD))
    
    colnames(T_IND) <- paste0(names(T_IND),"_T")
    
    prod_cont_ori <- prod_cont_ori %>%
      left_join(T_IND, by=c("MES"="MES_T",
                            "ORDEN_DEPTO"="ORDEN_DEPTO_T",
                            "INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO_T"))
    
    
    for(i in 5:n){
      if(i==n){
        a <- prod_cont_ori[,i]-prod_cont_ori[,i-1]
        b <- prod_cont_ori[,(i+(anio-2018))]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i])]  <- (a/b)*100
      }else{
        a <- prod_cont_ori[,i+1]-prod_cont_ori[,i]
        b <- prod_cont_ori[,(i+(anio-2018)+1)]
        prod_cont_ori[paste0("cont_",names(prod_cont_ori)[i+1])]  <- (a/b)*100
      }
    }
    
    
    prod_cont_ori <- prod_cont_ori %>%
      select(MES,ID_NUMORD,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,matches("^cont"))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_longer(cols = 5:ncol(prod_cont_ori),
                   names_to = "CONT_ANIO", values_to = "CONTRIBUCION")
    
    prod_cont_ori <- prod_cont_ori %>%
      arrange(CONT_ANIO)
    
    prod_cont_ori["ANIO_MES"] <- paste0(prod_cont_ori$CONT_ANIO,"_",prod_cont_ori$MES)
    
    prod_cont_ori<- subset(prod_cont_ori, select = -c(CONT_ANIO,MES))
    
    prod_cont_ori <- prod_cont_ori %>%
      pivot_wider(names_from = ANIO_MES, values_from = CONTRIBUCION)
    
    
    prod_cont_var <- prod_cont %>%
      select(MES,ID_NUMORD,ORDEN_DEPTO,INCLUSION_NOMBRE_DEPTO,matches("^var"))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_longer(cols = 5:ncol(prod_cont_var),
                   names_to = "VAR_ANIO", values_to = "VARIACION")
    
    prod_cont_var <- prod_cont_var %>%
      arrange(VAR_ANIO)
    
    prod_cont_var["ANIO_MES"] <- paste0(prod_cont_var$VAR_ANIO,"_",prod_cont_var$MES)
    
    prod_cont_var<- subset(prod_cont_var, select = -c(VAR_ANIO,MES))
    
    prod_cont_var <- prod_cont_var %>%
      pivot_wider(names_from = ANIO_MES, values_from = VARIACION)
    
    prod_cont_fin <- prod_cont_var %>%
      left_join(prod_cont_ori, by=c("ID_NUMORD"="ID_NUMORD",
                                    "ORDEN_DEPTO"="ORDEN_DEPTO",
                                    "INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
    
    
    prod_cont_fin <- prod_cont_fin[,c("ID_NUMORD","ORDEN_DEPTO",
                                      "INCLUSION_NOMBRE_DEPTO",
                                      paste0("var_",anio_inte,"_",mes_inte),
                                      paste0("var_",anio,"_",mes),
                                      paste0("cont_",anio_inte,"_",mes_inte),
                                      paste0("cont_",anio,"_",mes))]
    
    prod_cont_fin <- prod_cont_fin %>% 
      arrange(ORDEN_DEPTO,ID_NUMORD)
    
  }
  
  prod_deptoest <- Depto_est(PRODUCCIONREALPOND)
  prod_deptoest["DIFERENCIA"] <- prod_deptoest[,ncol(prod_deptoest)]-prod_deptoest[,(ncol(prod_deptoest)-1)]
  colnames(prod_deptoest) <- paste0(names(prod_deptoest),"_prod")
  
  vent_deptoest <- Depto_est(VENTASREALESPOND)
  vent_deptoest["DIFERENCIA"] <- vent_deptoest[,ncol(vent_deptoest)]-vent_deptoest[,(ncol(vent_deptoest)-1)]
  colnames(vent_deptoest) <- paste0(names(vent_deptoest),"_vent")
  
  
  emp_deptoest <- Depto_est(PERSONAL)
  emp_deptoest["DIFERENCIA"] <- emp_deptoest[,ncol(emp_deptoest)]-emp_deptoest[,(ncol(emp_deptoest)-1)]
  colnames(emp_deptoest) <- paste0(names(emp_deptoest),"_emp")
  
  sueld_deptoest <- Depto_est(TOTALSUELDOSREALES)
  sueld_deptoest["DIFERENCIA"] <- sueld_deptoest[,ncol(sueld_deptoest)]-sueld_deptoest[,(ncol(sueld_deptoest)-1)]
  colnames(sueld_deptoest) <- paste0(names(sueld_deptoest),"_sueld")
  
  horas_deptoest <- Depto_est(TOTALHORAS)
  horas_deptoest["DIFERENCIA"] <- horas_deptoest[,ncol(horas_deptoest)]-horas_deptoest[,(ncol(horas_deptoest)-1)]
  colnames(horas_deptoest) <- paste0(names(horas_deptoest),"_horas")
  
  
  Dep_est <- prod_deptoest %>% 
    left_join(vent_deptoest,by=c("ID_NUMORD_prod"="ID_NUMORD_vent",
                                 "ORDEN_DEPTO_prod"="ORDEN_DEPTO_vent",
                                 "INCLUSION_NOMBRE_DEPTO_prod"="INCLUSION_NOMBRE_DEPTO_vent")) %>% 
    left_join(emp_deptoest,by=c("ID_NUMORD_prod"="ID_NUMORD_emp",
                                "ORDEN_DEPTO_prod"="ORDEN_DEPTO_emp",
                                "INCLUSION_NOMBRE_DEPTO_prod"="INCLUSION_NOMBRE_DEPTO_emp")) %>% 
    left_join(sueld_deptoest,by=c("ID_NUMORD_prod"="ID_NUMORD_sueld",
                                  "ORDEN_DEPTO_prod"="ORDEN_DEPTO_sueld",
                                  "INCLUSION_NOMBRE_DEPTO_prod"="INCLUSION_NOMBRE_DEPTO_sueld")) %>% 
    left_join(horas_deptoest,by=c("ID_NUMORD_prod"="ID_NUMORD_horas",
                                  "ORDEN_DEPTO_prod"="ORDEN_DEPTO_horas",
                                  "INCLUSION_NOMBRE_DEPTO_prod"="INCLUSION_NOMBRE_DEPTO_horas"))
  
  
  n_dep_est <- colnames(Dep_est)
  
  nomb_estab <- data %>% 
    filter((ANIO==anio & MES==mes)) %>% 
    select(ID_NUMORD,NOMBRE_ESTAB,CLASE_CIIU4,NOVEDAD,
           ORDENDOMINDEPTO,AGREG_DOMINIO_REG) 
  
  
  Dep_est <- Dep_est %>% 
    left_join(nomb_estab %>% select(!c(ANIO,MES)), by=c("ID_NUMORD_prod"="ID_NUMORD"))
  
  Dep_est <- Dep_est %>% 
    select(ID_NUMORD_prod,NOMBRE_ESTAB,NOVEDAD,CLASE_CIIU4,
           DOMINIO_39,ORDENDOMINDEPTO,AGREG_DOMINIO_REG,
           c(n_dep_est[2:length(n_dep_est)])) %>% 
    arrange(ORDEN_DEPTO_prod,ORDENDOMINDEPTO,ID_NUMORD_prod) %>% 
    rename(Norden=ID_NUMORD_prod,
           Nombre=NOMBRE_ESTAB,
           Nov=NOVEDAD,
           EMMET_Clae=CLASE_CIIU4,
           Dom_39=DOMINIO_39,
           N_Dom=ORDENDOMINDEPTO,
           DominioDtp=AGREG_DOMINIO_REG,
           N_Dep=ORDEN_DEPTO_prod,
           Departamento=INCLUSION_NOMBRE_DEPTO_prod)
  
  addWorksheet(wb, sheetName = "PorestabalDpto")
  
  for (i in 10:ncol(Dep_est)){
    # Aplica el formato a la columna de números (columna A en este caso)
    addStyle(wb, sheet = "PorestabalDpto", style = num_formato, 
             rows = 2:(nrow(Dep_est) + 1), cols = i)
  }
  openxlsx::addStyle(wb, sheet ="PorestabalDpto", row = 1, col = 1:ncol(Dep_est), style = bold_style)
  
  
  
  # Escribir la base de datos en la primera hoja del archivo Excel
  writeData(wb, sheet = "PorestabalDpto", x = Dep_est)
  
  
  
  # Guardar archivo ---------------------------------------------------------
  
  openxlsx::saveWorkbook(wb, file = Salida, overwrite = TRUE)
  
  
  
}

