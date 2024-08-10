#' Cuadros desagregaion dominios
#'
#' @description Esta funcion crea un archivo de excel que contiene las variaciones
#' y contribuciones anueales desagregadas por dominio 39. Por cada dominio se crea
#' una hoja en donde se consigna las variaciones y contribuciones de las variables
#' de produccion, ventas, emppleados, sueldos y horas.
#' 
#' Adicionalmente se crea una hoja en donde se presentan los valores de defalctores
#' y se calcula la variación anual y mensual por cada dominio.
#'
#'  
#' @param mes Definir el mes a ejecutar, ej: 11
#' @param anio Definir el año a ejecutar, ej: 2022
#' @param directorio definir el directorio donde se encuentran ubicado los datos de entrada
#'
#' @return CSV file
#' @export
#'
#' @examples f7_cdominios(directorio="Documents/DANE/Procesos DIMPE /PilotoEMMET",
#'                        mes=11,anio=2022)
#'
#' @details Esta funcion se hace con la finalidad de crear cuadros informativos 
#' con datos desagregada según el interés, para facilitar el análisis del comportamiento
#' de cada secotr del dominio.
#' 
#'  

Cuadros_Dominios<-function(directorio,anio,mes){
  
  # Librerias ---------------------------------------------------------------
  
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(openxlsx)
  library(xlsx)
  library(stringr)
  
  source("https://raw.githubusercontent.com/sub-dane/EMMET/main/R/utils.R")
  
  # Cargar bases y variables ------------------------------------------------

  meses <- c("ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic")
  
  data<-read.csv(paste0(directorio,"/results/S4_tematica/EMMET_PANEL_tematica_",meses[mes],anio,".csv"),fileEncoding = "latin1")
  colnames(data) <- colnames_format(data)
  
  num_cols <- sapply(data, is.numeric)
  data[num_cols] <- lapply(data[num_cols], function(x) { x[is.na(x)] <- 0; x })
  
  DEFLACTOR <- read_excel(paste0(directorio,"/data/",anio,"/",meses[mes],"/DEFLACTOR_",meses[mes],anio,".xlsx"))
  
  
  # Crear un nuevo libro de Excel -------------------------------------------
  
  wb <- openxlsx::createWorkbook()
  Salida<-paste0(directorio,"/results/S5_anexos/cuadros_nacionales_",meses[mes],"_",anio,".xlsx")
 
  estilo_pink <- createStyle(fontColour = "black", fgFill = "#FFCCFF")
  estilo_green <- createStyle(fontColour = "black", fgFill = "#99FF99")
  bold_style <- createStyle(textDecoration = "bold")
  num_formato <- createStyle(numFmt = "0.0")
  num_style <- createStyle(numFmt = "#,##0.00")
  
  
  # Limpieza de nombres de variable -----------------------------------------
  
  colnames(data) <- colnames_format(data)
  data <-  data %>% mutate_at(vars(contains("OBSE")),~str_replace_all(.,pattern="[^[:alnum:]]",replacement=" "))
  
  
  # Directorio --------------------------------------------------------------
  
  directorio_hoja <- data %>% 
    filter((ANIO==anio & MES==mes)) %>% 
    select(ID_NUMORD,NUMEMP,NOVEDAD,NOMBRE_ESTAB,EMMET_CLASE,DOMINIO_39,INCLUSION_NOMBRE_DEPTO,
           AREA_METROPOLITANA,CIUDAD)
  
  directorio_hoja <- directorio_hoja %>% 
    rename(Nordes=ID_NUMORD,
           Nov=NOVEDAD,
           Departamento=INCLUSION_NOMBRE_DEPTO)
  
  
  addWorksheet(wb, sheetName = "Directorio")
  # Escribir la base de datos en la primera hoja del archivo Excel
  writeData(wb, sheet = "Directorio", x = directorio_hoja)
  
  
  # Aplicar el estilo de negrita a la primera fila
  openxlsx::addStyle(wb, sheet ="Directorio", row = 1, col = 1:ncol(directorio_hoja), style = bold_style)
  
  
  # Deflator ----------------------------------------------------------------
  
  deflactor <- DEFLACTOR %>% 
    select(CIIU4,ANO,MES,IPP_PyC)
  
  deflactor$IPP_PyC <- deflactor$IPP_PyC*100 
  
  deflactor <- deflactor %>% 
    pivot_wider(names_from = CIIU4, values_from = IPP_PyC)
  
  deflactor_anterior <- deflactor[c(nrow(deflactor)-12,nrow(deflactor)-24),]
  deflactor_anterior <- deflactor_anterior[1,]/deflactor_anterior[2,]*100-100
  
  deflactor_anual <- deflactor[c(nrow(deflactor),nrow(deflactor)-12),]
  deflactor_anual <- deflactor_anual[1,]/deflactor_anual[2,]*100-100
  
  deflactor_mensual <- deflactor[c(nrow(deflactor),nrow(deflactor)-1),]
  deflactor_mensual <- deflactor_mensual[1,]/deflactor_mensual[2,]*100-100
  
  deflactor$ANO <- as.character(deflactor$ANO)
  deflactor_anterior$ANO <- as.character(deflactor_anterior$ANO)
  deflactor_anual$ANO <- as.character(deflactor_anual$ANO)
  deflactor_mensual$ANO <- as.character(deflactor_mensual$ANO) 
  
  deflactor_anterior[1,1:2] <- c("Anual",anio-1)
  deflactor_anual[1,1:2] <- c("Anual",anio)
  deflactor_mensual[1,1:2] <- c("Mensual",anio)
  
  blanco <- rep("",ncol(deflactor))
  
  deflactor <- rbind(deflactor,blanco,deflactor_anterior,deflactor_anual,deflactor_mensual)
  
  
  deflactor_exp <- DEFLACTOR %>% 
    select(CIIU4,ANO,MES,IPP_EXP)
  
  deflactor_exp$IPP_EXP <- deflactor_exp$IPP_EXP*100 
  
  deflactor_exp <- deflactor_exp %>% 
    pivot_wider(names_from = CIIU4, values_from = IPP_EXP)
  
  deflactor_exp_anterior <- deflactor_exp[c(nrow(deflactor_exp)-12,nrow(deflactor_exp)-24),]
  deflactor_exp_anterior <- deflactor_exp_anterior[1,]/deflactor_exp_anterior[2,]*100-100
  
  deflactor_exp_anual <- deflactor_exp[c(nrow(deflactor_exp),nrow(deflactor_exp)-12),]
  deflactor_exp_anual <- deflactor_exp_anual[1,]/deflactor_exp_anual[2,]*100-100
  
  deflactor_exp_mensual <- deflactor_exp[c(nrow(deflactor_exp),nrow(deflactor_exp)-1),]
  deflactor_exp_mensual <- deflactor_exp_mensual[1,]/deflactor_exp_mensual[2,]*100-100
  
  
  deflactor_exp$ANO <- as.character(deflactor_exp$ANO)
  deflactor_exp_anterior$ANO <- as.character(deflactor_exp_anterior$ANO)
  deflactor_exp_anual$ANO <- as.character(deflactor_exp_anual$ANO)
  deflactor_exp_mensual$ANO <- as.character(deflactor_exp_mensual$ANO) 
  
  deflactor_exp_anterior[1,1:2] <- c("Anual",anio-1)
  deflactor_exp_anual[1,1:2] <- c("Anual",anio)
  deflactor_exp_mensual[1,1:2] <- c("Mensual",anio)
  
  deflactor <- rbind(deflactor,blanco,deflactor_exp,blanco,deflactor_exp_anterior,deflactor_exp_anual,deflactor_exp_mensual)
  deflactor <- deflactor %>%
    mutate_if(is.character, ~ gsub("\\.", ",", as.character(.)))
  
  
  addWorksheet(wb, "Deflactor")
  # Escribir la base de datos en la primera hoja del archivo Excel
  writeData(wb, sheet = "Deflactor", x = deflactor)
  # Aplicar el estilo de negrita a la primera fila
  openxlsx::addStyle(wb, sheet ="Deflactor", row = 1, col = 1:ncol(deflactor), style = bold_style)
  
  
  
  # Cont_Dominios -----------------------------------------------------------
  
  
  if(mes==1){
    mes_inte <- 12
    anio_inte <- anio-1
  }else{
    mes_inte <- mes-1
    anio_inte <- anio
  }
  
  
  #Funcion 
  cont_Dominios <- function(varia_int){
    prod_cont <- data %>% 
      select(ANIO,MES,DOMINIO_39,DESCRIPCIONDOMINIOEMMET39,{{varia_int}})
    
    prod_cont <- prod_cont %>%
      group_by(ANIO,MES,DOMINIO_39,DESCRIPCIONDOMINIOEMMET39) %>% 
      summarise(PRODREAL=sum({{varia_int}}))
    
    prod_cont <- prod_cont %>% 
      pivot_wider(names_from = ANIO, values_from = PRODREAL)
    
    T_IND <- data %>% 
      select(ANIO,MES,{{varia_int}}) %>% 
      group_by(ANIO,MES) %>% 
      summarise(DOMINIO_39=000,
                DESCRIPCIONDOMINIOEMMET39="Total Industria",
                TOTAL=sum({{varia_int}}))
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
      select(MES,DOMINIO_39,DESCRIPCIONDOMINIOEMMET39,c(as.character(2018:anio)))
    n=ncol(prod_cont_ori)
    
    T_IND <- data %>% 
      select(ANIO,MES,{{varia_int}}) %>% 
      group_by(ANIO,MES) %>% 
      summarise(TOTAL=sum({{varia_int}}))
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
      select(MES,DOMINIO_39,DESCRIPCIONDOMINIOEMMET39,matches("^cont"))
    
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
      select(MES,DOMINIO_39,DESCRIPCIONDOMINIOEMMET39,!c(as.character(2018:anio)))
    
    prod_cont_var <- prod_cont_var %>% 
      pivot_longer(cols = 4:ncol(prod_cont_var),
                   names_to = "VAR_ANIO", values_to = "VARIACION")
    
    prod_cont_var <- prod_cont_var %>% 
      arrange(VAR_ANIO)
    
    prod_cont_var["ANIO_MES"] <- paste0(prod_cont_var$VAR_ANIO,"_",prod_cont_var$MES)
    
    prod_cont_var<- subset(prod_cont_var, select = -c(VAR_ANIO,MES))
    
    prod_cont_var <- prod_cont_var %>% 
      pivot_wider(names_from = ANIO_MES, values_from = VARIACION)
    
    prod_cont <- prod_cont_var %>% 
      left_join(prod_cont_ori, by=c("DOMINIO_39"="DOMINIO_39","DESCRIPCIONDOMINIOEMMET39"="DESCRIPCIONDOMINIOEMMET39"))
    
    prod_cont <- prod_cont_ori <- prod_cont[, colSums(is.na(prod_cont)) != nrow(prod_cont)]
  }
  
  #Produccion
  
  prod_cont <- cont_Dominios(PRODUCCIONREALPOND)
  
  prod_cont["DIFERENCIA_cont"] <- prod_cont[,(ncol(prod_cont))]-prod_cont[,ncol(prod_cont)-1]
  
  #Ventas
  
  vent_cont <- cont_Dominios(VENTASREALESPOND)
  
  vent_cont <- vent_cont[,c("DOMINIO_39","DESCRIPCIONDOMINIOEMMET39",
                            paste0("var_",anio_inte,"_",mes_inte),
                            paste0("var_",anio,"_",mes),
                            paste0("cont_",anio_inte,"_",mes_inte),
                            paste0("cont_",anio,"_",mes))]
  
  vent_cont["DIFERENCIA_cont"] <- vent_cont[,ncol(vent_cont)]-vent_cont[,ncol(vent_cont)-1]
  
  
  #Empleo
  
  data <- data %>% 
    group_by(ANIO,MES,DOMINIO_39) %>% 
    mutate(PERSONAL=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL))
  
  emp_cont <- cont_Dominios(PERSONAL)
  
  emp_cont <- emp_cont[,c("DOMINIO_39","DESCRIPCIONDOMINIOEMMET39",
                          paste0("var_",anio_inte,"_",mes_inte),
                          paste0("var_",anio,"_",mes),
                          paste0("cont_",anio_inte,"_",mes_inte),
                          paste0("cont_",anio,"_",mes))]
  
  emp_cont["DIFERENCIA_cont"] <- emp_cont[,ncol(emp_cont)]-emp_cont[,ncol(emp_cont)-1]
  
  colnames(prod_cont) <- paste0(names(prod_cont),"_prod")
  colnames(vent_cont) <- paste0(names(vent_cont),"_vent")
  colnames(emp_cont) <- paste0(names(emp_cont),"_emp")
  
  Cont_Dominios <- prod_cont %>% 
    left_join(vent_cont, by=c("DOMINIO_39_prod"="DOMINIO_39_vent",
                              "DESCRIPCIONDOMINIOEMMET39_prod"="DESCRIPCIONDOMINIOEMMET39_vent")) %>% 
    left_join(emp_cont,by=c("DOMINIO_39_prod"="DOMINIO_39_emp",
                            "DESCRIPCIONDOMINIOEMMET39_prod"="DESCRIPCIONDOMINIOEMMET39_emp"))
  
  Cont_Dominios <- Cont_Dominios %>% 
    rename(DOMINIO_39=DOMINIO_39_prod,
           DESCRIPCION=DESCRIPCIONDOMINIOEMMET39_prod)
  
  Cont_Dominios <- Cont_Dominios %>% 
    arrange(DOMINIO_39)
  
  addWorksheet(wb, "Cont_Dominio")
  
  for (i in 3:ncol(Cont_Dominios)){
    # Aplica el formato a la columna de números (columna A en este caso)
    addStyle(wb, sheet = "Cont_Dominio", style = num_formato, 
             rows = 2:(nrow(Cont_Dominios) + 1), cols = i)
  }
  
  
  # Escribir el vector en la segunda hoja del archivo Excel
  writeData(wb, sheet = "Cont_Dominio", x = Cont_Dominios)
  openxlsx::addStyle(wb, sheet ="Cont_Dominio", row = 1, col = 1:ncol(Cont_Dominios), style = bold_style)
  
  # Dominios ----------------------------------------------------------------
  
  
  if(mes==1){
    mes_inte <- 12
    anio_inte <- anio-1
  }else{
    mes_inte <- mes-1
    anio_inte <- anio
  }
  
  #Función
  
  Dominios <- function (dominio,var_inte){
    dominio_inte <- data %>% 
      select(NOMBRE_ESTAB,ANIO,MES, DOMINIO_39,ORDEN_DEPTO,CLASE_CIIU4,NOVEDAD,
             ID_NUMORD,{{var_inte}}) %>% 
      filter(DOMINIO_39==dominio) 
    
    prod_cont <- dominio_inte  %>%
      group_by(ANIO,MES,ID_NUMORD) %>% 
      summarise(PRODREAL=sum({{var_inte}}))
    
    prod_cont <- prod_cont %>% 
      pivot_wider(names_from = ANIO, values_from = PRODREAL)
    
    prod_cont <- replace(prod_cont,is.na(prod_cont),0)
    
    T_IND <- dominio_inte %>% 
      select(ANIO,MES,{{var_inte}}) %>% 
      group_by(ANIO,MES) %>% 
      summarise(ID_NUMORD=000,
                TOTAL=sum({{var_inte}}))
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>% 
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    T_IND <- replace(T_IND,is.na(T_IND),0)
    
    prod_cont <- rbind(prod_cont,T_IND)
    
    n=ncol(prod_cont)
    for(i in 3:ncol(prod_cont)){
      if(i==n){
        prod_cont[paste0("var_",names(prod_cont)[i])] <- prod_cont[,i]/prod_cont[,i-1]*100-100
      }else{
        prod_cont[paste0("var_",names(prod_cont)[i+1])] <- prod_cont[,i+1]/prod_cont[,i]*100-100  
      }
      
    }
    
    prod_cont_ori <- prod_cont %>% 
      select(MES,ID_NUMORD,!matches("^var"))
    n=ncol(prod_cont_ori)
    
    T_IND <- dominio_inte %>% 
      select(ANIO,MES,{{var_inte}}) %>% 
      group_by(ANIO,MES) %>% 
      summarise(TOTAL=sum({{var_inte}}))
    T_IND$ANIO <- as.character(T_IND$ANIO)
    
    T_IND <- T_IND %>% 
      pivot_wider(names_from = ANIO, values_from = TOTAL)
    
    T_IND <- replace(T_IND,is.na(T_IND),0)
    
    colnames(T_IND) <- paste0(names(T_IND),"_T")
    
    prod_cont_ori <- prod_cont_ori %>% 
      left_join(T_IND, by=c("MES"="MES_T"))
    
    
    
    for(i in 3:n){
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
      select(MES,ID_NUMORD,matches("^cont"))
    
    prod_cont_ori <- prod_cont_ori %>% 
      pivot_longer(cols = 3:ncol(prod_cont_ori),
                   names_to = "CONT_ANIO", values_to = "CONTRIBUCION")
    
    prod_cont_ori <- prod_cont_ori %>% 
      arrange(CONT_ANIO)
    
    prod_cont_ori["ANIO_MES"] <- paste0(prod_cont_ori$CONT_ANIO,"_",prod_cont_ori$MES)
    
    prod_cont_ori<- subset(prod_cont_ori, select = -c(CONT_ANIO,MES))
    
    prod_cont_ori <- prod_cont_ori %>% 
      pivot_wider(names_from = ANIO_MES, values_from = CONTRIBUCION)
    
    
    prod_cont_var <- prod_cont %>% 
      select(MES,ID_NUMORD,matches("^var"))
    
    prod_cont_var <- prod_cont_var %>% 
      pivot_longer(cols = 3:ncol(prod_cont_var),
                   names_to = "VAR_ANIO", values_to = "VARIACION")
    
    prod_cont_var <- prod_cont_var %>% 
      arrange(VAR_ANIO)
    
    prod_cont_var["ANIO_MES"] <- paste0(prod_cont_var$VAR_ANIO,"_",prod_cont_var$MES)
    
    prod_cont_var<- subset(prod_cont_var, select = -c(VAR_ANIO,MES))
    
    prod_cont_var <- prod_cont_var %>% 
      pivot_wider(names_from = ANIO_MES, values_from = VARIACION)
    
    
    prod_cont <- prod_cont %>% 
      select(!matches("^var"))
    
    prod_cont <- prod_cont %>% 
      pivot_longer(cols = 3:ncol(prod_cont),
                   names_to = "ANIO", values_to = "ORIGINALES")
    
    prod_cont <- prod_cont %>% 
      arrange(ANIO)
    
    prod_cont["ANIO_MES"] <- paste0(prod_cont$ANIO,"_",prod_cont$MES)
    
    prod_cont <- subset(prod_cont, select = -c(ANIO,MES))
    
    prod_cont <- prod_cont %>% 
      pivot_wider(names_from = ANIO_MES, values_from = ORIGINALES)
    
    
    prod_cont_fin <- prod_cont %>% 
      left_join(prod_cont_var, by=c("ID_NUMORD"="ID_NUMORD"))
    
    prod_cont_fin <- prod_cont_fin %>% 
      left_join(prod_cont_ori, by=c("ID_NUMORD"="ID_NUMORD"))
    
    
    prod_cont_fin <- prod_cont_fin[,c("ID_NUMORD",
                                      paste0(anio_inte-1,"_",mes_inte),
                                      paste0(anio_inte,"_",mes_inte),
                                      paste0(anio-1,"_",mes),
                                      paste0(anio,"_",mes),
                                      paste0("var_",anio_inte,"_",mes_inte),
                                      paste0("var_",anio,"_",mes),
                                      paste0("cont_",anio_inte,"_",mes_inte),
                                      paste0("cont_",anio,"_",mes))]
  }
  
  estilo <- function (nom_hoja){
    
    dominio <- prod_dom %>% 
      select(!c(matches("^var"),matches("^cont"),matches("^DIF"))) %>% 
      left_join(vent_dom %>% 
                  select(!c(matches("^var"),matches("^cont"),matches("^DIF"))),
                by=c("ID_NUMORD_prod"="ID_NUMORD_vent")) %>% 
      left_join(emp_dom %>%select(!c(matches("^var"),matches("^cont"),matches("^DIF"))),
                by=c("ID_NUMORD_prod"="ID_NUMORD_emp")) %>% 
      left_join(sueld_dom %>%select(!c(matches("^var"),matches("^cont"),matches("^DIF"))),
                by=c("ID_NUMORD_prod"="ID_NUMORD_sueld")) %>% 
      left_join(horas_dom %>%select(!c(matches("^var"),matches("^cont"),matches("^DIF"))),
                by=c("ID_NUMORD_prod"="ID_NUMORD_horas")) %>% 
      left_join(prod_dom %>%select(ID_NUMORD_prod,c(matches("^var"),matches("^cont"),matches("^DIF"))),
                by=c("ID_NUMORD_prod"="ID_NUMORD_prod")) %>% 
      left_join(vent_dom %>% select(ID_NUMORD_vent,c(matches("^var"),matches("^cont"),matches("^DIF"))),
                by=c("ID_NUMORD_prod"="ID_NUMORD_vent")) %>% 
      left_join(emp_dom %>% select(ID_NUMORD_emp,c(matches("^var"),matches("^cont"),matches("^DIF"))),
                by=c("ID_NUMORD_prod"="ID_NUMORD_emp")) %>% 
      left_join(sueld_dom %>% select(ID_NUMORD_sueld,c(matches("^var"),matches("^cont"),matches("^DIF"))),
                by=c("ID_NUMORD_prod"="ID_NUMORD_sueld")) %>% 
      left_join(horas_dom %>% select(ID_NUMORD_horas,c(matches("^var"),matches("^cont"),matches("^DIF"))),
                by=c("ID_NUMORD_prod"="ID_NUMORD_horas"))
    
    
    dominio_fin <- dominio %>% 
      left_join(nomb_estab, by=c("ID_NUMORD_prod"="ID_NUMORD")) %>% 
      select("NOMBRE_ESTAB","NOMBREDPTO","CLASE_CIIU4","NOVEDAD","ID_NUMORD_prod",colnames(dominio)[-1])
    
    dominio_fin <- dominio_fin %>% 
      arrange(ID_NUMORD_prod)
    
    dominio_fin <- dominio_fin %>% 
      rename(Nombre=NOMBRE_ESTAB,
             Departamento=NOMBREDPTO,
             CIIU=CLASE_CIIU4,
             Nov=NOVEDAD)
    
    nombres_columnas <- colnames(dominio_fin)
    
    # Usando grep para encontrar las columnas que empiezan con "var"
    columnas_var <- nombres_columnas[grep("^var", nombres_columnas)]
    
    
    
    
    addWorksheet(wb, nom_hoja)
    
    
    
    for (j in 6:ncol(dominio_fin)){
      # Aplica el formato a la columna de números (columna A en este caso)
      addStyle(wb, sheet = nom_hoja, style = num_formato, 
               rows = 2:(nrow(dominio_fin) + 1), cols = j)
    }
    
    for (j in columnas_var) {
      for (i in 1:nrow(dominio_fin)){
        if(is.na(dominio_fin[i,j])!=TRUE & dominio_fin[i,j] < -20){
          addStyle(wb, sheet = nom_hoja, rows = i+1, cols = which(names(dominio_fin) == j), style = estilo_pink)
        }else if(is.na(dominio_fin[i,j])!=TRUE & dominio_fin[i,j] > 20){
          addStyle(wb, sheet = nom_hoja, rows = i+1, cols = which(names(dominio_fin) == j), style = estilo_green)
        }
      }
      
    }
    
    
    
    
    #names(dominio)
    return(dominio_fin)
  }
  
  
  data <- data %>% 
    group_by(ANIO,MES,ID_NUMORD,DOMINIO_39) %>% 
    mutate(PERSONAL=sum(TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
  
  nomb_estab <- data %>% 
    filter((ANIO==anio & MES==mes)) %>% 
    select(ID_NUMORD,NOMBRE_ESTAB,NOMBREDPTO,CLASE_CIIU4,NOVEDAD) 
  
  
  
  # 1010 --------------------------------------------------------------------
  
  #sort(unique(data$DOMINIO_39),FALSE)
  for(i in sort(unique(data$DOMINIO_39),FALSE)){
    prod_dom <- Dominios(i,PRODUCCIONREALPOND)
    prod_dom["DIFERENCIA"] <- prod_dom[,ncol(prod_dom)]-prod_dom[,(ncol(prod_dom)-1)]
    colnames(prod_dom) <- paste0(names(prod_dom),"_prod")
    
    vent_dom <- Dominios(i,VENTASREALESPOND)
    vent_dom["DIFERENCIA"] <- vent_dom[,ncol(vent_dom)]-vent_dom[,(ncol(vent_dom)-1)]
    colnames(vent_dom) <- paste0(names(vent_dom),"_vent")
    
    emp_dom <- Dominios(i,II_TOT_TOT_PERS)
    emp_dom["DIFERENCIA"] <- emp_dom[,ncol(emp_dom)]-emp_dom[,(ncol(emp_dom)-1)]
    colnames(emp_dom) <- paste0(names(emp_dom),"_emp")
    
    sueld_dom <- Dominios(i,TOTALSUELDOSREALES)
    sueld_dom["DIFERENCIA"] <- sueld_dom[,ncol(sueld_dom)]-sueld_dom[,(ncol(sueld_dom)-1)]
    colnames(sueld_dom) <- paste0(names(sueld_dom),"_sueld")
    
    horas_dom <- Dominios(i,TOTALHORAS)
    horas_dom["DIFERENCIA"] <- horas_dom[,ncol(horas_dom)]-horas_dom[,(ncol(horas_dom)-1)]
    colnames(horas_dom) <- paste0(names(horas_dom),"_horas")
    
    
    
    dom_i <- estilo(as.character(i))
    
    
    # Aplicar el estilo de negrita a la primera fila
    openxlsx::addStyle(wb, sheet = as.character(i), row = 1, col = 1:ncol(dom_i), style = bold_style)
    
    
    #Escribir el vector en la segunda hoja del archivo Excel
    writeData(wb, sheet = as.character(i), x = dom_i)                                                
    
  }
  
  
  # Guardar archivo ---------------------------------------------------------
  
  openxlsx::saveWorkbook(wb, file = Salida, overwrite = TRUE)
  
  
}






