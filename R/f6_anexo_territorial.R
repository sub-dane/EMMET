#' Anexo territorial
#'
#' @description Esta funcion crea el archivo de Anexo Territorial.
#'  Tiene como insumo, la base de datos Tematica. El cuerpo de la función
#'  crea cada una de las hojas del reporte.
#'
#'  Los cuadros de salida muestran información complementaria a la registrada
#'  en el boletín de prensa con el fin de brindar la información a un nivel más
#'  desagregado tanto total nacional como a nivel de departamental, áreas
#'  metropolitanas y principales ciudades del país.
#'
#'  Los resultados se muestran con variaciones anuales, año corrido y doce meses,
#'  junto con sus respectivas contribuciones, según dominios, por las principales
#'  variables que se recolectan en el proceso: Producción (real y nominal),
#'  ventas (real y nominal) y empleo.
#'  De igual manera se presentan los resultados de sueldos causados y horas
#'  totales trabajadas para los dominios nacionales.
#'
#' @param mes Definir el mes a ejecutar, ej: 11
#' @param anio Definir el año a ejecutar, ej: 2022
#' @param directorio definir el directorio donde se encuentran ubicado los datos de entrada
#'
#' @return CSV file
#' @export
#'
#' @examples f6_aterritorial(directorio="Documents/DANE/Procesos DIMPE /PilotoEMMET",
#'                        mes=11,anio=2022)
#'
#' @details
#'  En los anexos territoriales se calculan las contribuciones y variaciones
#'  en tres diferentes periodos.
#'
#'  Es importante resaltar que el equipo en el que se ejecutará la función debe
#'  tener intalado Java. Ver \href{https://github.com/NataliArteaga/DANE.EMMET#readme}{README}
#'
#'  A continuación se muestran los  periodos
#'  y las formulas para realizar el calculo de estos:
#'
#'  Contribucion anual:
#'
#'  \deqn{
#'  CA_{trj} = \frac{(V_{trj} - V_{(t-12)rj)}}{\sum_{1}^n V_{(t-12)rj)}} *100
#'  }
#'
#'
#'
#'  Donde:
#'
#'  t: Mes de referencia de la publicación de la Operación Estadística
#'
#'  \eqn{V_{trj}}: Valor en el periodo t para el territorio r en el dominio j.
#'
#'  \eqn{V_{(t-12)rj}}: Valor en el periodo t-12 o en el año anterior, en el
#'  territorio r en el dominio j.
#'
#'  \eqn{\sum_{1}^n V_{(t-12)rj)}}: Sumatoria de los valores en el periodo t-12,
#'  en el territorio r y en el dominio j
#'
#'  Esta contribucion anual se interpreta como el aporte del domino j en el territorio
#'  r a la variación anual del mes de referencia en el domino j en el territorio r
#'
#'  Contribucion anio corrido:
#'
#'  \deqn{
#'  CAC_{Trj} = \frac{\sum_{i}^T(V_{trj} - \sum_{b}^{T-12} V_{trj)}}{\sum_{b}^{T-12} V_{trj)}} *100
#'  }
#'
#'
#'  Donde:
#'  t: Mes variando de enero a diciembre
#'
#'  T: Mes de referencia.
#'
#'  i: Siempre es el mes de enero.
#'
#'  b=i-12: Corresponde a enero del año anterior
#'
#'  \eqn{V_{trj}}: Valor de la variable en el periodo t en el territorio r en el dominio j
#'
#'  Esta contribucion de anio corrido se interpreta como el aporte del domino j
#'  en el territorio r a la variación año corrido del mes de referencia en el
#'  domino j en el territorio r.
#'
#'
#'  Contribucion anio acumulado:
#'
#'  \deqn{
#'  CAA_{Trj} = \frac{\sum_{a+1}^T(V_{trj} - \sum_{b+1}^{a} V_{trj)}}{\sum_{b+1}^{a} V_{trj)}} *100
#'  }
#'
#'  Donde:
#'
#'  t: Mes variando de enero a diciembre
#'
#'  T: Mes de referencia.
#'
#'  a=T-12
#'
#'  b=a-12: Corresponde al mes a del año anterior
#'
#'  \eqn{V_{trj}}: Valor de la variable en el periodo t en el territorio r en el
#'  dominio j
#'
#'  Nota: cuando las variables que denotan meses (a, b) son negativas representan
#'  el mes del año inmediatamente anterior.
#'
#'  Esta contribucion de anio acumulado se interpreta como el aporte del domino
#'  j en el territorio r a la variación acumulada anual del mes de referencia
#'  en el domino j en el territorio r.
#'
#'
#'  Variación anual:
#'
#'  Es la relación del índice o valor (para producción y ventas, categoría de
#'  contratación, sueldos, horas) en el mes de referencia (ti) con el índice o
#'  valor absoluto del mismo mes en el año anterior (t i-12), menos 1 por 100.
#'
#'  \deqn{
#'  VA = \frac{\text{índice o valor del mes de referencia}}
#'  {\text{índice o valor del mismo mes del año anterior}} -1 *100
#'  }
#'
#'  Se interpreta como el crecimiento o disminución porcentual, dependiendo de
#'  si el resultado es negativo o positivo, de la variable correspondiente en
#'  el mes de referencia, en relación al mismo mes del año anterior
#'
#'
#'  Variación Año Corrido:
#'
#'  \deqn{
#'  VAC = \frac{\sum \text{índice o valor de enero al mes de referencia del año actual}}
#'  {\sum \text{índice o valor de enero al mes de referencia del mismo mes del año anterior}} -1 *100
#'  }
#'
#'  Se interpreta como el crecimiento o disminución porcentual, dependiendo de
#'  si el resultado es negativo o positivo, de la variable correspondiente en lo
#'  corrido del año hasta el mes de referencia, en relación al mismo periodo del
#'  año anterior
#'
#'
#'  Variación Acumulado Anual:
#'
#'  \deqn{
#'  VAA = \frac{\sum \text{índice o valor desde} a_{+1} \text{enero al mes de referencia}}
#'  {\sum \text{índice o valor en el año anterior desde} a_{+1} \text{ hasta el mes de referencia}} -1 *100
#'  }
#'
#'  Donde:
#'
#'  t=mes de referencia
#'
#'  a=t-12
#'
#'  Se interpreta como el crecimiento o disminución porcentual, dependiendo de
#'  si el resultado es negativo o positivo, de la variable correspondiente en l
#'  os últimos 12 meses hasta el mes de referencia, en relación al mismo periodo
#'  del año anterior.
#'  Contribuciones porcentuales: aporte en puntos porcentuales de las variaciones
#'  individuales a la variación de un agregado.
#'
#'  La función escribe, en formato excel, las hojas:
#'
#'  1. Var y Cont Anual Dpto:
#'  Se caclcula la contribución y variación anual (%)
#'  del valor de la producción, ventas, y empleo, según departamento
#'
#'  2. Var y Cont Anual Desagreg Dp:
#'  Se caclcula la contribución y variación anual (%)
#'   del valor de la producción, ventas, y empleo, según clase industrial
#'   por departamento
#'
#'  3.Var y Cont Anual Áreas metrop:
#'  Se caclcula la contribución y variación anual (%)
#'  del valor de la producción, ventas, y empleo, según área metropolitana
#'
#'  4. Var y Cont Anual Ciudades:
#'  Se caclcula la contribución y variación anual (%)
#'  del valor de la producción, ventas, y empleo, según ciudad
#'
#'  5. Var y Cont Año corrido Dpto:
#'  Se caclcula la contribución y variación aaño corrido (%)
#'  del valor de la producción, ventas, y empleo, según departamento
#'
#'  6.Var y Cont Año corri Desag Dp:
#'  Se caclcula la contribución y variación año corrido (%)
#'  del valor de la producción, ventas, y empleo, según clase
#'  industrial por departamento
#'
#'  7.Var y Cont Año corrido Áreas met:
#'  Se caclcula la contribución y variación a año corrido (%)
#'  el valor de la producción, ventas, y empleo, según área metropolitana
#'
#'  8. Var y Cont Año corrid Ciudad:
#'  Se caclcula la contribución y variación año corrido (%)
#'  del valor de la producción, ventas, y empleo, según ciudad
#'
#'  9. Var y Cont doce meses Dpto:
#'  Se caclcula la contribución y variación doce meses (%)
#'  del valor de la producción, ventas, y empleo, según departamento
#'
#'  10.Var y Cont doce meses Desa:
#'  Se caclcula la contribución y variación  doce meses (%)
#'  del valor de la producción, ventas, y empleo, según clase
#'  industrial por departamento
#'
#'  11.Var y Cont docemeses Áreas:
#'  Se caclcula la contribución y variación doce meses (%)
#'  del valor de la producción, ventas, y empleo, según área metropolitana
#'
#'  12. Var y Cont docemeses Ciu:
#'  Se caclcula la contribución y variación ddoce meses (%)
#'  del valor de la producción, ventas, y empleo, según ciudad
#'
#'  13. Índices Departamentos:
#'  Se caclcula los índices, de producción nominal y real,
#'  ventas nominal y real, empleo según departamento y clase industrial
#'
#'  14. Índices Áreas Metropolitana:
#'  Se caclcula los índices de producción nominal y real, ventas nominal
#'  y real, empleo según área metropolitana
#'
#'  15. Índices Ciudades:
#'  Se caclcula los índices, de producción nominal y real, ventas nominal
#'  y real, empleo según ciudades
#'
#'  16. Var y Cont Trienal Dpto:
#'  Se caclcula la contribución y variación  trienal (%), es decir; usando
#'  como año base los datos del año 2019, ddel valor de la producción, ventas,
#'  y empleo, según departamento
#'
#'  17. Var y Cont Trienal Desa:
#'  Se caclcula la contribución y variación  trienal (%), es decir usando
#'  como año base los dtos del año 2019, del valor de la producción, ventas,
#'  y empleo, según clase industrial por departamento
#'
#'  18.Var y Cont Trienal Áreas:
#'  Se caclcula la contribución y variación  trienal (%), es decir usando
#'  como año base los dtos del año 2019, agrupando los datos por
#'  áreas metropolitanas de la produccion, ventas y personal
#'
#'  19. Var y Cont Trienal Ciuda:
#'  Se caclcula la contribución y y variación  trienal (%), es decir usando
#'  como año base los dtos del año 2019, del valor de la producción, ventas,
#'  y empleo, según ciudad.
#'
#'
#'  Finalmente se exporta una archivo excel, que contiene la informacion de las
#'  19 hojas.


f6_aterritorial <- function(directorio,
                            mes,
                            anio){


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
  library(xlsx)
  library(DT)
  library(data.table)

  source("https://raw.githubusercontent.com/NataliArteaga/DANE.EMMET/main/R/utils.R")

  # Cargar bases y variables ------------------------------------------------

  meses <- c("ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic")
  meses_enu <- c("Enero","Febrero","Marzo","Abril","Mayo","Junio",
                 "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
  data<-read.csv(paste0(directorio,"/results/S4_tematica/EMMET_PANEL_tematica_",meses[mes],anio,".csv"),fileEncoding = "latin1")

  deptos <- data %>%
    select(INCLUSION_NOMBRE_DEPTO,ORDEN_DEPTO)
  deptos <- unique(deptos)
  
  areas <- data %>%
    select(AREA_METROPOLITANA,ORDEN_AREA)
  areas <- unique(areas)
  
  ciudades <- data %>%
    select(CIUDAD,ORDEN_CIUDAD)
  ciudades <- unique(ciudades)
  
  # Archivos de entrada y salida --------------------------------------------
  
  formato <- paste0(directorio,"/data/anexos_territorial_emmet_formato.xlsx")
  Salida<-paste0(directorio,"/results/S5_anexos/anexos_territorial_emmet_",meses[mes],"_",anio,".xlsx")
  
  temp <- unlist(str_split(Salida, pattern = "/"))
  Salida2 <- paste(temp[-length(temp)], collapse="/")
  temp2 <- unlist(str_split(formato, pattern = "/"))
  temp2 <- temp2[length(temp2)]
  temp <- temp[length(temp)]
  file.copy(formato,Salida2)
  file.rename(file.path(Salida2,temp2),file.path(Salida2,temp))
  
  # Limpieza de nombres de variable -----------------------------------------
  
  colnames(data) <- colnames_format(data)
  data <-  data %>% mutate_at(vars(contains("OBSE")),~str_replace_all(.,pattern="[^[:alnum:]]",replacement=" "))
  
  
  # Se carga el formato de excel --------------------------------------------
  
  wb <- openxlsx::loadWorkbook(Salida)
  sheets <- getSheetNames(Salida)
  #names(sheets)
  
  filas_blanco<- function(df) {
    blank_row <- df[1, ]
    blank_row[,] <- NA
    
    df_with_blank <- df %>%
      group_by(ORDEN_DEPTO) %>%
      do(rbind(., blank_row)) 
    return(df_with_blank)
  }
  
  # Formatos ----------------------------------------------------------------
  
  
  colgr <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    fgFill = "#F2F2F2",
    bgFill = "#FFFFFF",
    halign = "center",
    valign = "center",
    numFmt = "0.00",
    wrapText = TRUE
  )
  
  
  colbl <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    fgFill = "#FFFFFF",
    halign = "center",
    valign = "center",
    numFmt = "0.00",
    wrapText = TRUE
  )
  
  ultbl <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    border = "Bottom: thin",
    borderColour = "#000000",
    fgFill = "#FFFFFF",
    bgFill = "#FFFFFF",
    halign = "center",
    valign = "center",
    wrapText = TRUE
  )
  
  
  ultcgr <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    border = "Right",
    borderColour = "#000000",
    halign = "center",
    valign = "center",
    fgFill = "#F2F2F2",
    bgFill = "#FFFFFF",
    numFmt = "0.00",
    wrapText = TRUE
  )
  
  
  ultcbl <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    border = "Right",
    borderColour = "#000000",
    halign = "center",
    valign = "center",
    fgFill = "#FFFFFF",
    bgFill = "#FFFFFF",
    numFmt = "0.00",
    wrapText = TRUE
  )
  
  rowbl <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    border = "Top",
    borderColour = "#000000",
    halign = "left",
    fgFill = "#FFFFFF",
    bgFill = "#FFFFFF",
    numFmt = "0.00",
    wrapText = TRUE
  )
  
  ultrbl <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    border = "Bottom: thin",
    borderColour = "#000000",
    fgFill = "#FFFFFF",
    bgFill = "#FFFFFF",
    numFmt = "0.00",
    wrapText = TRUE
  )
  
  
  
  # Funciones ---------------------------------------------------------------
  
  #Funcion para crear las variables produccion_total, ventas_total y personal_total
  cont_tot_summ <- function(datos,periodo){
    if(periodo==1){
      contribucion <- datos  %>%
        summarise(produccion_total=sum(PRODUCCIONREALPOND),
                  ventas_total=sum(VENTASREALESPOND),
                  personal_total=sum(TOTPERS))
    }
    if(periodo==2){
      contribucion <- datos  %>%
        summarise(produccionnom_total = sum(PRODUCCIONNOMPOND),
                  produccion_total = sum(PRODUCCIONREALPOND),
                  ventas_total=sum(VENTASREALESPOND),
                  personal_total=sum(TOTPERS))
      
    }
    if(periodo==3){
      contribucion <- datos  %>%
        summarise(produccionNom_total = mean(produccionNom_total),
                  produccion_total = mean(produccion_total),
                  ventasnom_total=mean(ventasnom_total),
                  ventas_total=mean(ventas_total),
                  personal_total=mean(personal_total))
    }
    if(periodo==4){
      contribucion <- datos  %>%
        summarise(produccionNom_total = sum(PRODUCCIONNOMPOND),
                  produccion_total = sum(PRODUCCIONREALPOND),
                  ventasnom_total=sum(VENTASNOMINPOND),
                  ventas_total=sum(VENTASREALESPOND),
                  personal_total=sum(TOTPERS))
    }
    return(contribucion)
  }
  
  #Funcion para crear las variavles pro,vent,per de acuerdo al periodo que se
  #esté manejando
  
  cont_summ <- function(datos,periodo){
    if(periodo==1){
      contribucion <- datos %>%
        summarise(prod = sum(PRODUCCIONREALPOND),
                  vent=sum(VENTASREALESPOND),
                  per=sum(PERSONAL))
    }
    if(periodo==2){
      contribucion <- datos %>%
        summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
                  ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
                  personal=(per[2]-per[1])/contribucion_total$personal_total)
    }
    if(periodo==3){
      contribucion <- datos %>%
        summarise(prod = sum(PRODUCCIONREALPOND),
                  vent=sum(VENTASREALESPOND),
                  per=sum(PERSONAL)) %>%
        group_by(INCLUSION_NOMBRE_DEPTO,ORDENDOMINDEPTO,AGREG_DOMINIO_REG) %>%
        summarise(produccion=(prod[2]-prod[1]),
                  ventas=(vent[2]-vent[1]),
                  personal=(per[2]-per[1]))
    }
    if(periodo==4){
      contribucion <- datos %>%
        summarise(produccionNom_mensual = sum(PRODUCCIONNOMPOND),
                  produccion_mensual = sum(PRODUCCIONREALPOND),
                  ventasnom_mensual=sum(VENTASNOMINPOND),
                  ventas_mensual=sum(VENTASREALESPOND),
                  personal_mensual=sum(TOTPERS))
    }
    if(periodo==5){
      contribucion <- datos %>%
        summarise(produccion=(produccion/produccion_total),
                  ventas=(ventas/ventas_total),
                  personas=(personal/personal_total))
      
    }
    return(contribucion)
  }
  
  #Funcion para realizar los pivotes en las tablas de acuerdo al periodo que
  #se esté trabajando
  tabla_piv_pas <- function(tabla, periodo){
    if(periodo==1){
      tabla <- tabla %>%
        pivot_wider(names_from = c("ANIO"),values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
      tabla[paste0("varprodnom_",anio)] <- (tabla[paste0("produccionNom_",anio)]-tabla[paste0("produccionNom_",anio-1)])/tabla[paste0("produccionNom_",anio-1)]
      tabla[paste0("varprod_",anio)] <- (tabla[paste0("produccion_",anio)]-tabla[paste0("produccion_",anio-1)])/tabla[paste0("produccion_",anio-1)]
      tabla[paste0("varventasnom_",anio)]<- (tabla[paste0("ventasNom_",anio)]-tabla[paste0("ventasNom_",anio-1)])/tabla[paste0("ventasNom_",anio-1)]
      tabla[paste0("varventas_",anio)]<- (tabla[paste0("ventas_",anio)]-tabla[paste0("ventas_",anio-1)])/tabla[paste0("ventas_",anio-1)]
      tabla[paste0("varpersonas_",anio)] <- (tabla[paste0("personas_",anio)]-tabla[paste0("personas_",anio-1)])/tabla[paste0("personas_",anio-1)]
      
    }
    if(periodo==2){
      tabla <- tabla %>%
        pivot_wider(names_from = c("ANIO2"),values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
      tabla[paste0("varprodnom_",anio)] <- (tabla[paste0("produccionNom_",anio)]-tabla[paste0("produccionNom_",anio-1)])/tabla[paste0("produccionNom_",anio-1)]
      tabla[paste0("varprod_",anio)] <- (tabla[paste0("produccion_",anio)]-tabla[paste0("produccion_",anio-1)])/tabla[paste0("produccion_",anio-1)]
      tabla[paste0("varventasnom_",anio)]<- (tabla[paste0("ventasNom_",anio)]-tabla[paste0("ventasNom_",anio-1)])/tabla[paste0("ventasNom_",anio-1)]
      tabla[paste0("varventas_",anio)]<- (tabla[paste0("ventas_",anio)]-tabla[paste0("ventas_",anio-1)])/tabla[paste0("ventas_",anio-1)]
      tabla[paste0("varpersonas_",anio)] <- (tabla[paste0("personas_",anio)]-tabla[paste0("personas_",anio-1)])/tabla[paste0("personas_",anio-1)]
      
      
    }
    if(periodo==3){
      tabla <- tabla %>%
        pivot_wider(names_from = c("ANIO"),values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
      tabla[paste0("varprodnom_",anio)] <- (tabla[paste0("produccionNom_",anio)]-tabla[paste0("produccionNom_",2019)])/tabla[paste0("produccionNom_",2019)]
      tabla[paste0("varprod_",anio)] <- (tabla[paste0("produccion_",anio)]-tabla[paste0("produccion_",2019)])/tabla[paste0("produccion_",2019)]
      tabla[paste0("varventasnom_",anio)]<- (tabla[paste0("ventasNom_",anio)]-tabla[paste0("ventasNom_",2019)])/tabla[paste0("ventasNom_",2019)]
      tabla[paste0("varventas_",anio)]<- (tabla[paste0("ventas_",anio)]-tabla[paste0("ventas_",2019)])/tabla[paste0("ventas_",2019)]
      tabla[paste0("varpersonas_",anio)] <- (tabla[paste0("personas_",anio)]-tabla[paste0("personas_",2019)])/tabla[paste0("personas_",2019)]
      
    }
    return(tabla)
  }
  
  tabla_sum_mut <- function(data,periodo){
    if(periodo==1){
      #1,2,3,4,5,6,7,8,9,10,11,12,16,17,18,19
      tabla1 <- data %>%
        summarise(produccionNom=sum(PRODUCCIONNOMPOND),
                  produccion=sum(PRODUCCIONREALPOND),
                  ventasNom = sum(VENTASNOMINPOND),
                  ventas = sum(VENTASREALESPOND),
                  personas=sum(TOTPERS))
    }
    if(periodo==2){
      #12,14,15
      tabla1<-data %>%
        mutate(produccionNom=round((produccionNom_mensual/produccionNom_total)*100,1),
               produccion=round((produccion_mensual/produccion_total)*100,1),
               ventasnom=round((ventasnom_mensual/ventasnom_total)*100,1),
               ventas=round((ventas_mensual/ventas_total)*100,1),
               personal=round((personal_mensual/personal_total)*100,1))
      
    }
    return(tabla1)
  }
  
  #Funcion de acople
  tabla_sapp <- function(tabla){
    tabla$ANIO <- sapply(strsplit(as.character(tabla$variables), "_"), `[`, 2)
    tabla$variables <- sapply(strsplit(as.character(tabla$variables), "_"), `[`, 1)
    tabla <- tabla %>% filter(gsub("var","",variables)!=variables )
    tabla <- tabla %>% pivot_wider(names_from = variables,values_from = value)
    
    return(tabla)
  }
  
  # 1. Var y Cont Anual Dpto ------------------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(MES==mes & ANIO%in%c(anio-1))
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual
  contribucion <- data %>%
    filter(MES==mes & ANIO%in%c(anio-1,anio)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO,MES,INCLUSION_NOMBRE_DEPTO)
  contribucion <- cont_summ(contribucion,1) %>%
    group_by(INCLUSION_NOMBRE_DEPTO)
  contribucion <- cont_summ(contribucion,2) %>%
    arrange(produccion)
  
  
  #Calculo de la variación por departamentos
  tabla1 <- data %>%
    filter(ANIO%in%c(anio-1,anio) & MES%in%mes) %>%
    group_by(ANIO,MES,INCLUSION_NOMBRE_DEPTO)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,1)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:2)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Empalme de la variacion y contribucion anual por departamentos
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("INCLUSION_NOMBRE_DEPTO"))
  tabla1 <- tabla1[,c("INCLUSION_NOMBRE_DEPTO","varprodnom","varprod","produccion",
                      "varventasnom","varventas","ventas","varpersonas","personal")]
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  tabla1 <- tabla1 %>% inner_join(deptos, by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  tabla1 <- tabla1 %>% arrange(ORDEN_DEPTO)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  #Total industria
  
  contribucion1 <- data %>%
    filter(MES==mes & ANIO%in%c(anio-1,anio)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO,MES)
  contribucion1 <- cont_summ(contribucion1,1)
  contribucion1["INCLUSION_NOMBRE_DEPTO"] <- "Total Industria"
  contribucion1 <- contribucion1 %>%
    group_by(INCLUSION_NOMBRE_DEPTO)
  contribucion1 <- cont_summ(contribucion1,2)
  
  
  #Calculo de la variación por departamentos
  tabla2 <- data %>%
    filter(ANIO%in%c(anio-1,anio) & MES%in%mes) %>%
    group_by(ANIO,MES)
  tabla2 <- tabla_sum_mut(tabla2,1)
  tabla2 <- tabla_piv_pas(tabla2,1)
  tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1:2)],names_to = "variables",values_to = "value" )
  
  tabla2 <- tabla_sapp(tabla2)
  
  
  tabla2["INCLUSION_NOMBRE_DEPTO"] <- "Total Industria"
  
  #Empalme de la variacion y contribucion anual por departamentos
  tabla2 <- inner_join(x=tabla2,y=contribucion1,by=c("INCLUSION_NOMBRE_DEPTO"))
  tabla2 <- tabla2[,c("INCLUSION_NOMBRE_DEPTO","varprodnom","varprod","produccion",
                      "varventasnom","varventas","ventas","varpersonas","personal")]
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personal")){
    tabla2[,i] <-  tabla2[,i]*100
  }
  
  tabla1_1 <- rbind(tabla2,tabla1)
  
  #Exportar
  
  sheet <- sheets[3]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1_1[,2:ncol(tabla1_1)]),
                      startRow = 12, startCol = 2,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0(meses_enu[mes],"(",anio,"/",anio-1,")p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 28, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  
  # 2. Var y Cont Anual Desagreg Dp -----------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(MES==mes & ANIO%in%c(anio-1)) %>%
    group_by(INCLUSION_NOMBRE_DEPTO)
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual
  contribucion <- data %>%
    filter(MES==mes & ANIO%in%c(anio-1,anio)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO,MES,INCLUSION_NOMBRE_DEPTO,ORDENDOMINDEPTO,AGREG_DOMINIO_REG)
  contribucion <- cont_summ(contribucion,3) %>%
    arrange(produccion)
  
  ##Calculo de la contribucion por sector
  contribucion_sector<-contribucion %>%
    left_join(contribucion_total,by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  
  
  contribucion<- cont_summ(contribucion_sector,5)
  
  #Calculo de la variación por departamentos
  tabla1 <- data %>%
    filter(ANIO%in%c(anio-1,anio) & MES%in%mes) %>%
    group_by(ANIO,MES,INCLUSION_NOMBRE_DEPTO,ORDENDOMINDEPTO,AGREG_DOMINIO_REG)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,1)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:5)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Empalme de la variacion y contribucion anual por departamentos
  tabla1 <- tabla1 %>%
    inner_join(y=contribucion,by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO",
                                   "ORDENDOMINDEPTO"="ORDENDOMINDEPTO"))
  tabla1 <- tabla1[,c("INCLUSION_NOMBRE_DEPTO","AGREG_DOMINIO_REG","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personas")]
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personas")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  tabla1_1["AGREG_DOMINIO_REG"]<- ""
  tabla1_1 <- tabla1_1 %>%
    rename(personas=personal)
  tabla1_1 <- tabla1_1 %>%
    select(names(tabla1))
  tabla1_1 <- tabla1_1[2:nrow(tabla1_1),]
  tabla1_1$produccion <- tabla1_1$varprod
  tabla1_1$ventas <- tabla1_1$varventas
  tabla1_1$personas <- tabla1_1$varpersonas
  
  conteo <- tabla1 %>%
    group_by(INCLUSION_NOMBRE_DEPTO) %>%
    summarise(Total=n()) %>%
    filter(Total==1)
  
  tabla1_1 <- tabla1_1 %>%
    filter( !INCLUSION_NOMBRE_DEPTO %in% conteo$INCLUSION_NOMBRE_DEPTO)
  
  tabla1 <- rbind(tabla1_1,tabla1)
  
  tabla1 <- tabla1 %>% inner_join(deptos, by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  
  tabla1_a <- tabla1 %>% 
    filter(!INCLUSION_NOMBRE_DEPTO %in% c("Cauca","Tolima","Boyacá","Córdoba","Otros Departamentos"))
  
  tabla1_b <- tabla1 %>% 
    filter(INCLUSION_NOMBRE_DEPTO %in% c("Cauca","Tolima","Boyacá","Córdoba","Otros Departamentos")) %>% 
    arrange(ORDEN_DEPTO)
  
  
  tabla1_a <- filas_blanco(tabla1_a)
  tabla1 <- rbind(tabla1_a,tabla1_b)
  
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  
  tabla2["AGREG_DOMINIO_REG"]<- ""
  tabla2 <- tabla2 %>%
    rename(personas=personal)
  tabla2 <- tabla2 %>%
    select(names(tabla1))
  
  tabla1 <- rbind(tabla2,tabla1)
  
  
  #Exportar
  
  sheet <- sheets[4]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,3:ncol(tabla1)]),
                      startRow = 12, startCol = 3,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0(meses_enu[mes],"(",anio,"/",anio-1,")p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 93, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 3.Var y Cont Anual Áreas metrop -----------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(MES==mes & ANIO%in%c(anio-1))
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual por areas mtp
  contribucion <- data %>%
    filter(MES==mes & ANIO%in%c(anio,anio-1)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO,MES,AREA_METROPOLITANA)
  contribucion <- cont_summ(contribucion,1) %>%
    group_by(AREA_METROPOLITANA)
  contribucion <- cont_summ(contribucion,2) %>%
    arrange(produccion)
  
  
  #Calculo de la variacion por areas mtp
  tabla1 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%mes) %>%
    group_by(ANIO,MES,AREA_METROPOLITANA)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,1)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:5)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Emplame de la contriucion y variacion por areas mtp
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("AREA_METROPOLITANA"))
  tabla1 <- tabla1[,c("AREA_METROPOLITANA","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  for( i in c("varprodnom",
              "varprod","produccion","varventasnom","varventas","ventas",
              "varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  
  tabla1 <- tabla1 %>% inner_join(areas, by=c("AREA_METROPOLITANA"="AREA_METROPOLITANA"))
  tabla1 <- tabla1 %>% arrange(ORDEN_AREA)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  
  tabla2 <- tabla2 %>%
    rename(personal=personas,
           AREA_METROPOLITANA=INCLUSION_NOMBRE_DEPTO)
  tabla2 <- tabla2 %>%
    select(names(tabla1))
  tabla1 <- rbind(tabla2,tabla1)
  
  #Exportar
  
  sheet <- sheets[5]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,2:ncol(tabla1)]),
                      startRow = 12, startCol = 2,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0(meses_enu[mes],"(",anio,"/",anio-1,")p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 18, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 4. Var y Cont Anual Ciudades --------------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(MES==mes & ANIO%in%c(anio-1))
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual por ciudad
  contribucion <- data %>%
    filter(MES==mes & ANIO%in%c(anio,anio-1)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO,MES,CIUDAD)
  contribucion <- cont_summ(contribucion,1) %>%
    group_by(CIUDAD)
  contribucion <- cont_summ(contribucion,2) %>%
    arrange(produccion)
  
  #Calculo de la variación por ciudad
  tabla1 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%mes) %>%
    group_by(ANIO,MES,CIUDAD)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,1)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:2)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Empalme de la variacion y contribucion anual por ciudad
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("CIUDAD"))
  tabla1 <- tabla1[,c("CIUDAD","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  tabla1 <- tabla1 %>% inner_join(ciudades, by=c("CIUDAD"="CIUDAD"))
  tabla1 <- tabla1 %>% arrange(ORDEN_CIUDAD)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  tabla2 <- tabla2 %>%
    rename(CIUDAD=AREA_METROPOLITANA)
  tabla2 <- tabla2 %>%
    select(names(tabla1))
  tabla1 <- rbind(tabla2,tabla1)
  
  
  #Exportar
  
  sheet <- sheets[6]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,2:ncol(tabla1)]),
                      startRow = 12, startCol = 2,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0(meses_enu[mes],"(",anio,"/",anio-1,")p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 24, startCol = 2,colNames=FALSE, rowNames=FALSE)
  
  # 5. Var y Cont Año corrido Dpto  -----------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio-1))
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual por dpto
  contribucion <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio,anio-1)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO,INCLUSION_NOMBRE_DEPTO)
  contribucion <- cont_summ(contribucion,1) %>%
    group_by(INCLUSION_NOMBRE_DEPTO)
  contribucion <- cont_summ(contribucion,2) %>%
    arrange(produccion)
  
  #Calculo de la variación por dpto
  tabla1 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%c(1:mes)) %>%
    group_by(ANIO,INCLUSION_NOMBRE_DEPTO)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  
  tabla1 <- tabla_piv_pas(tabla1,1)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:5)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Empalme de la variacion y contribucion anual por dpto
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("INCLUSION_NOMBRE_DEPTO"))
  tabla1 <- tabla1[,c("INCLUSION_NOMBRE_DEPTO","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  for( i in c("varprodnom",
              "varprod","produccion","varventasnom","varventas","ventas",
              "varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  
  tabla1 <- tabla1 %>% inner_join(deptos, by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  tabla1 <- tabla1 %>% arrange(ORDEN_DEPTO)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  
  #Total Industria
  
  contribucion1 <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio,anio-1)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO)
  contribucion1 <- cont_summ(contribucion1,1)
  contribucion1["INCLUSION_NOMBRE_DEPTO"] <- "Total Industria"
  contribucion1 <- contribucion1 %>%
    group_by(INCLUSION_NOMBRE_DEPTO)
  contribucion1 <- cont_summ(contribucion1,2)
  
  
  #Calculo de la variación por departamentos
  tabla2 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%c(1:mes))%>%
    group_by(ANIO)
  tabla2 <- tabla_sum_mut(tabla2,1)
  tabla2 <- tabla_piv_pas(tabla2,1)
  tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1:2)],names_to = "variables",values_to = "value" )
  
  tabla2 <- tabla_sapp(tabla2)
  
  
  tabla2["INCLUSION_NOMBRE_DEPTO"] <- "Total Industria"
  
  #Empalme de la variacion y contribucion anual por departamentos
  tabla2 <- inner_join(x=tabla2,y=contribucion1,by=c("INCLUSION_NOMBRE_DEPTO"))
  tabla2 <- tabla2[,c("INCLUSION_NOMBRE_DEPTO","varprodnom","varprod","produccion",
                      "varventasnom","varventas","ventas","varpersonas","personal")]
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personal")){
    tabla2[,i] <-  tabla2[,i]*100
  }
  
  tabla1_1 <- rbind(tabla2,tabla1)
  
  
  #Exportar
  
  sheet <- sheets[7]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1_1[,2:ncol(tabla1_1)]),
                      startRow = 12, startCol = 2,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Enero - ",meses_enu[mes],"(",anio,"/",anio-1,")p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 28, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 6.Var y Cont Año corri Desag Dp -----------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio-1)) %>%
    group_by(INCLUSION_NOMBRE_DEPTO)
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual por dpto
  contribucion <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio,anio-1)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO,INCLUSION_NOMBRE_DEPTO,ORDENDOMINDEPTO,AGREG_DOMINIO_REG)
  contribucion <- cont_summ(contribucion,3) %>%
    arrange(produccion)
  
  ##Calculo de la contribucion por sector
  contribucion_sector<-contribucion %>%
    left_join(contribucion_total,by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  
  
  contribucion<- cont_summ(contribucion_sector,5)
  
  #Calculo de la variación por dpto
  tabla1 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%c(1:mes)) %>%
    group_by(ANIO,INCLUSION_NOMBRE_DEPTO,ORDENDOMINDEPTO,AGREG_DOMINIO_REG)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,1)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:4)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Empalme de la variacion y contribucion anual por dpto
  tabla1 <- tabla1 %>%
    inner_join(y=contribucion,by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO",
                                   "ORDENDOMINDEPTO"="ORDENDOMINDEPTO"))
  
  tabla1 <- tabla1[,c("INCLUSION_NOMBRE_DEPTO","AGREG_DOMINIO_REG","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personas")]
  
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personas")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  
  tabla1_1["AGREG_DOMINIO_REG"]<- ""
  tabla1_1 <- tabla1_1 %>%
    rename(personas=personal)
  tabla1_1 <- tabla1_1 %>%
    select(names(tabla1))
  tabla1_1 <- tabla1_1[2:nrow(tabla1_1),]
  tabla1_1$produccion <- tabla1_1$varprod
  tabla1_1$ventas <- tabla1_1$varventas
  tabla1_1$personas <- tabla1_1$varpersonas
  
  conteo <- tabla1 %>%
    group_by(INCLUSION_NOMBRE_DEPTO) %>%
    summarise(Total=n()) %>%
    filter(Total==1)
  
  tabla1_1 <- tabla1_1 %>%
    filter( !INCLUSION_NOMBRE_DEPTO %in% conteo$INCLUSION_NOMBRE_DEPTO)
  
  tabla1 <- rbind(tabla1_1,tabla1)
  
  tabla1 <- tabla1 %>% inner_join(deptos, by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  tabla1_a <- tabla1 %>% 
    filter(!INCLUSION_NOMBRE_DEPTO %in% c("Cauca","Tolima","Boyacá","Córdoba","Otros Departamentos"))
  
  tabla1_b <- tabla1 %>% 
    filter(INCLUSION_NOMBRE_DEPTO %in% c("Cauca","Tolima","Boyacá","Córdoba","Otros Departamentos")) %>% 
    arrange(ORDEN_DEPTO)
  
  
  tabla1_a <- filas_blanco(tabla1_a)
  tabla1 <- rbind(tabla1_a,tabla1_b)
  
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  
  tabla2["AGREG_DOMINIO_REG"]<- ""
  tabla2 <- tabla2 %>%
    rename(personas=personal)
  tabla2 <- tabla2 %>%
    select(names(tabla1))
  
  tabla1 <- rbind(tabla2,tabla1)
  
  
  #Exportar
  
  sheet <- sheets[8]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,3:ncol(tabla1)]),
                      startRow = 12, startCol = 3,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Enero - ",meses_enu[mes],"(",anio,"/",anio-1,")p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 93, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 7.Var y Cont Anio corrido areas met  -----------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio-1))
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual por area mtp
  contribucion <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio,anio-1)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO,AREA_METROPOLITANA)
  contribucion <- cont_summ(contribucion,1) %>%
    group_by(AREA_METROPOLITANA)
  contribucion <- cont_summ(contribucion,2) %>%
    arrange(produccion)
  
  #Calculo de la variación por area mtp
  tabla1 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%c(1:mes)) %>%
    group_by(ANIO,AREA_METROPOLITANA)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,1)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:5)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Empalme de la variacion y contribucion anual por area mtp
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("AREA_METROPOLITANA"))
  tabla1 <- tabla1[,c("AREA_METROPOLITANA","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  for( i in c("varprodnom",
              "varprod","produccion","varventasnom","varventas","ventas",
              "varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  tabla1 <- tabla1 %>% inner_join(areas, by=c("AREA_METROPOLITANA"="AREA_METROPOLITANA"))
  tabla1 <- tabla1 %>% arrange(ORDEN_AREA)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  
  tabla2 <- tabla2 %>%
    rename(personal=personas,
           AREA_METROPOLITANA=INCLUSION_NOMBRE_DEPTO)
  tabla2 <- tabla2 %>%
    select(names(tabla1))
  tabla1 <- rbind(tabla2,tabla1)
  
  
  #Exportar
  
  sheet <- sheets[9]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,2:ncol(tabla1)]),
                      startRow = 12, startCol = 2,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Enero - ",meses_enu[mes],"(",anio,"/",anio-1,")p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 18, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 8. Var y Cont Año corrid Ciudad -----------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio-1))
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual por ciudad
  contribucion <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio,anio-1)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO,CIUDAD)
  contribucion <- cont_summ(contribucion,1) %>%
    group_by(CIUDAD)
  contribucion <- cont_summ(contribucion,2) %>%
    arrange(produccion)
  
  #Calculo de la variacion por ciudad
  tabla1 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%c(1:mes)) %>%
    group_by(ANIO,CIUDAD)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,1)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Empalme de la variacion y contribucion anual por ciudad
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("CIUDAD"))
  tabla1 <- tabla1[,c("CIUDAD","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  
  tabla1 <- tabla1 %>% inner_join(ciudades, by=c("CIUDAD"="CIUDAD"))
  tabla1 <- tabla1 %>% arrange(ORDEN_CIUDAD)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  
  tabla2 <- tabla2 %>%
    rename(CIUDAD=AREA_METROPOLITANA)
  tabla2 <- tabla2 %>%
    select(names(tabla1))
  
  tabla1 <- rbind(tabla2,tabla1)
  
  
  #Esportar
  
  sheet <- sheets[10]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,2:ncol(tabla1)]),
                      startRow = 12, startCol = 2,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Enero - ",meses[mes],"(",anio,"/",anio-1,")p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 24, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 9. Var y Cont doce meses Dpto -------------------------------------------
  
  #Creacion de la variable para anio corrido
  data$ANIO2 <- as.numeric(ifelse(data$MES%in%c((mes+1):12),data$ANIO+1,data$ANIO))
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(ANIO2%in%(anio-1))
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual por dtp
  contribucion <- data %>%
    filter(ANIO2%in%c(anio-1,anio)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO2,INCLUSION_NOMBRE_DEPTO)
  contribucion <- cont_summ(contribucion,1) %>%
    group_by(INCLUSION_NOMBRE_DEPTO)
  contribucion <- cont_summ(contribucion,2) %>%
    arrange(produccion)
  
  #Calculo de la variacion mensual por dtp
  tabla1 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2,INCLUSION_NOMBRE_DEPTO)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,2)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  
  #Empalme de la contribucion y la variacion
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("INCLUSION_NOMBRE_DEPTO"))
  tabla1 <- tabla1[,c("INCLUSION_NOMBRE_DEPTO","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  for( i in c("varprodnom",
              "varprod","produccion","varventasnom","varventas","ventas",
              "varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  tabla1 <- tabla1 %>% inner_join(deptos, by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  tabla1 <- tabla1 %>% arrange(ORDEN_DEPTO)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  #Total Industria
  
  contribucion1 <- data %>%
    filter(ANIO2%in%c(anio-1,anio)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO2)
  contribucion1 <- cont_summ(contribucion1,1)
  contribucion1["INCLUSION_NOMBRE_DEPTO"] <- "Total Industria"
  contribucion1 <- contribucion1 %>%
    group_by(INCLUSION_NOMBRE_DEPTO)
  contribucion1 <- cont_summ(contribucion1,2)
  
  
  #Calculo de la variación por departamentos
  tabla2 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2)
  tabla2 <- tabla_sum_mut(tabla2,1)
  
  tabla2 <- tabla_piv_pas(tabla2,2)
  tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1)],names_to = "variables",values_to = "value" )
  
  tabla2 <- tabla_sapp(tabla2)
  
  tabla2["INCLUSION_NOMBRE_DEPTO"] <- "Total Industria"
  
  #Empalme de la contribucion y la variacion
  tabla2 <- inner_join(x=tabla2,y=contribucion1,by=c("INCLUSION_NOMBRE_DEPTO"))
  tabla2 <- tabla2[,c("INCLUSION_NOMBRE_DEPTO","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  for( i in c("varprodnom",
              "varprod","produccion","varventasnom","varventas","ventas",
              "varpersonas","personal")){
    tabla2[,i] <-  tabla2[,i]*100
  }
  
  tabla1_1 <- rbind(tabla2,tabla1)
  
  #Exportar
  
  sheet <- sheets[11]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1_1[,2:ncol(tabla1_1)]),
                      startRow = 12, startCol = 2,colNames=FALSE, rowNames=FALSE)
  
  #mes_sig=meses[mes+1]
  
  Enunciado<-paste0(meses_enu[mes+1]," ",anio-1," - ",meses_enu[mes]," ",anio," / ",meses_enu[mes+1]," ",anio-2,"-",meses_enu[mes]," ", anio-1,"p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 28, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 10.Var y Cont doce meses Desa  ------------------------------------------
  
  #Creacion de la variable para anio corrido
  data$ANIO2 <- as.numeric(ifelse(data$MES%in%c((mes+1):12),data$ANIO+1,data$ANIO))
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(ANIO2%in%(anio-1)) %>%
    group_by(INCLUSION_NOMBRE_DEPTO)
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual
  contribucion <- data %>%
    filter(ANIO2%in%c(anio-1,anio)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO2,INCLUSION_NOMBRE_DEPTO,ORDENDOMINDEPTO,AGREG_DOMINIO_REG)
  contribucion <- cont_summ(contribucion,3) %>%
    arrange(produccion)
  
  ##Calculo de la contribucion por sector
  contribucion_sector<-contribucion %>%
    left_join(contribucion_total,by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  
  contribucion<- cont_summ(contribucion_sector,5)
  
  #Calculo de la variacion
  tabla1 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2,INCLUSION_NOMBRE_DEPTO,ORDENDOMINDEPTO,AGREG_DOMINIO_REG)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,2)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:4)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Empalme de la contribucion y la variacion
  tabla1 <- tabla1 %>%
    inner_join(y=contribucion,by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO",
                                   "ORDENDOMINDEPTO"="ORDENDOMINDEPTO"))
  
  tabla1 <- tabla1[,c("INCLUSION_NOMBRE_DEPTO","AGREG_DOMINIO_REG","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personas")]
  
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personas")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  
  tabla1_1["AGREG_DOMINIO_REG"]<- ""
  tabla1_1 <- tabla1_1 %>%
    rename(personas=personal)
  tabla1_1 <- tabla1_1 %>%
    select(names(tabla1))
  tabla1_1 <- tabla1_1[2:nrow(tabla1_1),]
  tabla1_1$produccion <- tabla1_1$varprod
  tabla1_1$ventas <- tabla1_1$varventas
  tabla1_1$personas <- tabla1_1$varpersonas
  
  conteo <- tabla1 %>%
    group_by(INCLUSION_NOMBRE_DEPTO) %>%
    summarise(Total=n()) %>%
    filter(Total==1)
  
  tabla1_1 <- tabla1_1 %>%
    filter( !INCLUSION_NOMBRE_DEPTO %in% conteo$INCLUSION_NOMBRE_DEPTO)
  
  tabla1 <- rbind(tabla1_1,tabla1)
  
  tabla1 <- tabla1 %>% inner_join(deptos, by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  tabla1_a <- tabla1 %>% 
    filter(!INCLUSION_NOMBRE_DEPTO %in% c("Cauca","Tolima","Boyacá","Córdoba","Otros Departamentos"))
  
  tabla1_b <- tabla1 %>% 
    filter(INCLUSION_NOMBRE_DEPTO %in% c("Cauca","Tolima","Boyacá","Córdoba","Otros Departamentos")) %>% 
    arrange(ORDEN_DEPTO)
  
  
  tabla1_a <- filas_blanco(tabla1_a)
  tabla1 <- rbind(tabla1_a,tabla1_b)
  
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  
  tabla2["AGREG_DOMINIO_REG"]<- ""
  tabla2 <- tabla2 %>%
    rename(personas=personal)
  tabla2 <- tabla2 %>%
    select(names(tabla1))
  
  tabla1 <- rbind(tabla2,tabla1)
  
  #Exportar
  
  sheet <- sheets[12]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,3:ncol(tabla1)]),
                      startRow = 12, startCol = 3,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0(meses_enu[mes+1]," ",anio-1," - ",meses_enu[mes]," ",anio," / ",meses_enu[mes+1]," ",anio-2," - ",meses_enu[mes]," ", anio-1,"p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 93, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 11.Var y Cont docemeses Áreas -------------------------------------------
  
  #Creacion de la variable para anio corrido
  data$ANIO2 <- as.numeric(ifelse(data$MES%in%c((mes+1):12),data$ANIO+1,data$ANIO))
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(ANIO2%in%(anio-1))
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual por area mtp
  contribucion <- data %>%
    filter(ANIO2%in%c(anio-1,anio)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO2,AREA_METROPOLITANA)
  contribucion <- cont_summ(contribucion,1) %>%
    group_by(AREA_METROPOLITANA)
  contribucion <- cont_summ(contribucion,2) %>%
    arrange(produccion)
  
  
  #Calculo de la variacion por area mtp
  tabla1 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2,AREA_METROPOLITANA)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,2)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:5)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Empalme de la contribucion y la variacion
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("AREA_METROPOLITANA"))
  tabla1 <- tabla1[,c("AREA_METROPOLITANA","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  for( i in c("varprodnom",
              "varprod","produccion","varventasnom","varventas","ventas",
              "varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  tabla1 <- tabla1 %>% inner_join(areas, by=c("AREA_METROPOLITANA"="AREA_METROPOLITANA"))
  tabla1 <- tabla1 %>% arrange(ORDEN_AREA)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  
  tabla2 <- tabla2 %>%
    rename(personal=personas,
           AREA_METROPOLITANA=INCLUSION_NOMBRE_DEPTO)
  tabla2 <- tabla2 %>%
    select(names(tabla1))
  tabla1 <- rbind(tabla2,tabla1)
  
  
  #Exportar
  
  sheet <- sheets[13]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,2:ncol(tabla1)]),
                      startRow = 12, startCol = 2,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0(meses_enu[mes+1]," ",anio-1," - ",meses_enu[mes]," ",anio," / ",meses_enu[mes+1]," ",anio-2," - ",meses_enu[mes+1]," ", anio-1,"p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 18, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 12. Var y Cont docemeses Ciu --------------------------------------------
  
  #Creacion de la variable para anio corrido
  data$ANIO2 <- as.numeric(ifelse(data$MES%in%c((mes+1):12),data$ANIO+1,data$ANIO))
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(ANIO2%in%(anio-1))
  contribucion_total <- cont_tot_summ(contribucion_total,1)
  
  #Calculo de la contribucion mensual por ciudad
  contribucion <- data %>%
    filter(ANIO2%in%c(anio-1,anio)) %>%
    mutate(PERSONAL=TOTPERS) %>%
    group_by(ANIO2,CIUDAD)
  contribucion <- cont_summ(contribucion,1) %>%
    group_by(CIUDAD)
  contribucion <- cont_summ(contribucion,2) %>%
    arrange(produccion)
  
  #Calculo de la variacion por ciudad
  tabla1 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2,CIUDAD)
  tabla1 <- tabla_sum_mut(tabla1,1)
  
  tabla1 <- tabla_piv_pas(tabla1,2)
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1)],names_to = "variables",values_to = "value" )
  
  tabla1 <- tabla_sapp(tabla1)
  
  #Empalme de la contribucion y la variacion
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("CIUDAD"))
  tabla1 <- tabla1[,c("CIUDAD","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  tabla1 <- tabla1 %>% inner_join(ciudades, by=c("CIUDAD"="CIUDAD"))
  tabla1 <- tabla1 %>% arrange(ORDEN_CIUDAD)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  
  tabla2 <- tabla2 %>%
    rename(CIUDAD=AREA_METROPOLITANA)
  tabla2 <- tabla2 %>%
    select(names(tabla1))
  tabla1 <- rbind(tabla2,tabla1)
  
  
  #Exportar
  
  sheet <- sheets[14]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,2:ncol(tabla1)]),
                      startRow = 12, startCol = 2,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0(meses_enu[mes+1]," ",anio-1," - ",meses_enu[mes]," ",anio," / ",meses_enu[mes+1]," ",anio-2,"-",meses_enu[mes]," ", anio-1,"p")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 24, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 13. Índices Departamentos -----------------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    group_by(INCLUSION_NOMBRE_DEPTO,ANIO,MES,AGREG_DOMINIO_REG)
  #contribucion_total <- cont_tot_summ(contribucion_total,4)
  
  contribucion_total <- contribucion_total  %>%
    summarise(produccionNom_total = sum(PRODUCCIONNOMPOND),
              produccion_total = sum(PRODUCCIONREALPOND),
              ventasnom_total=sum(VENTASNOMINPOND),
              ventas_total=sum(VENTASREALESPOND),
              personal_total=sum(TOTPERS))
  
  
  #Calculo de la contribucion mensual por dpto
  contribucion_total <- contribucion_total %>%
    group_by(INCLUSION_NOMBRE_DEPTO,ANIO,AGREG_DOMINIO_REG)
  #contribucion_total <- cont_tot_summ(contribucion_total,3)
  
  contribucion_total <- contribucion_total  %>%
    filter(ANIO==2018) %>%
    summarise(produccionNom_total = mean(produccionNom_total),
              produccion_total = mean(produccion_total),
              ventasnom_total=mean(ventasnom_total),
              ventas_total=mean(ventas_total),
              personal_total=mean(personal_total))
  
  #Calculo de indices
  contribucion_mensual <- data %>%
    group_by(INCLUSION_NOMBRE_DEPTO,ANIO,MES,AGREG_DOMINIO_REG)
  #contribucion_mensual <- cont_summ(contribucion_mensual,4)
  contribucion_mensual <- contribucion_mensual %>%
    summarise(produccionNom_mensual = sum(PRODUCCIONNOMPOND),
              produccion_mensual = sum(PRODUCCIONREALPOND),
              ventasnom_mensual=sum(VENTASNOMINPOND),
              ventas_mensual=sum(VENTASREALESPOND),
              personal_mensual=sum(TOTPERS))
  
  contribucion<-contribucion_mensual %>%
    left_join(contribucion_total,by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO",
                                      "AGREG_DOMINIO_REG"="AGREG_DOMINIO_REG"))
  
  # tabla1 <- tabla_sum_mut(contribucion,2) %>%
  #   select(INCLUSION_NOMBRE_DEPTO,ANIO.x,MES,AGREG_DOMINIO_REG.x,produccionNom,
  #          produccion,ventasnom,ventas,personal)
  
  
  tabla1<-contribucion %>%
    mutate(produccionNom=(produccionNom_mensual/produccionNom_total)*100,
           produccion   =(produccion_mensual/produccion_total)*100,
           ventasnom    =(ventasnom_mensual/ventasnom_total)*100,
           ventas       =(ventas_mensual/ventas_total)*100,
           personal     =(personal_mensual/personal_total)*100)%>%
    select(INCLUSION_NOMBRE_DEPTO,ANIO.x,MES,AGREG_DOMINIO_REG,produccionNom,
           produccion,ventasnom,ventas,personal) %>%
    rename("ANIO"="ANIO.x")
  
  
  #Calculo de la contribucion mensual
  contribucion_mensual <- data %>%
    group_by(ANIO,MES,INCLUSION_NOMBRE_DEPTO) %>%
    summarise(produccionNom_mensual=sum(PRODUCCIONNOMPOND),
              produccion_mensual   =sum(PRODUCCIONREALPOND),
              ventasNom_mensual    = sum(VENTASNOMINPOND),
              ventas_mensual       = sum(VENTASREALESPOND),
              #personas_mensual=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas_mensual     =sum(TOTPERS))
  
  #Calculo de la contribucion por el anio base
  contribucion_base <- contribucion_mensual %>%
    filter(ANIO==2018) %>%
    group_by(INCLUSION_NOMBRE_DEPTO) %>%
    summarise(produccionNom_total=mean(produccionNom_mensual),
              produccion_total=mean(produccion_mensual),
              ventasNom_total= mean(ventasNom_mensual),
              ventas_total= mean(ventas_mensual),
              personas_total=mean(personas_mensual))
  
  
  contribucion_mensual<-contribucion_mensual %>%
    mutate(produccionNom_total=contribucion_base$produccionNom_total,
           produccion_total=contribucion_base$produccion_total,
           ventasNom_total=contribucion_base$ventasNom_total,
           ventas_total=contribucion_base$ventas_total,
           personas_total=contribucion_base$personas_total)
  
  
  #Calculo de la variables nominales y totales
  tabla1_1<-contribucion_mensual %>%
    mutate(produccionNom =(produccionNom_mensual/produccionNom_total)*100,
           produccion    =(produccion_mensual/produccion_total)*100,
           ventasnom     =(ventasNom_mensual/ventasNom_total)*100,
           ventas        =(ventas_mensual/ventas_total)*100,
           personal      =(personas_mensual/personas_total)*100) %>%
    mutate(AGREG_DOMINIO_REG= "Total") %>%
    select(INCLUSION_NOMBRE_DEPTO,ANIO,MES,AGREG_DOMINIO_REG,produccionNom,
           produccion,ventasnom,ventas,personal)
  
  tabla1<- tabla1 %>% filter(!grepl("Total", AGREG_DOMINIO_REG, ignore.case = TRUE))
  
  tabla1<-rbind(tabla1,tabla1_1)
  
  tabla1 <- tabla1 %>% inner_join(deptos, by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  orden_agreg <- unique(data %>% select(AGREG_DOMINIO_REG,ORDENDOMINDEPTO))
  orden_agreg <- orden_agreg %>%
    mutate(ORDENDOMINDEPTO = ifelse(AGREG_DOMINIO_REG == "Total", 0, ORDENDOMINDEPTO))
  tabla1 <- tabla1 %>% 
    right_join(orden_agreg,
               by=c("AGREG_DOMINIO_REG"="AGREG_DOMINIO_REG"))
  tabla1 <- tabla1 %>% arrange(ORDEN_DEPTO,ORDENDOMINDEPTO) %>% 
    select(!c(ORDEN_DEPTO,ORDENDOMINDEPTO))
  
  
  
  #Exportar
  
  sheet <- sheets[15]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1),
                      startRow = 12, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Enero 2018 - ",meses_enu[mes]," ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+14), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fuente. DANE - EMMET")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+15), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("p: provisionales")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+16), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Nota: La diferencia entre el total y la suma de los dominios se debe a aproximaciones decimales.")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+17), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Otros departamentos: Amazonas, Arauca, Caquetá, Casanare, Cesar, Chocó, Huila, La Guajira, Magdalena, Meta, Nariño, Norte de Santander, Putumayo, Quindío, San Andrés, Sucre")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+18), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Alimentos y bebidas incluye las divisiones CIIU4 10 y 11")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+19), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Textiles y confecciones  incluye las divisiones CIIU4 13 y 14")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+20), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Curtido de cuero y calzado incluye la división CIIU4 15")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+21), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Madera y muebles  incluye las divisiones CIIU4 16 y 31")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+22), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Papel e imprentas  incluye las divisiones CIIU4 17 y 18")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+23), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Sustancias y productos químicos, farmacéuticos, de caucho y plástico  incluye las divisiones CIIU4 20 a 22")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+24), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Minerales no metálicos  division CIIU4 23")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+25), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Productos metálicos  incluye las divisiones CIIU4 24 y 25")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+26), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Vehículos de transporte, carrocerías, autopartes y otro equipo de transporte  incluye las divisiones CIIU4 29 y 30")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+27), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  filas_pares <- seq(12, nrow(tabla1)+12, by = 2)
  filas_impares <- seq(13, nrow(tabla1)+12, by = 2)
  addStyle(wb, sheet, style=colgr, rows = filas_pares, cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl, rows = filas_impares, cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultbl, rows = (nrow(tabla1)+12), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcgr, rows = filas_pares, cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = filas_impares, cols = ncol(tabla1), gridExpand = TRUE)
  
  addStyle(wb, sheet, style=rowbl, rows = (nrow(tabla1)+14), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl, rows = (nrow(tabla1)+14):(nrow(tabla1)+27), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = (nrow(tabla1)+14):(nrow(tabla1)+27), cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultrbl, rows = (nrow(tabla1)+27), cols = 1:ncol(tabla1), gridExpand = TRUE)
  
  # 14. Índices Áreas Metropolitana -----------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    group_by(AREA_METROPOLITANA,ANIO,MES)
  #contribucion_total <- cont_tot_summ(contribucion_total,4)
  
  contribucion_total <- contribucion_total  %>%
    summarise(produccionNom_total = sum(PRODUCCIONNOMPOND),
              produccion_total = sum(PRODUCCIONREALPOND),
              ventasnom_total=sum(VENTASNOMINPOND),
              ventas_total=sum(VENTASREALESPOND),
              personal_total=sum(TOTPERS))
  
  #Calculo de la contribucion mensual por area mtp
  contribucion_total <- contribucion_total %>%
    group_by(AREA_METROPOLITANA,ANIO)
  #contribucion_total <- cont_tot_summ(contribucion_total,3)
  
  contribucion_total <- contribucion_total  %>%
    filter(ANIO==2018) %>%
    summarise(produccionNom_total = mean(produccionNom_total),
              produccion_total = mean(produccion_total),
              ventasnom_total=mean(ventasnom_total),
              ventas_total=mean(ventas_total),
              personal_total=mean(personal_total))
  
  #Calculo de los indices
  contribucion_mensual <- data %>%
    group_by(AREA_METROPOLITANA,ANIO,MES)
  #contribucion_mensual <- cont_summ(contribucion_mensual,4)
  
  contribucion_mensual <- contribucion_mensual %>%
    summarise(produccionNom_mensual = sum(PRODUCCIONNOMPOND),
              produccion_mensual = sum(PRODUCCIONREALPOND),
              ventasnom_mensual=sum(VENTASNOMINPOND),
              ventas_mensual=sum(VENTASREALESPOND),
              personal_mensual=sum(TOTPERS))
  
  contribucion<-contribucion_mensual %>%
    left_join(contribucion_total,by=c("AREA_METROPOLITANA"="AREA_METROPOLITANA"))
  
  
  
  # tabla1 <- tabla_sum_mut(contribucion,2) %>%
  #   select(AREA_METROPOLITANA,ANIO,MES,produccionNom,
  #          produccion,ventasnom,ventas,personal)
  
  
  tabla1<-contribucion %>%
    mutate(produccionNom=(produccionNom_mensual/produccionNom_total)*100,
           produccion   =(produccion_mensual/produccion_total)*100,
           ventasnom    =(ventasnom_mensual/ventasnom_total)*100,
           ventas       =(ventas_mensual/ventas_total)*100,
           personal     =(personal_mensual/personal_total)*100) %>%
    select(AREA_METROPOLITANA,ANIO.x,MES,produccionNom,
           produccion,ventasnom,ventas,personal)
  
  tabla1 <- tabla1 %>% inner_join(areas, by=c("AREA_METROPOLITANA"="AREA_METROPOLITANA"))
  tabla1 <- tabla1 %>% arrange(ORDEN_AREA)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  
  #Exportar
  
  sheet <- sheets[16]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1),
                      startRow = 12, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Enero 2018 - ",meses_enu[mes]," ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+14), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fuente. DANE - EMMET")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+15), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("p: provisionales")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+16), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Nota: La diferencia entre el total y la suma de los dominios se debe a aproximaciones decimales")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+17), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  
  filas_pares <- seq(12, nrow(tabla1)+12, by = 2)
  filas_impares <- seq(13, nrow(tabla1)+12, by = 2)
  addStyle(wb, sheet, style=colgr, rows = filas_pares, cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl, rows = filas_impares, cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultbl, rows = (nrow(tabla1)+12), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcgr, rows = filas_pares, cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = filas_impares, cols = ncol(tabla1), gridExpand = TRUE)
  
  addStyle(wb, sheet, style=colbl, rows = (nrow(tabla1)+14):(nrow(tabla1)+17), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=rowbl, rows = (nrow(tabla1)+14), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = (nrow(tabla1)+14):(nrow(tabla1)+17), cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultrbl, rows = (nrow(tabla1)+17), cols = 1:ncol(tabla1), gridExpand = TRUE)
  
  # 15. Índices Ciudades ----------------------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    group_by(CIUDAD,ANIO,MES)
  #contribucion_total <- cont_tot_summ(contribucion_total,4)
  
  contribucion_total <- contribucion_total  %>%
    summarise(produccionNom_total = sum(PRODUCCIONNOMPOND),
              produccion_total = sum(PRODUCCIONREALPOND),
              ventasnom_total=sum(VENTASNOMINPOND),
              ventas_total=sum(VENTASREALESPOND),
              personal_total=sum(TOTPERS))
  
  
  #Calculo de la contribucion mensual por ciudad
  contribucion_total <- contribucion_total %>%
    group_by(CIUDAD,ANIO)
  #contribucion_total <- cont_tot_summ(contribucion_total,3)
  
  contribucion_total <- contribucion_total  %>%
    filter(ANIO==2018) %>%
    summarise(produccionNom_total = mean(produccionNom_total),
              produccion_total = mean(produccion_total),
              ventasnom_total=mean(ventasnom_total),
              ventas_total=mean(ventas_total),
              personal_total=mean(personal_total))
  
  
  #Calculo de los indices
  contribucion_mensual <- data %>%
    group_by(CIUDAD,ANIO,MES)
  #contribucion_mensual <- cont_summ(contribucion_mensual,4)
  
  contribucion_mensual <- contribucion_mensual %>%
    summarise(produccionNom_mensual = sum(PRODUCCIONNOMPOND),
              produccion_mensual = sum(PRODUCCIONREALPOND),
              ventasnom_mensual=sum(VENTASNOMINPOND),
              ventas_mensual=sum(VENTASREALESPOND),
              personal_mensual=sum(TOTPERS))
  
  contribucion<-contribucion_mensual %>%
    left_join(contribucion_total,by=c("CIUDAD"="CIUDAD"))
  
  # tabla1 <- tabla_sum_mut(contribucion,2) %>%
  #   select(CIUDAD,ANIO,MES,produccionNom,
  #          produccion,ventasnom,ventas,personal)
  
  tabla1<-contribucion %>%
    mutate(produccionNom=(produccionNom_mensual/produccionNom_total)*100,
           produccion   =(produccion_mensual/produccion_total)*100,
           ventasnom    =(ventasnom_mensual/ventasnom_total)*100,
           ventas       =(ventas_mensual/ventas_total)*100,
           personal     =(personal_mensual/personal_total)*100) %>%
    select(CIUDAD,ANIO.x,MES,produccionNom,
           produccion,ventasnom,ventas,personal)
  
  
  tabla1 <- tabla1 %>% inner_join(ciudades, by=c("CIUDAD"="CIUDAD"))
  tabla1 <- tabla1 %>% arrange(ORDEN_CIUDAD)
  tabla1 <- tabla1[,-ncol(tabla1)]
  
  #Exportar
  
  sheet <- sheets[17]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1),
                      startRow = 12, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Enero 2018 - ",meses_enu[mes]," ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+14), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fuente. DANE - EMMET")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+15), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("p: provisionales")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+16), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Nota: La diferencia entre el total y la suma de los dominios se debe a aproximaciones decimales")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+17), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  
  filas_pares <- seq(12, nrow(tabla1)+12, by = 2)
  filas_impares <- seq(13, nrow(tabla1)+12, by = 2)
  addStyle(wb, sheet, style=colgr, rows = filas_pares, cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl, rows = filas_impares, cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultbl, rows = (nrow(tabla1)+12), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcgr, rows = filas_pares, cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = filas_impares, cols = ncol(tabla1), gridExpand = TRUE)
  
  addStyle(wb, sheet, style=rowbl, rows = (nrow(tabla1)+14), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl, rows = (nrow(tabla1)+14):(nrow(tabla1)+17), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = (nrow(tabla1)+14):(nrow(tabla1)+17), cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultrbl, rows = (nrow(tabla1)+17), cols = 1:ncol(tabla1), gridExpand = TRUE)
  
  
  
  # # 16. Var y Cont Trienal Dpto  --------------------------------------------
  # 
  # #Calculo de la contribucion total
  # contribucion_total <- data %>%
  #   filter(MES==mes & ANIO%in%c(2019))
  # contribucion_total <- cont_tot_summ(contribucion_total,1)
  # 
  # #Calculo de la contribucion mensual
  # contribucion <- data %>%
  #   filter(MES==mes & ANIO%in%c(anio,2019)) %>%
  #   mutate(PERSONAL=TOTPERS) %>%
  #   group_by(ANIO,MES,INCLUSION_NOMBRE_DEPTO)
  # contribucion <- cont_summ(contribucion,1) %>%
  #   group_by(INCLUSION_NOMBRE_DEPTO)
  # contribucion <- cont_summ(contribucion,2) %>%
  #   arrange(produccion)
  # 
  # 
  # #Calculo de la variación por dpto
  # tabla1 <- data %>%
  #   filter(ANIO%in%c(anio,2019) & MES%in%mes) %>%
  #   group_by(ANIO,MES,INCLUSION_NOMBRE_DEPTO)
  # tabla1 <- tabla_sum_mut(tabla1,1)
  # 
  # tabla1 <- tabla_piv_pas(tabla1,3)
  # tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:5)],names_to = "variables",values_to = "value" )
  # 
  # tabla1 <- tabla_sapp(tabla1)
  # 
  # #Empalme de la contribucion y la variacion
  # tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("INCLUSION_NOMBRE_DEPTO"))
  # tabla1 <- tabla1[,c("INCLUSION_NOMBRE_DEPTO","varprodnom",
  #                     "varprod","produccion","varventasnom","varventas","ventas",
  #                     "varpersonas","personal")]
  # 
  # for( i in c("varprodnom",
  #             "varprod","produccion","varventasnom","varventas","ventas",
  #             "varpersonas","personal")){
  #   tabla1[,i] <-  tabla1[,i]*100
  # }
  # 
  # tabla1 <- tabla1 %>% inner_join(deptos, by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  # tabla1 <- tabla1 %>% arrange(ORDEN_DEPTO)
  # tabla1 <- tabla1[,-ncol(tabla1)]
  # 
  # #Total Industria
  # 
  # #Calculo de la contribucion mensual
  # contribucion1 <- data %>%
  #   filter(MES==mes & ANIO%in%c(anio,2019)) %>%
  #   mutate(PERSONAL=TOTPERS) %>%
  #   group_by(ANIO,MES)
  # contribucion1 <- cont_summ(contribucion1,1)
  # contribucion1["INCLUSION_NOMBRE_DEPTO"] <- "Total Industria"
  # contribucion1 <- contribucion1 %>%
  #   group_by(INCLUSION_NOMBRE_DEPTO)
  # contribucion1 <- cont_summ(contribucion1,2)
  # 
  # 
  # #Calculo de la variación por dpto
  # tabla2 <- data %>%
  #   filter(ANIO%in%c(anio,2019) & MES%in%mes) %>%
  #   group_by(ANIO,MES)
  # tabla2 <- tabla_sum_mut(tabla2,1)
  # 
  # tabla2 <- tabla_piv_pas(tabla2,3)
  # tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1:5)],names_to = "variables",values_to = "value" )
  # 
  # tabla2 <- tabla_sapp(tabla2)
  # 
  # tabla2["INCLUSION_NOMBRE_DEPTO"] <- "Total Industria"
  # 
  # #Empalme de la contribucion y la variacion
  # tabla2 <- inner_join(x=tabla2,y=contribucion1,by=c("INCLUSION_NOMBRE_DEPTO"))
  # tabla2 <- tabla2[,c("INCLUSION_NOMBRE_DEPTO","varprodnom",
  #                     "varprod","produccion","varventasnom","varventas","ventas",
  #                     "varpersonas","personal")]
  # 
  # for( i in c("varprodnom",
  #             "varprod","produccion","varventasnom","varventas","ventas",
  #             "varpersonas","personal")){
  #   tabla2[,i] <-  tabla2[,i]*100
  # }
  # 
  # 
  # tabla1_1 <- rbind(tabla2,tabla1)
  # 
  # 
  # #Exportar
  # 
  # sheet <- sheets[[18]]
  # addDataFrame(data.frame(tabla1_1), sheet, col.names=FALSE, row.names=FALSE, startRow = 12, startColumn = 1)
  # 
  # 
  # Enunciado<-paste0(meses_enu[mes]," (",anio,"/","2019)p")
  # addDataFrame(data.frame(Enunciado), sheet, col.names=FALSE, row.names=FALSE, startRow = 8, startColumn = 1)
  # 
  # Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  # addDataFrame(data.frame(Enunciado), sheet, col.names=FALSE, row.names=FALSE,
  #              startRow = 28, startColumn = 1)
  # 
  # # 17. Var y Cont Trienal Desagre  -----------------------------------------
  # 
  # #Calculo de la contribucion total
  # contribucion_total <- data %>%
  #   filter(MES==mes & ANIO%in%c(2019)) %>%
  #   group_by(INCLUSION_NOMBRE_DEPTO)
  # contribucion_total <- cont_tot_summ(contribucion_total,2)
  # 
  # #Calculo de la contribucion mensual
  # contribucion <- data %>%
  #   filter(MES==mes & ANIO%in%c(anio,2019)) %>%
  #   mutate(PERSONAL=TOTPERS) %>%
  #   group_by(ANIO,INCLUSION_NOMBRE_DEPTO,ORDENDOMINDEPTO,AGREG_DOMINIO_REG)
  # contribucion <- cont_summ(contribucion,3) %>%
  #   arrange(produccion)
  # 
  # contribucion_sector<-contribucion %>%
  #   left_join(contribucion_total,by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  # 
  # contribucion<- cont_summ(contribucion_sector,5)
  # 
  # #Calculo de la variacion
  # tabla1 <- data %>%
  #   filter(ANIO%in%c(anio,2019) & MES%in%mes) %>%
  #   group_by(ANIO,INCLUSION_NOMBRE_DEPTO,ORDENDOMINDEPTO,AGREG_DOMINIO_REG)
  # tabla1 <- tabla_sum_mut(tabla1,1)
  # 
  # tabla1 <- tabla_piv_pas(tabla1,3)
  # tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:4)],names_to = "variables",values_to = "value" )
  # 
  # tabla1 <- tabla_sapp(tabla1)
  # 
  # #Empalme de la contribucion y la variacion
  # tabla1 <- tabla1 %>%
  #   inner_join(y=contribucion,by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO",
  #                                  "ORDENDOMINDEPTO"="ORDENDOMINDEPTO"))
  # 
  # tabla1 <- tabla1[,c("INCLUSION_NOMBRE_DEPTO","AGREG_DOMINIO_REG","varprodnom",
  #                     "varprod","produccion","varventasnom","varventas","ventas",
  #                     "varpersonas","personas")]
  # 
  # 
  # for( i in c("varprodnom","varprod","produccion","varventasnom",
  #             "varventas","ventas","varpersonas","personas")){
  #   tabla1[,i] <-  tabla1[,i]*100
  # }
  # 
  # tabla1_1["AGREG_DOMINIO_REG"]<- ""
  # tabla1_1 <- tabla1_1 %>%
  #   rename(personas=personal)
  # tabla1_1 <- tabla1_1 %>%
  #   select(names(tabla1))
  # tabla1_1 <- tabla1_1[2:nrow(tabla1_1),]
  # tabla1_1$produccion <- tabla1_1$varprod
  # tabla1_1$ventas <- tabla1_1$varventas
  # tabla1_1$personas <- tabla1_1$varpersonas
  # 
  # conteo <- tabla1 %>%
  #   group_by(INCLUSION_NOMBRE_DEPTO) %>%
  #   summarise(Total=n()) %>%
  #   filter(Total==1)
  # 
  # tabla1_1 <- tabla1_1 %>%
  #   filter( !INCLUSION_NOMBRE_DEPTO %in% conteo$INCLUSION_NOMBRE_DEPTO)
  # 
  # tabla1 <- rbind(tabla1_1,tabla1)
  # 
  # tabla1 <- tabla1 %>% inner_join(deptos, by=c("INCLUSION_NOMBRE_DEPTO"="INCLUSION_NOMBRE_DEPTO"))
  # tabla1 <- tabla1 %>% arrange(ORDEN_DEPTO)
  # tabla1 <- tabla1[,-ncol(tabla1)]
  # 
  # 
  # tabla2["AGREG_DOMINIO_REG"]<- ""
  # tabla2 <- tabla2 %>%
  #   rename(personas=personal)
  # tabla2 <- tabla2 %>%
  #   select(names(tabla1))
  # 
  # tabla1 <- rbind(tabla2,tabla1)
  # 
  # #Exportar
  # 
  # sheet <- sheets[[19]]
  # addDataFrame(data.frame(tabla1), sheet, col.names=FALSE, row.names=FALSE, startRow = 12, startColumn = 1)
  # 
  # 
  # Enunciado<-paste0(meses_enu[mes]," (",anio,"/","2019)p")
  # addDataFrame(data.frame(Enunciado), sheet, col.names=FALSE, row.names=FALSE, startRow = 8, startColumn = 1)
  # 
  # Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  # addDataFrame(data.frame(Enunciado), sheet, col.names=FALSE, row.names=FALSE,
  #              startRow = 93, startColumn = 1)
  # 
  # # 18.Var y Cont Trienal Áreas me ------------------------------------------
  # 
  # #Calculo de la contribucion total
  # contribucion_total <- data %>%
  #   filter(MES==mes & ANIO%in%c(2019))
  # contribucion_total <- cont_tot_summ(contribucion_total,1)
  # 
  # #Calculo de la contribucion mensual por area mtp
  # contribucion <- data %>%
  #   filter(MES==mes & ANIO%in%c(anio,2019)) %>%
  #   mutate(PERSONAL=TOTPERS) %>%
  #   group_by(ANIO,MES,AREA_METROPOLITANA)
  # contribucion <- cont_summ(contribucion,1) %>%
  #   group_by(AREA_METROPOLITANA)
  # contribucion <- cont_summ(contribucion,2) %>%
  #   arrange(produccion)
  # 
  # 
  # #Calculo de la variacion por area mtp
  # tabla1 <- data %>%
  #   filter(ANIO%in%c(anio,2019) & MES%in%mes) %>%
  #   group_by(ANIO,MES,AREA_METROPOLITANA)
  # tabla1 <- tabla_sum_mut(tabla1,1)
  # 
  # tabla1 <- tabla_piv_pas(tabla1,3)
  # tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:5)],names_to = "variables",values_to = "value" )
  # 
  # tabla1 <- tabla_sapp(tabla1)
  # 
  # #Empalme de la contribucion y la variacion
  # tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("AREA_METROPOLITANA"))
  # tabla1 <- tabla1[,c("AREA_METROPOLITANA","varprodnom",
  #                     "varprod","produccion","varventasnom","varventas","ventas",
  #                     "varpersonas","personal")]
  # 
  # for( i in c("varprodnom",
  #             "varprod","produccion","varventasnom","varventas","ventas",
  #             "varpersonas","personal")){
  #   tabla1[,i] <-  tabla1[,i]*100
  # }
  # 
  # 
  # tabla1 <- tabla1 %>% inner_join(areas, by=c("AREA_METROPOLITANA"="AREA_METROPOLITANA"))
  # tabla1 <- tabla1 %>% arrange(ORDEN_AREA)
  # tabla1 <- tabla1[,-ncol(tabla1)]
  # 
  # 
  # tabla2 <- tabla2 %>%
  #   rename(personal=personas,
  #          AREA_METROPOLITANA=INCLUSION_NOMBRE_DEPTO)
  # tabla2 <- tabla2 %>%
  #   select(names(tabla1))
  # tabla1 <- rbind(tabla2,tabla1)
  # 
  # 
  # 
  # #Exportar
  # 
  # sheet <- sheets[[20]]
  # addDataFrame(data.frame(tabla1), sheet, col.names=FALSE, row.names=FALSE, startRow = 12, startColumn = 1)
  # 
  # 
  # Enunciado<-paste0(meses_enu[mes]," (",anio,"/","2019)p")
  # addDataFrame(data.frame(Enunciado), sheet, col.names=FALSE, row.names=FALSE, startRow = 8, startColumn = 1)
  # 
  # Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  # addDataFrame(data.frame(Enunciado), sheet, col.names=FALSE, row.names=FALSE,
  #              startRow = 18, startColumn = 1)
  # 
  # # 19. Var y Cont Trienal Ciudad ------------------------------------------
  # 
  # #Calculo de la contribucion total
  # contribucion_total <- data %>%
  #   filter(MES==mes & ANIO%in%c(2019))
  # contribucion_total <- cont_tot_summ(contribucion_total,1)
  # 
  # #Calculo de la contribucion mensual por ciudad
  # contribucion <- data %>%
  #   filter(MES==mes & ANIO%in%c(anio,2019)) %>%
  #   mutate(PERSONAL=TOTPERS) %>%
  #   group_by(ANIO,MES,CIUDAD)
  # contribucion <- cont_summ(contribucion,1) %>%
  #   group_by(CIUDAD)
  # contribucion <- cont_summ(contribucion,2) %>%
  #   arrange(produccion)
  # 
  # 
  # #Calculo de la variacion por ciudad
  # tabla1 <- data %>%
  #   filter(ANIO%in%c(anio,2019) & MES%in%mes) %>%
  #   group_by(ANIO,MES,CIUDAD)
  # tabla1 <- tabla_sum_mut(tabla1,1)
  # 
  # tabla1 <- tabla_piv_pas(tabla1,3)
  # tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:2)],names_to = "variables",values_to = "value" )
  # 
  # tabla1 <- tabla_sapp(tabla1)
  # 
  # #Empalme de la contribucion y la variacion
  # tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("CIUDAD"))
  # tabla1 <- tabla1[,c("CIUDAD","varprodnom",
  #                     "varprod","produccion","varventasnom","varventas","ventas",
  #                     "varpersonas","personal")]
  # 
  # 
  # for( i in c("varprodnom","varprod","produccion","varventasnom",
  #             "varventas","ventas","varpersonas","personal")){
  #   tabla1[,i] <-  tabla1[,i]*100
  # }
  # 
  # tabla1 <- tabla1 %>% inner_join(ciudades, by=c("CIUDAD"="CIUDAD"))
  # tabla1 <- tabla1 %>% arrange(ORDEN_CIUDAD)
  # tabla1 <- tabla1[,-ncol(tabla1)]
  # 
  # tabla2 <- tabla2 %>%
  #   rename(CIUDAD=AREA_METROPOLITANA)
  # tabla2 <- tabla2 %>%
  #   select(names(tabla1))
  # tabla1 <- rbind(tabla2,tabla1)
  # 
  # 
  # #Exportar
  # 
  # sheet <- sheets[[21]]
  # addDataFrame(data.frame(tabla1), sheet, col.names=FALSE, row.names=FALSE, startRow = 12, startColumn = 1)
  # 
  # 
  # Enunciado<-paste0(meses_enu[mes]," (",anio,"/","2019)p")
  # addDataFrame(data.frame(Enunciado), sheet, col.names=FALSE, row.names=FALSE, startRow = 8, startColumn = 1)
  # 
  # Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  # addDataFrame(data.frame(Enunciado), sheet, col.names=FALSE, row.names=FALSE,
  #              startRow = 24, startColumn = 1)
  # 
  # Guardar archivo de salida -----------------------------------------------
  
  openxlsx::saveWorkbook(wb,Salida,overwrite = TRUE)

}
