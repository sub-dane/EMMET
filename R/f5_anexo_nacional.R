#' Anexo Nacional
#'
#' @description Esta funcion crea el archivo de Anexo Nacional.
#'  Tiene como insumo, la base de datos Tematica. El cuerpo de la función
#'  crea cada una de las hojas del reporte.
#'
#'  Los cuadros de salida o anexos estadísticos de EMMET, muestran información
#'  complementaria a la registrada en el boletín de prensa con el fin de brindar
#'  la información a un nivel más desagregado tanto total nacional como
#'  desagregada a nivel de departamentos, áreas metropolitanas y principales
#'  ciudades del país.
#'
#'  Los resultados se muestran con variaciones anuales, año corrido y doce meses,
#'  junto con sus respectivas contribuciones, según dominios, por las principales
#'  variables que se recolectan en el proceso: Producción (real y nominal),
#'  ventas (real y nominal) y empleo, desagregando a su vez esta última variable
#'  por área funcional y tipo de contrato.
#'  De igual manera se presentan los resultados de sueldos causados y horas
#'  totales trabajadas para los dominios nacionales.
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
#' @examples f5_anacional(directorio="Documents/DANE/Procesos DIMPE /PilotoEMMET",
#'                        mes=11,anio=2022)
#'
#'
#' @details En los anexos nacionales se calculan las contribuciones y variaciones
#'  en tres diferentes periodos.
#'
#'  Es importante resaltar que el equipo en el que se ejecutará la función debe
#'  tener intalado Java. Ver \href{https://github.com/NataliArteaga/DANE.EMMET#readme}{README}
#'
#'  A continuación se muestran los  periodos
#'  y las formulas para realizar el cálculo de estos:
#'
#'  Contribución anual:
#'
#'   \deqn{
#'  CA_{trj} = \frac{(V_{trj} - V_{(t-12)rj)}}{\sum_{1}^n V_{(t-12)rj)}} *100
#'  }
#'
#'
#'  Donde:
#'
#'  t: Mes de referencia de la publicación de la operación estadística
#'
#'  \eqn{V_{trj}}: Valor en el periodo t para el territorio r en el dominio j.
#'
#'  \eqn{V_{(t-12)rj}}: Valor en el periodo t-12 o en el año anterior, en el
#'  territorio r en el dominio j.
#'
#'  \eqn{\sum_{1}^n V_{(t-12)rj)}}: Sumatoria de los valores en el periodo t-12,
#'  en el territorio r y en el dominio j
#'
#'  Esta contribución anual se interpreta como el aporte del domino j
#'  en el territorio r a la variación anual del mes de referencia en el
#'  domino j en el territorio r
#'
#'  Contribución año corrido:
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
#'  \eqn{V_{trj}}: Valor de la variable en el periodo t en el territorio r
#'  en el dominio j
#'
#'  Esta contribucion de año corrido se interpreta como el aporte del domino j
#'  en el territorio r a la variación año corrido del mes de referencia en el
#'  domino j en el territorio r.
#'
#'
#'  Contribución año acumulado:
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
#'  Esta contribución de año acumulado se interpreta como el aporte del domino
#'  j en el territorio r a la variación acumulada anual del mes de referencia
#'  en el domino j en el territorio r.
#'
#'
#'  Variación anual:
#'
#'  Es la relación del índice o valor (para producción y ventas, categoría de
#'  contratación, sueldos, horas) en el mes de referencia (ti) con el índice o
#'  valor absoluto del mismo mes en el año anterior (t^{i-12}), menos 1 por 100.
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
#'\deqn{
#'  VAA = \frac{\sum \text{índice o valor desde} a_{+1} \text{hasta el mes de referencia}}
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
#'  si el resultado es negativo o positivo, de la variable correspondiente en
#'  los últimos 12 meses hasta el mes de referencia, en relación al mismo periodo
#'  del año anterior.
#'
#'  Contribuciones porcentuales: aporte en puntos porcentuales de las variaciones
#'  individuales a la variación de un agregado.
#'
#'
#'  La función escribe, en formato excel, las hojas:
#'
#'  1. Var y cont_Anual:
#'  Se caclcula la Variación anual (%) y contribución, del valor
#'  de la producción, ventas, y empleo, según las clase industrial
#'
#'  2. Var Anual_Emp_Sueldos_Horas:
#'  Se caclculan las Variaciones anuales, según clase industrial,
#'  producción, ventas, personal ocupado, sueldos y horas totales
#'  trabajadas
#'
#'  3. Var y cont_año corrido:
#'  Se caclcula la Variación año corrido (%) y contribución, del valor de
#'  la producción, ventas, y empleo, según las clase industrial
#'
#'  4. Var año corr Emp_Sueldos_Hor:
#'  Se caclculan las variaciones año corrido, según clase industrial
#'  producción, ventas, personal ocupado, sueldos y horas totales
#'  trabajadas
#'
#'
#'  5. Var y cont_doce meses:
#'  Se caclcula la  aariación doce meses (%) y contribución, del valor
#'  de la producción, ventas, y empleo, según las clase industrial
#'
#'
#'  6. Var doce meses Emp_Sueldos_H:
#'  Se caclculan las variaciones doce meses, según clase industrial,
#'  producción, ventas, personal ocupado, sueldos y horas totales trabajadas .
#'
#'
#'  7. Indices total por clase:
#'  Se calculan los índices de producción nominal y real, ventas nominal y real,
#'  empleo, sueldos y  horas Totales trabajadas, según clase industrial para
#'  cada uno de los dominios que se encuentran en la encuesta mensual
#'  manufacturera.
#'
#'  8. Índices Desestacionalizados:
#'  Para la contrucción de la desestacionalización se usa la librería: seasonal.
#'  Esta libreria es una adaptacion del programa X13 seasonal, para el programa
#'  de R.
#'
#'  Para esta función se crea un calendario en donde se ingresan los festivos
#'  fijos y los de fecha variables, para asignar la ponderación a los días, en
#'  donde finalmente quedan expresadas los días de Lunes a Sábados.
#'
#'  Posteriormente para realizar la desestacionalización en las variables de
#'  producción, ventas y empleo, se escoge de manera individual la variable de
#'  interés con la columna de las fechas, este nuevo data frame
#'  (ej: producción, fecha) se convierte en una serie de tiempo, con frecuencia
#'  mensual.
#'
#'  Para la desestacionalización, se usa la serie de tiempo construida
#'  anteriormente, y se especifican:
#'
#'  * fecha de inicio y fin de la serie (ej: 2001.1,2021.11).
#'
#'  * variables regresoras: el calendario, construido anteriormente,
#'
#'  * variables dicotomicas de outliers: estas son festivos patrios,
#'       años bisiesto, y fechas de alto impacto, como:
#'
#'       2016.Jul: Paro camionero
#'
#'       2016.Aug:
#'
#'       2020.Mar: Inicio de pandemia Covid-19
#'
#'       2020.Apr: Inicio de la cuarentena
#'
#'       2021.May: Paro Nacional
#'
#'  Adicionalmente se especifican los parámetros estacionales y la informacion
#'  de la serie que se desea observar.
#'
#'  Finalmente se crea una tabla con la fecha y los datos desestacionalizados.
#'
#'  Esto, por cada una de las variables de interés.
#'
#'
#'  9. Enlace legal hasta 2014:
#'  Se realiza una conexión entre la serie de la metodología actual, con las
#'  variaciones de los indices de la metedología anterior con el fin de tener una
#'  serie más larga de comparación de la información.
#'
#'
#'  10. Enlace legal hasta 2001:
#'  Se realiza una conexión entre la serie de la metodología actual, con las
#'  variaciones de los indices de la metedología anterior con el fin de tener una
#'  serie más larga de comparación de la información.
#'
#'  11. Var y cont_Trienal:
#'  Se caclcula la contribución y variación trienal, es decir usando como año
#'  base los datos del año 2019, del valor de la producción, ventas, y empleo,
#'  según clase industrial.
#'
#'
f5_anacional <- function(directorio,
                         mes,
                         anio){
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
  library(data.table)
  library(purrr)
  source("https://raw.githubusercontent.com/sub-dane/EMMET/main/R/utils.R")


  # Cargar bases y variables ------------------------------------------------


  meses <- tolower(meses)
  meses_enu <- c("Enero","Febrero","Marzo","Abril","Mayo","Junio",
                 "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
  meses_min<-tolower(meses_enu)
  data <- read.csv(paste0(directorio,"/results/S4_tematica/EMMET_PANEL_tematica_",meses[mes],anio,".csv"),fileEncoding = "latin1")
  indices_14<- read_xlsx(paste0(directorio,"/data/Indices_2014.xlsx"))
  indices_01<- read_xlsx(paste0(directorio,"/data/Indices_2001.xlsx"))
  
  
  # Archivos de entrada y salida --------------------------------------------
  
  formato<-paste0(directorio,"/data/anexos_nacional_emmet_formato.xlsx")
  Salida<-paste0(directorio,"/results/S5_anexos/anexos_nacional_emmet_",meses[mes],"_",anio,".xlsx")
  
  
  temp <- unlist(str_split(Salida, pattern = "/"))
  Salida2 <- paste(temp[-length(temp)], collapse="/")
  temp2 <- unlist(str_split(formato, pattern = "/"))
  temp2 <- temp2[length(temp2)]
  temp <- temp[length(temp)]
  file.copy(formato,Salida2)
  file.rename(file.path(Salida2,temp2),file.path(Salida2,temp))
  
  
  # Limpieza de nombres de variable -----------------------------------------
  
  colnames_format <- function(base){
    colnames(base) <- toupper(colnames(base))
    colnames(base) <- gsub(" ","",colnames(base))
    colnames(base) <- gsub("__","_",colnames(base))
    colnames(base) <- stringi::stri_trans_general(colnames(base), id = "Latin-ASCII")
    return(colnames(base))
  }
  colnames(data) <- colnames_format(data)
  #colnames(data2) <- colnames_format(data2)
  data <-  data %>% mutate_at(vars(contains("OBSE")),~str_replace_all(.,pattern="[^[:alnum:]]",replacement=" "))
  
  colnames(indices_14)<-colnames_format(indices_14)
  colnames(indices_01)<-colnames_format(indices_01)
  
  
  
  # Se carga el formato de excel --------------------------------------------
  
  wb <- openxlsx::loadWorkbook(Salida)
  sheets <- getSheetNames(Salida)
  #names(wb)
  
  #Estilos
  #num_formato <- createStyle(numFmt = "0.0")
  #boldStyle <- createStyle(textDecoration = "bold")
  
  
  # Formatos ----------------------------------------------------------------
  
  
  colgr <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    fgFill = "#F2F2F2",
    bgFill = "#FFFFFF",
    halign = "center",
    valign = "center",
    numFmt = "0.0"
  )
  
  
  colbl <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    fgFill = "#FFFFFF",
    halign = "center",
    valign = "center",
    numFmt = "0.0"
  )

  colgr_in <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    fgFill = "#F2F2F2",
    bgFill = "#FFFFFF"
  )
  
  
  colbl_in <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    fgFill = "#FFFFFF"
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
    valign = "center"
  )
  
  ultblfc <- createStyle(
      fontName = "Segoe UI",
      fontSize = 9,
      fontColour = "#000000",
      border = "Bottom: thin, Right: thin",
      borderColour = "#000000",
      fgFill = "#FFFFFF",
      bgFill = "#FFFFFF",
      halign = "center",
      valign = "center"
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
    numFmt = "0.0"
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
    numFmt = "0.0"
  )
  
  rowbl <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    border = "Top",
    borderColour = "#000000",
    halign = "left",
    fgFill = "#FFFFFF",
    bgFill = "#FFFFFF"
  )

  rowblf <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    border ="Top: thin, Right: thin",
    borderColour = "#000000",
    halign = "left",
    fgFill = "#FFFFFF",
    bgFill = "#FFFFFF"
  )  
  
  ultrbl <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    border = "Bottom: thin",
    borderColour = "#000000",
    fgFill = "#FFFFFF",
    bgFill = "#FFFFFF",
  )
  ultrblf <- createStyle(
    fontName = "Segoe UI",
    fontSize = 9,
    fontColour = "#000000",
    border = "Bottom: thin, Right: thin",
    borderColour = "#000000",
    fgFill = "#FFFFFF",
    bgFill = "#FFFFFF",
  )  
  
  # Funcion -----------------------------------------------------------------
  
  
  #Funcion para crear las variables produccion_total, ventas_total y personal_total
  
  contr_sum_an <- function(tabla){
    tabla1 <- tabla %>% summarise(produccion_total = sum(PRODUCCIONREALPOND),
                                  ventas_total=sum(VENTASREALESPOND),
                                  personal_total=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    return(tabla1)
  }
  
  
  #Funcion para crear las variavles pro,vent,per de acuerdo al periodo que se
  #esté manejando
  
  contr_fin <- function(periodo,tabla){
    if(periodo==6){
      contribucion <- tabla %>%
        summarise(prod = sum(PRODUCCIONREALPOND),
                  vent=sum(VENTASREALESPOND),
                  per=sum(PERSONAL)) %>%
        group_by(INCLUSION_NOMBRE_DEPTO)  %>%
        summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
                  ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
                  personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
        arrange(produccion)
      
    }else{
      contribucion <- tabla %>%
        summarise(prod = sum(PRODUCCIONREALPOND),
                  vent=sum(VENTASREALESPOND),
                  per=sum(PERSONAL)) %>%
        group_by(DOMINIO_39,DOMINIO39_DESCRIP) %>%
        summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
                  ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
                  personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
        arrange(produccion)
    }
    
    return(contribucion)
  }
  
  
  #Funcion para realizar los pivotes en las tablas de acuerdo al periodo que
  #se esté trabajando
  tabla_pivot <- function(periodo,tabla){
    if(periodo==5){
      tabla1 <- tabla %>%
        pivot_wider(names_from = c("ANIO2"),
                    values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
    }else {
      tabla1 <- tabla %>%
        pivot_wider(names_from = c("ANIO"),
                    values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
    }
    return(tabla1)
  }
  
  
  #Funcion para crear las nuevas variables concatenando columnas específicas del
  # data frame, según el periodo que se esté manejando
  tabla_paste_an <- function(periodo,base){
    if(periodo==11){
      base[paste0("varprodnom_",anio)] <- (base[paste0("produccionNom_",anio)]-base[paste0("produccionNom_",2019)])/base[paste0("produccionNom_",2019)]
      base[paste0("varprod_",anio)] <- (base[paste0("produccion_",anio)]-base[paste0("produccion_",2019)])/base[paste0("produccion_",2019)]
      base[paste0("varventasnom_",anio)]<- (base[paste0("ventasNom_",anio)]-base[paste0("ventasNom_",2019)])/base[paste0("ventasNom_",2019)]
      base[paste0("varventas_",anio)]<- (base[paste0("ventas_",anio)]-base[paste0("ventas_",2019)])/base[paste0("ventas_",2019)]
      base[paste0("varpersonas_",anio)] <- (base[paste0("personas_",anio)]-base[paste0("personas_",2019)])/base[paste0("personas_",2019)]
    }else if(periodo==2 | periodo==4 | periodo==6){
      base[paste0("varprod_",anio)] <- (base[paste0("produccion_",anio)]-base[paste0("produccion_",anio-1)])/base[paste0("produccion_",anio-1)]
      base[paste0("varventas_",anio)]<- (base[paste0("ventas_",anio)]-base[paste0("ventas_",anio-1)])/base[paste0("ventas_",anio-1)]
      base[paste0("varpersonas_",anio)] <- (base[paste0("personas_",anio)]-base[paste0("personas_",anio-1)])/base[paste0("personas_",anio-1)]
      base[paste0("varempleo_",anio)] <- (base[paste0("empleo_",anio)]-base[paste0("empleo_",anio-1)])/base[paste0("empleo_",anio-1)]
      base[paste0("varemptem_",anio)] <- (base[paste0("emptem_",anio)]-base[paste0("emptem_",anio-1)])/base[paste0("emptem_",anio-1)]
      base[paste0("varempleados_",anio)] <- (base[paste0("empleados_",anio)]-base[paste0("empleados_",anio-1)])/base[paste0("empleados_",anio-1)]
      base[paste0("varoperarios_",anio)] <- (base[paste0("operarios_",anio)]-base[paste0("operarios_",anio-1)])/base[paste0("operarios_",anio-1)]
      base[paste0("varsueldos_",anio)] <- (base[paste0("sueldos_",anio)]-base[paste0("sueldos_",anio-1)])/base[paste0("sueldos_",anio-1)]
      base[paste0("varsuelemplper_",anio)] <- (base[paste0("suelemplper_",anio)]-base[paste0("suelemplper_",anio-1)])/base[paste0("suelemplper_",anio-1)]
      base[paste0("varsuelempltem_",anio)] <- (base[paste0("suelempltem_",anio)]-base[paste0("suelempltem_",anio-1)])/base[paste0("suelempltem_",anio-1)]
      base[paste0("varsueltotemp_",anio)] <- (base[paste0("sueltotemp_",anio)]-base[paste0("sueltotemp_",anio-1)])/base[paste0("sueltotemp_",anio-1)]
      base[paste0("varsueltotope_",anio)] <- (base[paste0("sueltotope_",anio)]-base[paste0("sueltotope_",anio-1)])/base[paste0("sueltotope_",anio-1)]
      base[paste0("varhoras_",anio)] <- (base[paste0("horas_",anio)]-base[paste0("horas_",anio-1)])/base[paste0("horas_",anio-1)]
    }else{
      base[paste0("varprodnom_",anio)] <- (base[paste0("produccionNom_",anio)]-base[paste0("produccionNom_",anio-1)])/base[paste0("produccionNom_",anio-1)]
      base[paste0("varprod_",anio)] <- (base[paste0("produccion_",anio)]-base[paste0("produccion_",anio-1)])/base[paste0("produccion_",anio-1)]
      base[paste0("varventasnom_",anio)]<- (base[paste0("ventasNom_",anio)]-base[paste0("ventasNom_",anio-1)])/base[paste0("ventasNom_",anio-1)]
      base[paste0("varventas_",anio)]<- (base[paste0("ventas_",anio)]-base[paste0("ventas_",anio-1)])/base[paste0("ventas_",anio-1)]
      base[paste0("varpersonas_",anio)] <- (base[paste0("personas_",anio)]-base[paste0("personas_",anio-1)])/base[paste0("personas_",anio-1)]
    }
    
    return(base)
  }
  
  #Funcion para crear variables en las tablas finales,
  #según el periodo que se esté manejando
  tabla_summarise <- function(periodo,tabla){
    if(periodo==2 | periodo==4 | periodo==6){
      tabla1 <- tabla  %>%
        summarise(produccion=sum(PRODUCCIONREALPOND),
                  ventas = sum(VENTASREALESPOND),
                  personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
                  empleo=sum(TOTALEMPLEOPERMANENTE),
                  emptem=sum(TOTALEMPLEOTEMPORAL),
                  empleados=sum(TOTALEMPLEOADMON),
                  operarios=sum(TOTALEMPLEOPRODUC),
                  sueldos=sum(TOTALSUELDOSREALES),
                  suelemplper=sum(SUELDOSPERMANENTESREALES),
                  suelempltem=sum(SUELDOSTEMPORALESREALES),
                  sueltotemp=sum(SUELDOSADMONREAL),
                  sueltotope=sum(SUELDOSPRODUCREAL),
                  horas=sum(TOTALHORAS))
      
    }else{
      tabla1 <- tabla  %>%
        summarise(produccionNom=sum(PRODUCCIONNOMPOND),
                  produccion=sum(PRODUCCIONREALPOND),
                  ventasNom = sum(VENTASNOMINPOND),
                  ventas = sum(VENTASREALESPOND),
                  personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    }
    return(tabla1)
    
  }
  
  tabla_acople<-function(tabla){
    tabla$ANIO <- sapply(strsplit(as.character(tabla$variables), "_"), `[`, 2)
    tabla$variables <- sapply(strsplit(as.character(tabla$variables), "_"), `[`, 1)
    tabla <- tabla %>% filter(gsub("var","",variables)!=variables )
    
    tabla <- tabla %>% pivot_wider(names_from = variables,values_from = value)
    
    return(tabla)
  }
  
  # data$PRODUCCIONREALPOND<- round(data$PRODUCCIONREALPOND,2)
  # data$VENTASREALESPOND <- round(data$VENTASREALESPOND,2)
  
  # 1. Var y cont_Anual -----------------------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <-data %>%
    filter(MES==mes & ANIO%in%c(anio-1)) %>%
    summarise(produccion_total = sum(PRODUCCIONREALPOND),
              ventas_total=sum(VENTASREALESPOND),
              personal_total=sum(TOTPERS))
  
  #contribucion_total <- contr_sum_an(contribucion_total)
  
  #Calculo de la contribucion por meses
  contribucion <- data %>%
    filter(MES==mes & ANIO%in%c(anio,anio-1)) %>%
    #mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
    group_by(ANIO,MES,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(prod = sum(PRODUCCIONREALPOND),
              vent=sum(VENTASREALESPOND),
              per=sum(TOTPERS)) %>%
    group_by(DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
              ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
              personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
    arrange(produccion)
  
  #contribucion <- contr_fin(1,contribucion)
  
  #Calculo de la variación por dominos
  tabla1 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%mes) %>%
    group_by(ANIO,MES,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccionNom=sum(PRODUCCIONNOMPOND),
              produccion=sum(PRODUCCIONREALPOND),
              ventasNom = sum(VENTASNOMINPOND),
              ventas = sum(VENTASREALESPOND),
              #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)
              personas=sum(TOTPERS))
  
  #tabla1 <- tabla_summarise(1,tabla1)
  
  tabla1 <- tabla1 %>%
    pivot_wider(names_from = c("ANIO"),
                values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
  
  
  #tabla1 <- tabla_pivot(1,tabla1)
  
  tabla1[paste0("varprodnom_",anio)] <- (tabla1[paste0("produccionNom_",anio)]-tabla1[paste0("produccionNom_",anio-1)])/tabla1[paste0("produccionNom_",anio-1)]
  tabla1[paste0("varprod_",anio)] <- (tabla1[paste0("produccion_",anio)]-tabla1[paste0("produccion_",anio-1)])/tabla1[paste0("produccion_",anio-1)]
  tabla1[paste0("varventasnom_",anio)]<- (tabla1[paste0("ventasNom_",anio)]-tabla1[paste0("ventasNom_",anio-1)])/tabla1[paste0("ventasNom_",anio-1)]
  tabla1[paste0("varventas_",anio)]<- (tabla1[paste0("ventas_",anio)]-tabla1[paste0("ventas_",anio-1)])/tabla1[paste0("ventas_",anio-1)]
  tabla1[paste0("varpersonas_",anio)] <- (tabla1[paste0("personas_",anio)]-tabla1[paste0("personas_",anio-1)])/tabla1[paste0("personas_",anio-1)]
  
  #tabla1 <- tabla_paste_an(1,tabla1)
  
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:3)],names_to = "variables",values_to = "value" )
  
  #tabla1 <- tabla_acople(tabla1)
  tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
  tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
  tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
  tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
  
  
  #Empalme de la variacion y contribucion anual por dominios
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("DOMINIO_39","DOMINIO39_DESCRIP"))
  tabla1 <- tabla1[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprodnom",
                      "varprod","produccion","varventasnom","varventas",
                      "ventas","varpersonas","personal")]
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  tabla1 <- tabla1 %>% arrange(DOMINIO_39)
  
  #Total Industria
  
  #Calculo de la contribucion por meses
  contribucion1 <- data %>%
    filter(MES==mes & ANIO%in%c(anio,anio-1)) %>%
    #mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
    group_by(ANIO,MES) %>%
    summarise(prod = sum(PRODUCCIONREALPOND),
              vent=sum(VENTASREALESPOND),
              per=sum(TOTPERS))
  contribucion1["DOMINIO39_DESCRIP"] <- ""
  contribucion1["DOMINIO_39"] <- "Total Industia"
  contribucion1 <- contribucion1 %>%
    group_by(DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
              ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
              personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
    arrange(produccion)
  
  #contribucion <- contr_fin(1,contribucion)
  
  #Calculo de la variación por dominos
  tabla2 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%mes) %>%
    group_by(ANIO,MES) %>%
    summarise(produccionNom=sum(PRODUCCIONNOMPOND),
              produccion=sum(PRODUCCIONREALPOND),
              ventasNom = sum(VENTASNOMINPOND),
              ventas = sum(VENTASREALESPOND),
              #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)
              personas=sum(TOTPERS))
  
  #tabla1 <- tabla_summarise(1,tabla1)
  
  tabla2 <- tabla2 %>%
    pivot_wider(names_from = c("ANIO"),
                values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
  
  
  #tabla1 <- tabla_pivot(1,tabla1)
  
  tabla2[paste0("varprodnom_",anio)] <- (tabla2[paste0("produccionNom_",anio)]-tabla2[paste0("produccionNom_",anio-1)])/tabla2[paste0("produccionNom_",anio-1)]
  tabla2[paste0("varprod_",anio)] <- (tabla2[paste0("produccion_",anio)]-tabla2[paste0("produccion_",anio-1)])/tabla2[paste0("produccion_",anio-1)]
  tabla2[paste0("varventasnom_",anio)]<- (tabla2[paste0("ventasNom_",anio)]-tabla2[paste0("ventasNom_",anio-1)])/tabla2[paste0("ventasNom_",anio-1)]
  tabla2[paste0("varventas_",anio)]<- (tabla2[paste0("ventas_",anio)]-tabla2[paste0("ventas_",anio-1)])/tabla2[paste0("ventas_",anio-1)]
  tabla2[paste0("varpersonas_",anio)] <- (tabla2[paste0("personas_",anio)]-tabla2[paste0("personas_",anio-1)])/tabla2[paste0("personas_",anio-1)]
  
  #tabla1 <- tabla_paste_an(1,tabla1)
  
  tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1:3)],names_to = "variables",values_to = "value" )
  
  #tabla1 <- tabla_acople(tabla1)
  tabla2$ANIO <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 2)
  tabla2$variables <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 1)
  tabla2 <- tabla2 %>% filter(gsub("var","",variables)!=variables )
  tabla2 <- tabla2 %>% pivot_wider(names_from = variables,values_from = value)
  
  tabla2["DOMINIO39_DESCRIP"] <- ""
  tabla2["DOMINIO_39"] <- "Total Industia"
  
  #Empalme de la variacion y contribucion anual por dominios
  tabla2 <- inner_join(x=tabla2,y=contribucion1,by=c("DOMINIO_39","DOMINIO39_DESCRIP"))
  tabla2 <- tabla2[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprodnom",
                      "varprod","produccion","varventasnom","varventas",
                      "ventas","varpersonas","personal")]
  
  for( i in c("varprodnom","varprod","produccion","varventasnom",
              "varventas","ventas","varpersonas","personal")){
    tabla2[,i] <-  tabla2[,i]*100
  }
  
  tabla1 <- tabla1 %>% arrange(DOMINIO_39)
  tabla1 <- rbind(tabla2,tabla1)
  
  #Exportar
  
  sheet <- sheets[2]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,3:ncol(tabla1)]),
                      startRow = 13, startCol = 3,colNames=FALSE, rowNames=FALSE)
  
  addSuperSubScriptToCell_general(wb,sheet,row=9,col=1,texto = paste0(meses_enu[mes]," (",anio," / ",anio-1,")[p]"))
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 55, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  
  # 2. Var Anual_Emp_Sueldos_Horas  -----------------------------------------
  
  #Calculo de la variación por dominos
  tabla1 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%mes) %>%
    group_by(ANIO,MES,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccion=sum(PRODUCCIONREALPOND),
              ventas = sum(VENTASREALESPOND),
              #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas=sum(TOTPERS),
              empleo=sum(TOTALEMPLEOPERMANENTE),
              emptem=sum(TOTALEMPLEOTEMPORAL),
              empleados=sum(TOTALEMPLEOADMON),
              operarios=sum(TOTALEMPLEOPRODUC),
              sueldos=sum(TOTALSUELDOSREALES),
              suelemplper=sum(SUELDOSPERMANENTESREALES),
              suelempltem=sum(SUELDOSTEMPORALESREALES),
              sueltotemp=sum(SUELDOSADMONREAL),
              sueltotope=sum(SUELDOSPRODUCREAL),
              horas=sum(TOTALHORAS))
  
  
  #tabla1 <- tabla_summarise(2,tabla1)
  
  
  tabla1 <- tabla1 %>%
    pivot_wider(names_from = c("ANIO"),
                values_from = c("produccion","ventas","personas","empleo","emptem",
                                "empleados","operarios","sueldos","suelemplper","suelempltem","sueltotemp",
                                "sueltotope","horas"))
  
  
  #tabla1 <- tabla_paste_an(2,tabla1)
  
  tabla1[paste0("varprod_",anio)] <- (tabla1[paste0("produccion_",anio)]-tabla1[paste0("produccion_",anio-1)])/tabla1[paste0("produccion_",anio-1)]
  tabla1[paste0("varventas_",anio)]<- (tabla1[paste0("ventas_",anio)]-tabla1[paste0("ventas_",anio-1)])/tabla1[paste0("ventas_",anio-1)]
  tabla1[paste0("varpersonas_",anio)] <- (tabla1[paste0("personas_",anio)]-tabla1[paste0("personas_",anio-1)])/tabla1[paste0("personas_",anio-1)]
  tabla1[paste0("varempleo_",anio)] <- (tabla1[paste0("empleo_",anio)]-tabla1[paste0("empleo_",anio-1)])/tabla1[paste0("empleo_",anio-1)]
  tabla1[paste0("varemptem_",anio)] <- (tabla1[paste0("emptem_",anio)]-tabla1[paste0("emptem_",anio-1)])/tabla1[paste0("emptem_",anio-1)]
  tabla1[paste0("varempleados_",anio)] <- (tabla1[paste0("empleados_",anio)]-tabla1[paste0("empleados_",anio-1)])/tabla1[paste0("empleados_",anio-1)]
  tabla1[paste0("varoperarios_",anio)] <- (tabla1[paste0("operarios_",anio)]-tabla1[paste0("operarios_",anio-1)])/tabla1[paste0("operarios_",anio-1)]
  tabla1[paste0("varsueldos_",anio)] <- (tabla1[paste0("sueldos_",anio)]-tabla1[paste0("sueldos_",anio-1)])/tabla1[paste0("sueldos_",anio-1)]
  tabla1[paste0("varsuelemplper_",anio)] <- (tabla1[paste0("suelemplper_",anio)]-tabla1[paste0("suelemplper_",anio-1)])/tabla1[paste0("suelemplper_",anio-1)]
  tabla1[paste0("varsuelempltem_",anio)] <- (tabla1[paste0("suelempltem_",anio)]-tabla1[paste0("suelempltem_",anio-1)])/tabla1[paste0("suelempltem_",anio-1)]
  tabla1[paste0("varsueltotemp_",anio)] <- (tabla1[paste0("sueltotemp_",anio)]-tabla1[paste0("sueltotemp_",anio-1)])/tabla1[paste0("sueltotemp_",anio-1)]
  tabla1[paste0("varsueltotope_",anio)] <- (tabla1[paste0("sueltotope_",anio)]-tabla1[paste0("sueltotope_",anio-1)])/tabla1[paste0("sueltotope_",anio-1)]
  tabla1[paste0("varhoras_",anio)] <- (tabla1[paste0("horas_",anio)]-tabla1[paste0("horas_",anio-1)])/tabla1[paste0("horas_",anio-1)]
  
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:3)],names_to = "variables",values_to = "value" )
  
  
  #tabla1 <- tabla_acople(tabla1)
  tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
  tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
  tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
  tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
  
  
  #Empalme de la variacion y contribucion anual por dominios
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("DOMINIO_39","DOMINIO39_DESCRIP"))
  tabla1 <- tabla1[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprod",
                      "varventas","varpersonas","varempleo","varemptem",
                      "varempleados","varoperarios","varsueldos","varsuelemplper",
                      "varsuelempltem","varsueltotemp","varsueltotope","varhoras")]
  
  for( i in c("varprod","varventas","varpersonas","varempleo","varemptem",
              "varempleados","varoperarios","varsueldos","varsuelemplper",
              "varsuelempltem","varsueltotemp","varsueltotope","varhoras")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  #Totales Industria
  
  #Calculo de la variación por dominos
  tabla2 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%mes) %>%
    group_by(ANIO,MES) %>%
    summarise(produccion=sum(PRODUCCIONREALPOND),
              ventas = sum(VENTASREALESPOND),
              #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas=sum(TOTPERS),
              empleo=sum(TOTALEMPLEOPERMANENTE),
              emptem=sum(TOTALEMPLEOTEMPORAL),
              empleados=sum(TOTALEMPLEOADMON),
              operarios=sum(TOTALEMPLEOPRODUC),
              sueldos=sum(TOTALSUELDOSREALES),
              suelemplper=sum(SUELDOSPERMANENTESREALES),
              suelempltem=sum(SUELDOSTEMPORALESREALES),
              sueltotemp=sum(SUELDOSADMONREAL),
              sueltotope=sum(SUELDOSPRODUCREAL),
              horas=sum(TOTALHORAS))
  
  
  #tabla1 <- tabla_summarise(2,tabla1)
  
  
  tabla2 <- tabla2 %>%
    pivot_wider(names_from = c("ANIO"),
                values_from = c("produccion","ventas","personas","empleo","emptem",
                                "empleados","operarios","sueldos","suelemplper","suelempltem","sueltotemp",
                                "sueltotope","horas"))
  
  
  #tabla1 <- tabla_paste_an(2,tabla1)
  
  tabla2[paste0("varprod_",anio)] <- (tabla2[paste0("produccion_",anio)]-tabla2[paste0("produccion_",anio-1)])/tabla2[paste0("produccion_",anio-1)]
  tabla2[paste0("varventas_",anio)]<- (tabla2[paste0("ventas_",anio)]-tabla2[paste0("ventas_",anio-1)])/tabla2[paste0("ventas_",anio-1)]
  tabla2[paste0("varpersonas_",anio)] <- (tabla2[paste0("personas_",anio)]-tabla2[paste0("personas_",anio-1)])/tabla2[paste0("personas_",anio-1)]
  tabla2[paste0("varempleo_",anio)] <- (tabla2[paste0("empleo_",anio)]-tabla2[paste0("empleo_",anio-1)])/tabla2[paste0("empleo_",anio-1)]
  tabla2[paste0("varemptem_",anio)] <- (tabla2[paste0("emptem_",anio)]-tabla2[paste0("emptem_",anio-1)])/tabla2[paste0("emptem_",anio-1)]
  tabla2[paste0("varempleados_",anio)] <- (tabla2[paste0("empleados_",anio)]-tabla2[paste0("empleados_",anio-1)])/tabla2[paste0("empleados_",anio-1)]
  tabla2[paste0("varoperarios_",anio)] <- (tabla2[paste0("operarios_",anio)]-tabla2[paste0("operarios_",anio-1)])/tabla2[paste0("operarios_",anio-1)]
  tabla2[paste0("varsueldos_",anio)] <- (tabla2[paste0("sueldos_",anio)]-tabla2[paste0("sueldos_",anio-1)])/tabla2[paste0("sueldos_",anio-1)]
  tabla2[paste0("varsuelemplper_",anio)] <- (tabla2[paste0("suelemplper_",anio)]-tabla2[paste0("suelemplper_",anio-1)])/tabla2[paste0("suelemplper_",anio-1)]
  tabla2[paste0("varsuelempltem_",anio)] <- (tabla2[paste0("suelempltem_",anio)]-tabla2[paste0("suelempltem_",anio-1)])/tabla2[paste0("suelempltem_",anio-1)]
  tabla2[paste0("varsueltotemp_",anio)] <- (tabla2[paste0("sueltotemp_",anio)]-tabla2[paste0("sueltotemp_",anio-1)])/tabla2[paste0("sueltotemp_",anio-1)]
  tabla2[paste0("varsueltotope_",anio)] <- (tabla2[paste0("sueltotope_",anio)]-tabla2[paste0("sueltotope_",anio-1)])/tabla2[paste0("sueltotope_",anio-1)]
  tabla2[paste0("varhoras_",anio)] <- (tabla2[paste0("horas_",anio)]-tabla2[paste0("horas_",anio-1)])/tabla2[paste0("horas_",anio-1)]
  
  tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1:3)],names_to = "variables",values_to = "value" )
  
  
  #tabla2 <- tabla_acople(tabla2)
  tabla2$ANIO <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 2)
  tabla2$variables <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 1)
  tabla2 <- tabla2 %>% filter(gsub("var","",variables)!=variables )
  tabla2 <- tabla2 %>% pivot_wider(names_from = variables,values_from = value)
  
  tabla2["DOMINIO39_DESCRIP"] <- ""
  tabla2["DOMINIO_39"] <- "Total Industia"
  
  #Empalme de la variacion y contribucion anual por dominios
  tabla2 <- inner_join(x=tabla2,y=contribucion1,by=c("DOMINIO_39","DOMINIO39_DESCRIP"))
  tabla2 <- tabla2[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprod",
                      "varventas","varpersonas","varempleo","varemptem",
                      "varempleados","varoperarios","varsueldos","varsuelemplper",
                      "varsuelempltem","varsueltotemp","varsueltotope","varhoras")]
  
  for( i in c("varprod","varventas","varpersonas","varempleo","varemptem",
              "varempleados","varoperarios","varsueldos","varsuelemplper",
              "varsuelempltem","varsueltotemp","varsueltotope","varhoras")){
    tabla2[,i] <-  tabla2[,i]*100
  }
  
  tabla1 <- tabla1 %>%
    arrange(DOMINIO_39)
  
  tabla1 <- rbind(tabla2,tabla1)
  
  #Exportar
  
  sheet <- sheets[3]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,3:ncol(tabla1)]),
                      startRow = 14, startCol = 3,colNames=FALSE, rowNames=FALSE)
  
  addSuperSubScriptToCell_general(wb,sheet,row=9,col=1,texto = paste0(meses_enu[mes]," (",anio," / ",anio-1,")[p]"))

  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,data.frame(Enunciado),startRow = 55, startCol = 1,
                      colNames=FALSE, rowNames=FALSE)
  
  
  # 3. Var y cont_anio corrido -----------------------------------------------
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio-1)) %>%
    summarise(produccion_total = sum(PRODUCCIONREALPOND),
              ventas_total=sum(VENTASREALESPOND),
              personal_total=sum(TOTPERS))
  
  #contribucion_total <- contr_sum_an(contribucion_total)
  
  #Calculo de la contribucion mensual por dominio
  contribucion <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio,anio-1)) %>%
    #mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
    group_by(ANIO,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(prod = sum(PRODUCCIONREALPOND),
              vent=sum(VENTASREALESPOND),
              per=sum(TOTPERS)) %>%
    group_by(DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
              ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
              personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
    arrange(produccion)
  
  #contribucion <- contr_fin(3,contribucion)
  
  #Calculo de la variacion por dominio
  tabla1 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%c(1:mes)) %>%
    group_by(ANIO,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccionNom=sum(PRODUCCIONNOMPOND),
              produccion=sum(PRODUCCIONREALPOND),
              ventasNom = sum(VENTASNOMINPOND),
              ventas = sum(VENTASREALESPOND),
              #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas=sum(TOTPERS))
  
  #tabla1 <- tabla_summarise(3,tabla1)
  
  tabla1 <- tabla1 %>%
    pivot_wider(names_from = c("ANIO"),
                values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
  
  
  #tabla1 <- tabla_pivot(3,tabla1)
  
  #tabla1 <- tabla_paste_an(3,tabla1)
  
  tabla1[paste0("varprodnom_",anio)] <- (tabla1[paste0("produccionNom_",anio)]-tabla1[paste0("produccionNom_",anio-1)])/tabla1[paste0("produccionNom_",anio-1)]
  tabla1[paste0("varprod_",anio)] <- (tabla1[paste0("produccion_",anio)]-tabla1[paste0("produccion_",anio-1)])/tabla1[paste0("produccion_",anio-1)]
  tabla1[paste0("varventasnom_",anio)]<- (tabla1[paste0("ventasNom_",anio)]-tabla1[paste0("ventasNom_",anio-1)])/tabla1[paste0("ventasNom_",anio-1)]
  tabla1[paste0("varventas_",anio)]<- (tabla1[paste0("ventas_",anio)]-tabla1[paste0("ventas_",anio-1)])/tabla1[paste0("ventas_",anio-1)]
  tabla1[paste0("varpersonas_",anio)] <- (tabla1[paste0("personas_",anio)]-tabla1[paste0("personas_",anio-1)])/tabla1[paste0("personas_",anio-1)]
  
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:2)],names_to = "variables",values_to = "value" )
  
  #tabla1 <- tabla_acople(tabla1)
  tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
  tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
  tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
  tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
  
  
  
  #Empalme de la contribucion y variacion por dominio
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("DOMINIO_39","DOMINIO39_DESCRIP"))
  tabla1 <- tabla1[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprodnom","varprod","produccion",
                      "varventasnom","varventas","ventas","varpersonas","personal")]
  
  for( i in c("varprodnom","varprod","produccion",
              "varventasnom","varventas","ventas","varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  #Total Industria
  
  #contribucion_total <- contr_sum_an(contribucion_total)
  
  #Calculo de la contribucion mensual por dominio
  contribucion1 <- data %>%
    filter(MES%in%c(1:mes) & ANIO%in%c(anio,anio-1)) %>%
    #mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
    group_by(ANIO) %>%
    summarise(prod = sum(PRODUCCIONREALPOND),
              vent=sum(VENTASREALESPOND),
              per=sum(TOTPERS))
  contribucion1["DOMINIO_39"] <- "Total Industria"
  contribucion1["DOMINIO39_DESCRIP"] <- ""
  contribucion1 <- contribucion1 %>%
    group_by(DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
              ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
              personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
    arrange(produccion)
  
  #contribucion <- contr_fin(3,contribucion)
  
  #Calculo de la variacion por dominio
  tabla2 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%c(1:mes)) %>%
    group_by(ANIO) %>%
    summarise(produccionNom=sum(PRODUCCIONNOMPOND),
              produccion=sum(PRODUCCIONREALPOND),
              ventasNom = sum(VENTASNOMINPOND),
              ventas = sum(VENTASREALESPOND),
              #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas=sum(TOTPERS))
  
  #tabla2 <- tabla_summarise(3,tabla2)
  
  tabla2 <- tabla2 %>%
    pivot_wider(names_from = c("ANIO"),
                values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
  
  
  #tabla2 <- tabla_pivot(3,tabla2)
  
  #tabla2 <- tabla_paste_an(3,tabla2)
  
  tabla2[paste0("varprodnom_",anio)] <- (tabla2[paste0("produccionNom_",anio)]-tabla2[paste0("produccionNom_",anio-1)])/tabla2[paste0("produccionNom_",anio-1)]
  tabla2[paste0("varprod_",anio)] <- (tabla2[paste0("produccion_",anio)]-tabla2[paste0("produccion_",anio-1)])/tabla2[paste0("produccion_",anio-1)]
  tabla2[paste0("varventasnom_",anio)]<- (tabla2[paste0("ventasNom_",anio)]-tabla2[paste0("ventasNom_",anio-1)])/tabla2[paste0("ventasNom_",anio-1)]
  tabla2[paste0("varventas_",anio)]<- (tabla2[paste0("ventas_",anio)]-tabla2[paste0("ventas_",anio-1)])/tabla2[paste0("ventas_",anio-1)]
  tabla2[paste0("varpersonas_",anio)] <- (tabla2[paste0("personas_",anio)]-tabla2[paste0("personas_",anio-1)])/tabla2[paste0("personas_",anio-1)]
  
  tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1:2)],names_to = "variables",values_to = "value" )
  
  #tabla2 <- tabla_acople(tabla2)
  tabla2$ANIO <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 2)
  tabla2$variables <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 1)
  tabla2 <- tabla2 %>% filter(gsub("var","",variables)!=variables )
  tabla2 <- tabla2 %>% pivot_wider(names_from = variables,values_from = value)
  
  tabla2["DOMINIO_39"] <- "Total Industria"
  tabla2["DOMINIO39_DESCRIP"] <- ""
  
  #Empalme de la contribucion y variacion por dominio
  tabla2 <- inner_join(x=tabla2,y=contribucion1,by=c("DOMINIO_39","DOMINIO39_DESCRIP"))
  tabla2 <- tabla2[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprodnom","varprod","produccion",
                      "varventasnom","varventas","ventas","varpersonas","personal")]
  
  for( i in c("varprodnom","varprod","produccion",
              "varventasnom","varventas","ventas","varpersonas","personal")){
    tabla2[,i] <-  tabla2[,i]*100
  }
  
  tabla1 <- tabla1 %>%
    arrange(DOMINIO_39)
  
  tabla1 <- rbind(tabla2,tabla1)
  
  #Exportar
  
  sheet <- sheets[5]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,3:ncol(tabla1)]),
                      startRow = 13, startCol = 3,colNames=FALSE, rowNames=FALSE)
  
  addSuperSubScriptToCell_general(wb,sheet,row=9,col=1,texto = paste0("Enero - ",meses_min[mes]," (",anio," / ",anio-1,")[p]"))
  
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 55, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  
  # 4. Var anio corr Emp_Sueldos_Hor -----------------------------------------
  
  #Calculo de la variación por dominos
  tabla1 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%c(1:mes)) %>%
    group_by(ANIO,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccion=sum(PRODUCCIONREALPOND),
              ventas = sum(VENTASREALESPOND),
              personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              empleo=sum(TOTALEMPLEOPERMANENTE),
              emptem=sum(TOTALEMPLEOTEMPORAL),
              empleados=sum(TOTALEMPLEOADMON),
              operarios=sum(TOTALEMPLEOPRODUC),
              sueldos=sum(TOTALSUELDOSREALES),
              suelemplper=sum(SUELDOSPERMANENTESREALES),
              suelempltem=sum(SUELDOSTEMPORALESREALES),
              sueltotemp=sum(SUELDOSADMONREAL),
              sueltotope=sum(SUELDOSPRODUCREAL),
              horas=sum(TOTALHORAS))
  
  #tabla1 <- tabla_summarise(4,tabla1)
  
  
  tabla1 <- tabla1 %>%
    pivot_wider(names_from = c("ANIO"),
                values_from = c("produccion","ventas","personas","empleo",
                                "emptem","empleados","operarios","sueldos",
                                "suelemplper","suelempltem","sueltotemp",
                                "sueltotope","horas"))
  
  #tabla1 <- tabla_paste_an(4,tabla1)
  
  tabla1[paste0("varprod_",anio)] <- (tabla1[paste0("produccion_",anio)]-tabla1[paste0("produccion_",anio-1)])/tabla1[paste0("produccion_",anio-1)]
  tabla1[paste0("varventas_",anio)]<- (tabla1[paste0("ventas_",anio)]-tabla1[paste0("ventas_",anio-1)])/tabla1[paste0("ventas_",anio-1)]
  tabla1[paste0("varpersonas_",anio)] <- (tabla1[paste0("personas_",anio)]-tabla1[paste0("personas_",anio-1)])/tabla1[paste0("personas_",anio-1)]
  tabla1[paste0("varempleo_",anio)] <- (tabla1[paste0("empleo_",anio)]-tabla1[paste0("empleo_",anio-1)])/tabla1[paste0("empleo_",anio-1)]
  tabla1[paste0("varemptem_",anio)] <- (tabla1[paste0("emptem_",anio)]-tabla1[paste0("emptem_",anio-1)])/tabla1[paste0("emptem_",anio-1)]
  tabla1[paste0("varempleados_",anio)] <- (tabla1[paste0("empleados_",anio)]-tabla1[paste0("empleados_",anio-1)])/tabla1[paste0("empleados_",anio-1)]
  tabla1[paste0("varoperarios_",anio)] <- (tabla1[paste0("operarios_",anio)]-tabla1[paste0("operarios_",anio-1)])/tabla1[paste0("operarios_",anio-1)]
  tabla1[paste0("varsueldos_",anio)] <- (tabla1[paste0("sueldos_",anio)]-tabla1[paste0("sueldos_",anio-1)])/tabla1[paste0("sueldos_",anio-1)]
  tabla1[paste0("varsuelemplper_",anio)] <- (tabla1[paste0("suelemplper_",anio)]-tabla1[paste0("suelemplper_",anio-1)])/tabla1[paste0("suelemplper_",anio-1)]
  tabla1[paste0("varsuelempltem_",anio)] <- (tabla1[paste0("suelempltem_",anio)]-tabla1[paste0("suelempltem_",anio-1)])/tabla1[paste0("suelempltem_",anio-1)]
  tabla1[paste0("varsueltotemp_",anio)] <- (tabla1[paste0("sueltotemp_",anio)]-tabla1[paste0("sueltotemp_",anio-1)])/tabla1[paste0("sueltotemp_",anio-1)]
  tabla1[paste0("varsueltotope_",anio)] <- (tabla1[paste0("sueltotope_",anio)]-tabla1[paste0("sueltotope_",anio-1)])/tabla1[paste0("sueltotope_",anio-1)]
  tabla1[paste0("varhoras_",anio)] <- (tabla1[paste0("horas_",anio)]-tabla1[paste0("horas_",anio-1)])/tabla1[paste0("horas_",anio-1)]
  
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:3)],names_to = "variables",values_to = "value" )
  
  #tabla1 <- tabla_acople(tabla1)
  tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
  tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
  tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
  tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
  
  
  tabla1 <- tabla1[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprod",
                      "varventas","varpersonas","varempleo","varemptem",
                      "varempleados","varoperarios","varsueldos","varsuelemplper",
                      "varsuelempltem","varsueltotemp","varsueltotope","varhoras")]
  
  for( i in c("varprod","varventas","varpersonas","varempleo","varemptem",
              "varempleados","varoperarios","varsueldos","varsuelemplper",
              "varsuelempltem","varsueltotemp","varsueltotope","varhoras")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  
  #Total Industria
  
  #Calculo de la variación por dominos
  tabla2 <- data %>%
    filter(ANIO%in%c(anio,anio-1) & MES%in%c(1:mes)) %>%
    group_by(ANIO) %>%
    summarise(produccion=sum(PRODUCCIONREALPOND),
              ventas = sum(VENTASREALESPOND),
              personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              empleo=sum(TOTALEMPLEOPERMANENTE),
              emptem=sum(TOTALEMPLEOTEMPORAL),
              empleados=sum(TOTALEMPLEOADMON),
              operarios=sum(TOTALEMPLEOPRODUC),
              sueldos=sum(TOTALSUELDOSREALES),
              suelemplper=sum(SUELDOSPERMANENTESREALES),
              suelempltem=sum(SUELDOSTEMPORALESREALES),
              sueltotemp=sum(SUELDOSADMONREAL),
              sueltotope=sum(SUELDOSPRODUCREAL),
              horas=sum(TOTALHORAS))
  
  #tabla2 <- tabla_summarise(4,tabla2)
  
  
  tabla2 <- tabla2 %>%
    pivot_wider(names_from = c("ANIO"),
                values_from = c("produccion","ventas","personas","empleo",
                                "emptem","empleados","operarios","sueldos",
                                "suelemplper","suelempltem","sueltotemp",
                                "sueltotope","horas"))
  
  #tabla2 <- tabla_paste_an(4,tabla2)
  
  tabla2[paste0("varprod_",anio)] <- (tabla2[paste0("produccion_",anio)]-tabla2[paste0("produccion_",anio-1)])/tabla2[paste0("produccion_",anio-1)]
  tabla2[paste0("varventas_",anio)]<- (tabla2[paste0("ventas_",anio)]-tabla2[paste0("ventas_",anio-1)])/tabla2[paste0("ventas_",anio-1)]
  tabla2[paste0("varpersonas_",anio)] <- (tabla2[paste0("personas_",anio)]-tabla2[paste0("personas_",anio-1)])/tabla2[paste0("personas_",anio-1)]
  tabla2[paste0("varempleo_",anio)] <- (tabla2[paste0("empleo_",anio)]-tabla2[paste0("empleo_",anio-1)])/tabla2[paste0("empleo_",anio-1)]
  tabla2[paste0("varemptem_",anio)] <- (tabla2[paste0("emptem_",anio)]-tabla2[paste0("emptem_",anio-1)])/tabla2[paste0("emptem_",anio-1)]
  tabla2[paste0("varempleados_",anio)] <- (tabla2[paste0("empleados_",anio)]-tabla2[paste0("empleados_",anio-1)])/tabla2[paste0("empleados_",anio-1)]
  tabla2[paste0("varoperarios_",anio)] <- (tabla2[paste0("operarios_",anio)]-tabla2[paste0("operarios_",anio-1)])/tabla2[paste0("operarios_",anio-1)]
  tabla2[paste0("varsueldos_",anio)] <- (tabla2[paste0("sueldos_",anio)]-tabla2[paste0("sueldos_",anio-1)])/tabla2[paste0("sueldos_",anio-1)]
  tabla2[paste0("varsuelemplper_",anio)] <- (tabla2[paste0("suelemplper_",anio)]-tabla2[paste0("suelemplper_",anio-1)])/tabla2[paste0("suelemplper_",anio-1)]
  tabla2[paste0("varsuelempltem_",anio)] <- (tabla2[paste0("suelempltem_",anio)]-tabla2[paste0("suelempltem_",anio-1)])/tabla2[paste0("suelempltem_",anio-1)]
  tabla2[paste0("varsueltotemp_",anio)] <- (tabla2[paste0("sueltotemp_",anio)]-tabla2[paste0("sueltotemp_",anio-1)])/tabla2[paste0("sueltotemp_",anio-1)]
  tabla2[paste0("varsueltotope_",anio)] <- (tabla2[paste0("sueltotope_",anio)]-tabla2[paste0("sueltotope_",anio-1)])/tabla2[paste0("sueltotope_",anio-1)]
  tabla2[paste0("varhoras_",anio)] <- (tabla2[paste0("horas_",anio)]-tabla2[paste0("horas_",anio-1)])/tabla2[paste0("horas_",anio-1)]
  
  tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1:3)],names_to = "variables",values_to = "value" )
  
  #tabla2 <- tabla_acople(tabla2)
  tabla2$ANIO <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 2)
  tabla2$variables <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 1)
  tabla2 <- tabla2 %>% filter(gsub("var","",variables)!=variables )
  tabla2 <- tabla2 %>% pivot_wider(names_from = variables,values_from = value)
  
  tabla2["DOMINIO_39"] <- "Total Industria"
  tabla2["DOMINIO39_DESCRIP"] <- ""
  
  tabla2 <- tabla2[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprod",
                      "varventas","varpersonas","varempleo","varemptem",
                      "varempleados","varoperarios","varsueldos","varsuelemplper",
                      "varsuelempltem","varsueltotemp","varsueltotope","varhoras")]
  
  for( i in c("varprod","varventas","varpersonas","varempleo","varemptem",
              "varempleados","varoperarios","varsueldos","varsuelemplper",
              "varsuelempltem","varsueltotemp","varsueltotope","varhoras")){
    tabla2[,i] <-  tabla2[,i]*100
  }
  
  tabla1 <- tabla1 %>% arrange(DOMINIO_39)
  
  tabla1 <- rbind(tabla2,tabla1)
  
  
  #Exportar
  
  sheet <- sheets[6]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,3:ncol(tabla1)]),
                      startRow = 14, startCol = 3,colNames=FALSE, rowNames=FALSE)
  
  addSuperSubScriptToCell_general(wb,sheet,row=9,col=1,texto = paste0("Enero - ",meses_min[mes]," (",anio," / ",anio-1,")[p]"))
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 55, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 5. Var y cont_doce meses ------------------------------------------------
  
  #Creacion de la variable para anio corrido
  data$ANIO2 <- as.numeric(ifelse(data$MES%in%c((mes+1):12),data$ANIO+1,data$ANIO))
  
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(ANIO2%in%(anio-1)) %>%
    summarise(produccion_total = sum(PRODUCCIONREALPOND),
              ventas_total=sum(VENTASREALESPOND),
              personal_total=sum(TOTPERS))
  
  #contribucion_total <- contr_sum_an(contribucion_total)
  
  #Calculo de la contribucion mensual por dominio
  contribucion <- data %>%
    filter(ANIO2%in%c(anio-1,anio)) %>%
    #mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
    group_by(ANIO2,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(prod = sum(PRODUCCIONREALPOND),
              vent=sum(VENTASREALESPOND),
              per=sum(TOTPERS)) %>%
    group_by(DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
              ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
              personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
    arrange(produccion)
  
  #contribucion <- contr_fin(5,contribucion)
  
  #Calculo de la variación por dominio
  tabla1 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccionNom=sum(PRODUCCIONNOMPOND),
              produccion=sum(PRODUCCIONREALPOND),
              ventasNom = sum(VENTASNOMINPOND),
              ventas = sum(VENTASREALESPOND),
              #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)
              personas=sum(TOTPERS))
  
  #tabla1 <- tabla_summarise(5,tabla1)
  
  tabla1 <- tabla1 %>%
    pivot_wider(names_from = c("ANIO2"),
                values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
  
  
  #tabla1 <- tabla_pivot(5,tabla1)
  
  #tabla1 <- tabla_paste_an(5,tabla1)
  
  tabla1[paste0("varprodnom_",anio)] <- (tabla1[paste0("produccionNom_",anio)]-tabla1[paste0("produccionNom_",anio-1)])/tabla1[paste0("produccionNom_",anio-1)]
  tabla1[paste0("varprod_",anio)] <- (tabla1[paste0("produccion_",anio)]-tabla1[paste0("produccion_",anio-1)])/tabla1[paste0("produccion_",anio-1)]
  tabla1[paste0("varventasnom_",anio)]<- (tabla1[paste0("ventasNom_",anio)]-tabla1[paste0("ventasNom_",anio-1)])/tabla1[paste0("ventasNom_",anio-1)]
  tabla1[paste0("varventas_",anio)]<- (tabla1[paste0("ventas_",anio)]-tabla1[paste0("ventas_",anio-1)])/tabla1[paste0("ventas_",anio-1)]
  tabla1[paste0("varpersonas_",anio)] <- (tabla1[paste0("personas_",anio)]-tabla1[paste0("personas_",anio-1)])/tabla1[paste0("personas_",anio-1)]
  
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:3)],names_to = "variables",values_to = "value" )
  
  #tabla1 <- tabla_acople(tabla1)
  tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
  tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
  tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
  tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
  
  
  #Empalme de la contribucion y la variacion
  tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("DOMINIO_39","DOMINIO39_DESCRIP"))
  tabla1 <- tabla1[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  for( i in c("varprodnom",
              "varprod","produccion","varventasnom","varventas","ventas",
              "varpersonas","personal")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  #Total Industria
  
  #Calculo de la contribucion mensual por dominio
  contribucion1 <- data %>%
    filter(ANIO2%in%c(anio-1,anio)) %>%
    #mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
    group_by(ANIO2) %>%
    summarise(prod = sum(PRODUCCIONREALPOND),
              vent=sum(VENTASREALESPOND),
              per=sum(TOTPERS))
  contribucion1["DOMINIO_39"] <- "Total Industria"
  contribucion1["DOMINIO39_DESCRIP"] <- ""
  contribucion1 <- contribucion1 %>%
    group_by(DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
              ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
              personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
    arrange(produccion)
  
  #contribucion <- contr_fin(5,contribucion)
  
  #Calculo de la variación por dominio
  tabla2 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2) %>%
    summarise(produccionNom=sum(PRODUCCIONNOMPOND),
              produccion=sum(PRODUCCIONREALPOND),
              ventasNom = sum(VENTASNOMINPOND),
              ventas = sum(VENTASREALESPOND),
              #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)
              personas=sum(TOTPERS))
  
  #tabla2 <- tabla_summarise(5,tabla2)
  
  tabla2 <- tabla2 %>%
    pivot_wider(names_from = c("ANIO2"),
                values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
  
  
  #tabla2 <- tabla_pivot(5,tabla2)
  
  #tabla2 <- tabla_paste_an(5,tabla2)
  
  tabla2[paste0("varprodnom_",anio)] <- (tabla2[paste0("produccionNom_",anio)]-tabla2[paste0("produccionNom_",anio-1)])/tabla2[paste0("produccionNom_",anio-1)]
  tabla2[paste0("varprod_",anio)] <- (tabla2[paste0("produccion_",anio)]-tabla2[paste0("produccion_",anio-1)])/tabla2[paste0("produccion_",anio-1)]
  tabla2[paste0("varventasnom_",anio)]<- (tabla2[paste0("ventasNom_",anio)]-tabla2[paste0("ventasNom_",anio-1)])/tabla2[paste0("ventasNom_",anio-1)]
  tabla2[paste0("varventas_",anio)]<- (tabla2[paste0("ventas_",anio)]-tabla2[paste0("ventas_",anio-1)])/tabla2[paste0("ventas_",anio-1)]
  tabla2[paste0("varpersonas_",anio)] <- (tabla2[paste0("personas_",anio)]-tabla2[paste0("personas_",anio-1)])/tabla2[paste0("personas_",anio-1)]
  
  tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1:3)],names_to = "variables",values_to = "value" )
  
  #tabla2 <- tabla_acople(tabla2)
  tabla2$ANIO <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 2)
  tabla2$variables <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 1)
  tabla2 <- tabla2 %>% filter(gsub("var","",variables)!=variables )
  tabla2 <- tabla2 %>% pivot_wider(names_from = variables,values_from = value)
  
  
  tabla2["DOMINIO_39"] <- "Total Industria"
  tabla2["DOMINIO39_DESCRIP"] <- ""
  
  #Empalme de la contribucion y la variacion
  tabla2 <- inner_join(x=tabla2,y=contribucion1,by=c("DOMINIO_39","DOMINIO39_DESCRIP"))
  tabla2 <- tabla2[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprodnom",
                      "varprod","produccion","varventasnom","varventas","ventas",
                      "varpersonas","personal")]
  
  for( i in c("varprodnom",
              "varprod","produccion","varventasnom","varventas","ventas",
              "varpersonas","personal")){
    tabla2[,i] <-  tabla2[,i]*100
  }
  
  
  tabla1 <- tabla1 %>% arrange(DOMINIO_39)
  tabla1 <- rbind(tabla2,tabla1)
  
  
  
  #Exportar
  
  sheet <- sheets[7]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,3:ncol(tabla1)]),
                      startRow = 13, startCol = 3,colNames=FALSE, rowNames=FALSE)
  
  
if(mes==12){
  addSuperSubScriptToCell_general(wb,sheet,row=9,col=1,texto = paste0(meses_enu[1]," ",anio," - ",meses_min[mes]," ",anio," / ",meses_min[1]," ",anio-1," - ",meses_min[mes]," ",anio-1,"[p]"))
}else{
  addSuperSubScriptToCell_general(wb,sheet,row=9,col=1,texto = paste0(meses_enu[mes+1]," ",anio-1," - ",meses_min[mes]," ",anio," / ",meses_min[mes+1]," ",anio-2," - ",meses_min[mes]," ",anio-1,"[p]"))
}
  
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 55, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  
  # 6. Var doce meses Emp_Sueldos_H -----------------------------------------
  
  #Creacion de la variable para anio corrido
  data$ANIO2 <- as.numeric(ifelse(data$MES%in%c((mes+1):12),data$ANIO+1,data$ANIO))
  
  #Calculo de la contribucion total
  contribucion_total <- data %>%
    filter(ANIO2%in%(anio-1)) %>%
    summarise(produccion_total = sum(PRODUCCIONREALPOND),
              ventas_total=sum(VENTASREALESPOND),
              personal_total=sum(TOTPERS))
  
  #contribucion_total <- contr_sum_an(contribucion_total)
  
  #Calculo de la contribucion mensual por departamento
  contribucion <- data %>%
    filter(ANIO2%in%c(anio-1,anio)) %>%
    #mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
    group_by(ANIO2,INCLUSION_NOMBRE_DEPTO) %>%
    summarise(prod = sum(PRODUCCIONREALPOND),
              vent=sum(VENTASREALESPOND),
              per=sum(TOTPERS)) %>%
    group_by(INCLUSION_NOMBRE_DEPTO)  %>%
    summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
              ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
              personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
    arrange(produccion)
  
  
  #contribucion <- contr_fin(6,contribucion)
  
  #Calculo de la variacion
  tabla1 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccion=sum(PRODUCCIONREALPOND),
              ventas = sum(VENTASREALESPOND),
              #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas=sum(TOTPERS),
              empleo=sum(TOTALEMPLEOPERMANENTE),
              emptem=sum(TOTALEMPLEOTEMPORAL),
              empleados=sum(TOTALEMPLEOADMON),
              operarios=sum(TOTALEMPLEOPRODUC),
              sueldos=sum(TOTALSUELDOSREALES),
              suelemplper=sum(SUELDOSPERMANENTESREALES),
              suelempltem=sum(SUELDOSTEMPORALESREALES),
              sueltotemp=sum(SUELDOSADMONREAL),
              sueltotope=sum(SUELDOSPRODUCREAL),
              horas=sum(TOTALHORAS))
  
  #tabla1 <- tabla_summarise(6,tabla1)
  
  tabla1_6 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccionNom=sum(PRODUCCIONNOMPOND),
              ventasNom = sum(VENTASNOMINPOND))
  
  #Empalme de la variacion y contribucion
  tabla1 <- tabla1 %>% left_join(tabla1_6,by=c("ANIO2"="ANIO2","DOMINIO_39"="DOMINIO_39",
                                               "DOMINIO39_DESCRIP"="DOMINIO39_DESCRIP"))
  
  tabla1 <- tabla1 %>%
    pivot_wider(names_from = c("ANIO2"),
                values_from = c("produccionNom","produccion","ventas","ventasNom",
                                "personas","empleo","emptem","empleados","operarios",
                                "sueldos","suelemplper","suelempltem","sueltotemp",
                                "sueltotope","horas"))
  
  
  #tabla1 <- tabla_paste_an(6,tabla1)
  
  tabla1[paste0("varprod_",anio)] <- (tabla1[paste0("produccion_",anio)]-tabla1[paste0("produccion_",anio-1)])/tabla1[paste0("produccion_",anio-1)]
  tabla1[paste0("varventas_",anio)]<- (tabla1[paste0("ventas_",anio)]-tabla1[paste0("ventas_",anio-1)])/tabla1[paste0("ventas_",anio-1)]
  tabla1[paste0("varpersonas_",anio)] <- (tabla1[paste0("personas_",anio)]-tabla1[paste0("personas_",anio-1)])/tabla1[paste0("personas_",anio-1)]
  tabla1[paste0("varempleo_",anio)] <- (tabla1[paste0("empleo_",anio)]-tabla1[paste0("empleo_",anio-1)])/tabla1[paste0("empleo_",anio-1)]
  tabla1[paste0("varemptem_",anio)] <- (tabla1[paste0("emptem_",anio)]-tabla1[paste0("emptem_",anio-1)])/tabla1[paste0("emptem_",anio-1)]
  tabla1[paste0("varempleados_",anio)] <- (tabla1[paste0("empleados_",anio)]-tabla1[paste0("empleados_",anio-1)])/tabla1[paste0("empleados_",anio-1)]
  tabla1[paste0("varoperarios_",anio)] <- (tabla1[paste0("operarios_",anio)]-tabla1[paste0("operarios_",anio-1)])/tabla1[paste0("operarios_",anio-1)]
  tabla1[paste0("varsueldos_",anio)] <- (tabla1[paste0("sueldos_",anio)]-tabla1[paste0("sueldos_",anio-1)])/tabla1[paste0("sueldos_",anio-1)]
  tabla1[paste0("varsuelemplper_",anio)] <- (tabla1[paste0("suelemplper_",anio)]-tabla1[paste0("suelemplper_",anio-1)])/tabla1[paste0("suelemplper_",anio-1)]
  tabla1[paste0("varsuelempltem_",anio)] <- (tabla1[paste0("suelempltem_",anio)]-tabla1[paste0("suelempltem_",anio-1)])/tabla1[paste0("suelempltem_",anio-1)]
  tabla1[paste0("varsueltotemp_",anio)] <- (tabla1[paste0("sueltotemp_",anio)]-tabla1[paste0("sueltotemp_",anio-1)])/tabla1[paste0("sueltotemp_",anio-1)]
  tabla1[paste0("varsueltotope_",anio)] <- (tabla1[paste0("sueltotope_",anio)]-tabla1[paste0("sueltotope_",anio-1)])/tabla1[paste0("sueltotope_",anio-1)]
  tabla1[paste0("varhoras_",anio)] <- (tabla1[paste0("horas_",anio)]-tabla1[paste0("horas_",anio-1)])/tabla1[paste0("horas_",anio-1)]
  
  
  
  tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:2)],names_to = "variables",values_to = "value" )
  
  
  #tabla1 <- tabla_acople(tabla1)
  tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
  tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
  tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
  tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
  
  
  
  
  tabla1 <- tabla1[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprod",
                      "varventas","varpersonas","varempleo","varemptem",
                      "varempleados","varoperarios","varsueldos","varsuelemplper",
                      "varsuelempltem","varsueltotemp","varsueltotope","varhoras")]
  
  for( i in c("varprod","varventas","varpersonas","varempleo","varemptem",
              "varempleados","varoperarios","varsueldos","varsuelemplper",
              "varsuelempltem","varsueltotemp","varsueltotope","varhoras")){
    tabla1[,i] <-  tabla1[,i]*100
  }
  
  
  #Total Industria
  
  #contribucion <- contr_fin(6,contribucion)
  
  #Calculo de la variacion
  tabla2 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2) %>%
    summarise(produccion=sum(PRODUCCIONREALPOND),
              ventas = sum(VENTASREALESPOND),
              #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas=sum(TOTPERS),
              empleo=sum(TOTALEMPLEOPERMANENTE),
              emptem=sum(TOTALEMPLEOTEMPORAL),
              empleados=sum(TOTALEMPLEOADMON),
              operarios=sum(TOTALEMPLEOPRODUC),
              sueldos=sum(TOTALSUELDOSREALES),
              suelemplper=sum(SUELDOSPERMANENTESREALES),
              suelempltem=sum(SUELDOSTEMPORALESREALES),
              sueltotemp=sum(SUELDOSADMONREAL),
              sueltotope=sum(SUELDOSPRODUCREAL),
              horas=sum(TOTALHORAS))
  
  #tabla2 <- tabla_summarise(6,tabla2)
  
  tabla2_6 <- data %>%
    filter(ANIO2%in%c(anio,anio-1)) %>%
    group_by(ANIO2) %>%
    summarise(produccionNom=sum(PRODUCCIONNOMPOND),
              ventasNom = sum(VENTASNOMINPOND))
  
  tabla2["DOMINIO_39"] <- "Total Industria"
  tabla2["DOMINIO39_DESCRIP"] <- ""
  tabla2_6["DOMINIO_39"] <- "Total Industria"
  tabla2_6["DOMINIO39_DESCRIP"] <- ""
  
  #Empalme de la variacion y contribucion
  tabla2 <- tabla2 %>% left_join(tabla2_6,by=c("ANIO2"="ANIO2","DOMINIO_39"="DOMINIO_39",
                                               "DOMINIO39_DESCRIP"="DOMINIO39_DESCRIP"))
  
  tabla2 <- tabla2 %>%
    pivot_wider(names_from = c("ANIO2"),
                values_from = c("produccionNom","produccion","ventas","ventasNom",
                                "personas","empleo","emptem","empleados","operarios",
                                "sueldos","suelemplper","suelempltem","sueltotemp",
                                "sueltotope","horas"))
  
  
  #tabla2 <- tabla_paste_an(6,tabla2)
  
  tabla2[paste0("varprod_",anio)] <- (tabla2[paste0("produccion_",anio)]-tabla2[paste0("produccion_",anio-1)])/tabla2[paste0("produccion_",anio-1)]
  tabla2[paste0("varventas_",anio)]<- (tabla2[paste0("ventas_",anio)]-tabla2[paste0("ventas_",anio-1)])/tabla2[paste0("ventas_",anio-1)]
  tabla2[paste0("varpersonas_",anio)] <- (tabla2[paste0("personas_",anio)]-tabla2[paste0("personas_",anio-1)])/tabla2[paste0("personas_",anio-1)]
  tabla2[paste0("varempleo_",anio)] <- (tabla2[paste0("empleo_",anio)]-tabla2[paste0("empleo_",anio-1)])/tabla2[paste0("empleo_",anio-1)]
  tabla2[paste0("varemptem_",anio)] <- (tabla2[paste0("emptem_",anio)]-tabla2[paste0("emptem_",anio-1)])/tabla2[paste0("emptem_",anio-1)]
  tabla2[paste0("varempleados_",anio)] <- (tabla2[paste0("empleados_",anio)]-tabla2[paste0("empleados_",anio-1)])/tabla2[paste0("empleados_",anio-1)]
  tabla2[paste0("varoperarios_",anio)] <- (tabla2[paste0("operarios_",anio)]-tabla2[paste0("operarios_",anio-1)])/tabla2[paste0("operarios_",anio-1)]
  tabla2[paste0("varsueldos_",anio)] <- (tabla2[paste0("sueldos_",anio)]-tabla2[paste0("sueldos_",anio-1)])/tabla2[paste0("sueldos_",anio-1)]
  tabla2[paste0("varsuelemplper_",anio)] <- (tabla2[paste0("suelemplper_",anio)]-tabla2[paste0("suelemplper_",anio-1)])/tabla2[paste0("suelemplper_",anio-1)]
  tabla2[paste0("varsuelempltem_",anio)] <- (tabla2[paste0("suelempltem_",anio)]-tabla2[paste0("suelempltem_",anio-1)])/tabla2[paste0("suelempltem_",anio-1)]
  tabla2[paste0("varsueltotemp_",anio)] <- (tabla2[paste0("sueltotemp_",anio)]-tabla2[paste0("sueltotemp_",anio-1)])/tabla2[paste0("sueltotemp_",anio-1)]
  tabla2[paste0("varsueltotope_",anio)] <- (tabla2[paste0("sueltotope_",anio)]-tabla2[paste0("sueltotope_",anio-1)])/tabla2[paste0("sueltotope_",anio-1)]
  tabla2[paste0("varhoras_",anio)] <- (tabla2[paste0("horas_",anio)]-tabla2[paste0("horas_",anio-1)])/tabla2[paste0("horas_",anio-1)]
  
  
  
  tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1:2)],names_to = "variables",values_to = "value" )
  
  
  #tabla2 <- tabla_acople(tabla2)
  tabla2$ANIO <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 2)
  tabla2$variables <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 1)
  tabla2 <- tabla2 %>% filter(gsub("var","",variables)!=variables )
  tabla2 <- tabla2 %>% pivot_wider(names_from = variables,values_from = value)
  
  
  
  
  tabla2 <- tabla2[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprod",
                      "varventas","varpersonas","varempleo","varemptem",
                      "varempleados","varoperarios","varsueldos","varsuelemplper",
                      "varsuelempltem","varsueltotemp","varsueltotope","varhoras")]
  
  for( i in c("varprod","varventas","varpersonas","varempleo","varemptem",
              "varempleados","varoperarios","varsueldos","varsuelemplper",
              "varsuelempltem","varsueltotemp","varsueltotope","varhoras")){
    tabla2[,i] <-  tabla2[,i]*100
  }
  
  tabla1 <- tabla1 %>% arrange(DOMINIO_39)
  tabla1 <- rbind(tabla2,tabla1)
  
  #Exportar
  
  sheet <- sheets[8]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1[,3:ncol(tabla1)]),
                      startRow = 14, startCol = 3,colNames=FALSE, rowNames=FALSE)
  
  if(mes==12){
    addSuperSubScriptToCell_general(wb,sheet,row=9,col=1,texto = paste0(meses_enu[1]," ",anio," - ",meses_min[mes]," ",anio," / ",meses_min[1]," ",anio-1," - ",meses_min[mes]," ",anio-1,"[p]"))
  }else{
    addSuperSubScriptToCell_general(wb,sheet,row=9,col=1,texto = paste0(meses_enu[mes+1]," ",anio-1," - ",meses_min[mes]," ",anio," / ",meses_min[mes+1]," ",anio-2," - ",meses_min[mes]," ",anio-1,"[p]"))
  }
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = 55, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  # 7. Indices total por clase  ---------------------------------------------
  
  #Creacion  de las variables nominales y totales y con estas se realiza el
  #calculo de la contribucion mensual por dominio
  contribucion_mensual <- data %>%
    group_by(ANIO,MES,DOMINIO_39,DOMINIO39_DESCRIP) %>%
    summarise(produccionNom_mensual=sum(PRODUCCIONNOMPOND),
              produccion_mensual=sum(PRODUCCIONREALPOND),
              ventasNom_mensual = sum(VENTASNOMINPOND),
              ventas_mensual = sum(VENTASREALESPOND),
              #personas_mensual=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas_mensual=sum(TOTPERS),
              empleo_mensual=sum(TOTALEMPLEOPERMANENTE),
              emptem_mensual=sum(TOTALEMPLEOTEMPORAL),
              empleados_mensual=sum(TOTALEMPLEOADMON),
              operarios_mensual=sum(TOTALEMPLEOPRODUC),
              sueldos_mensual=sum(TOTALSUELDOSREALES),
              suelemplper_mensual=sum(SUELDOSPERMANENTESREALES),
              suelempltem_mensual=sum(SUELDOSTEMPORALESREALES),
              sueltotemp_mensual=sum(SUELDOSADMONREAL),
              sueltotope_mensual=sum(SUELDOSPRODUCREAL),
              horas_mensual=sum(TOTALHORAS))
  
  #Calculo de la contribucion con el anio base
  contribucion_base<- contribucion_mensual %>%
    filter(ANIO==2018) %>%
    group_by(ANIO,DOMINIO_39) %>%
    summarise(produccionNom_total=mean(produccionNom_mensual),
              produccion_total=mean(produccion_mensual),
              ventasNom_total = mean(ventasNom_mensual),
              ventas_total = mean(ventas_mensual),
              personas_total=mean(personas_mensual),
              empleo_total=mean(empleo_mensual),
              emptem_total=mean(emptem_mensual),
              empleados_total=mean(empleados_mensual),
              operarios_total=mean(operarios_mensual),
              sueldos_total=mean(sueldos_mensual),
              suelemplper_total=mean(suelemplper_mensual),
              suelempltem_total=mean(suelempltem_mensual),
              sueltotemp_total=mean(sueltotemp_mensual),
              sueltotope_total=mean(sueltotope_mensual),
              horas_total=mean(horas_mensual))
  
  
  contribucion_base<-subset(contribucion_base, select = -ANIO)
  
  
  contribucion<-contribucion_mensual %>%
    left_join(contribucion_base,by=c("DOMINIO_39"="DOMINIO_39"))
  
  #Calculo de la variacion
  tabla1<-contribucion %>%
    mutate(produccionNom =(produccionNom_mensual/produccionNom_total)*100,
           produccion    =(produccion_mensual/produccion_total)*100,
           ventasNom     =(ventasNom_mensual/ventasNom_total)*100,
           ventas        =(ventas_mensual/ventas_total)*100,
           personas      =(personas_mensual/personas_total)*100,
           empleo        =(empleo_mensual/empleo_total)*100,
           emptem        =(emptem_mensual/emptem_total)*100,
           empleados     =(empleados_mensual/empleados_total)*100,
           operarios     =(operarios_mensual/operarios_total)*100,
           sueldos       =(sueldos_mensual/sueldos_total)*100,
           suelemplper   =(suelemplper_mensual/suelemplper_total)*100,
           suelempltem   =(suelempltem_mensual/suelempltem_total)*100,
           sueltotemp    =(sueltotemp_mensual/sueltotemp_total)*100,
           sueltotope    =(sueltotope_mensual/sueltotope_total)*100,
           horas         =(horas_mensual/horas_total)*100) %>%
    select(DOMINIO_39,ANIO,MES,DOMINIO39_DESCRIP,produccionNom,
           produccion,ventasNom,ventas,personas,empleo,emptem,empleados,
           operarios,sueldos,suelemplper,suelempltem,sueltotemp,sueltotope,
           horas)
  
  #Calculo de la contribucion mensual
  contribucion_mensual <- data %>%
    group_by(ANIO,MES) %>%
    summarise(produccionNom_mensual=sum(PRODUCCIONNOMPOND),
              produccion_mensual=sum(PRODUCCIONREALPOND),
              ventasNom_mensual = sum(VENTASNOMINPOND),
              ventas_mensual = sum(VENTASREALESPOND),
              #personas_mensual=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas_mensual=sum(TOTPERS),
              empleo_mensual=sum(TOTALEMPLEOPERMANENTE),
              emptem_mensual=sum(TOTALEMPLEOTEMPORAL),
              empleados_mensual=sum(TOTALEMPLEOADMON),
              operarios_mensual=sum(TOTALEMPLEOPRODUC),
              sueldos_mensual=sum(TOTALSUELDOSREALES),
              suelemplper_mensual=sum(SUELDOSPERMANENTESREALES),
              suelempltem_mensual=sum(SUELDOSTEMPORALESREALES),
              sueltotemp_mensual=sum(SUELDOSADMONREAL),
              sueltotope_mensual=sum(SUELDOSPRODUCREAL),
              horas_mensual=sum(TOTALHORAS))
  
  #Calculo de la contribucion por el anio base
  contribucion_base <- contribucion_mensual %>%
    filter(ANIO==2018) %>%
    group_by(ANIO) %>%
    summarise(produccionNom_total=mean(produccionNom_mensual),
              produccion_total=mean(produccion_mensual),
              ventasNom_total= mean(ventasNom_mensual),
              ventas_total= mean(ventas_mensual),
              personas_total=mean(personas_mensual),
              empleo_total=mean(empleo_mensual),
              emptem_total=mean(emptem_mensual),
              empleados_total=mean(empleados_mensual),
              operarios_total=mean(operarios_mensual),
              sueldos_total=mean(sueldos_mensual),
              suelemplper_total=mean(suelemplper_mensual),
              suelempltem_total=mean(suelempltem_mensual),
              sueltotemp_total=mean(sueltotemp_mensual),
              sueltotope_total=mean(sueltotope_mensual),
              horas_total=mean(horas_mensual))
  
  contribucion_base<-subset(contribucion_base, select = -ANIO)
  
  contribucion_mensual<-contribucion_mensual %>%
    mutate(produccionNom_total=contribucion_base$produccionNom_total,
           produccion_total=contribucion_base$produccion_total,
           ventasNom_total=contribucion_base$ventasNom_total,
           ventas_total=contribucion_base$ventas_total,
           personas_total=contribucion_base$personas_total,
           empleo_total=contribucion_base$empleo_total,
           emptem_total=contribucion_base$emptem_total,
           empleados_total=contribucion_base$empleados_total,
           operarios_total=contribucion_base$operarios_total,
           sueldos_total=contribucion_base$sueldos_total,
           suelemplper_total=contribucion_base$suelemplper_total,
           suelempltem_total=contribucion_base$suelempltem_total,
           sueltotemp_total=contribucion_base$sueltotemp_total,
           sueltotope_total=contribucion_base$sueltotope_total,
           horas_total=contribucion_base$horas_total)
  
  
  #Calculo de la variables nominales y totales
  tabla1_1<-contribucion_mensual %>%
    mutate(produccionNom =(produccionNom_mensual/produccionNom_total)*100,
           produccion    =(produccion_mensual/produccion_total)*100,
           ventasNom     =(ventasNom_mensual/ventasNom_total)*100,
           ventas        =(ventas_mensual/ventas_total)*100,
           personas      =(personas_mensual/personas_total)*100,
           empleo        =(empleo_mensual/empleo_total)*100,
           emptem        =(emptem_mensual/emptem_total)*100,
           empleados     =(empleados_mensual/empleados_total)*100,
           operarios     =(operarios_mensual/operarios_total)*100,
           sueldos       =(sueldos_mensual/sueldos_total)*100,
           suelemplper   =(suelemplper_mensual/suelemplper_total)*100,
           suelempltem   =(suelempltem_mensual/suelempltem_total)*100,
           sueltotemp    =(sueltotemp_mensual/sueltotemp_total)*100,
           sueltotope    =(sueltotope_mensual/sueltotope_total)*100,
           horas         =(horas_mensual/horas_total)*100) %>%
    select(ANIO,MES,produccionNom,produccion,ventasNom,ventas,personas,empleo,emptem,empleados,
           operarios,sueldos,suelemplper,suelempltem,sueltotemp,sueltotope,
           horas)
  tabla1_1["DOMINIO_39"]<-"T_IND"
  tabla1_1["DOMINIO39_DESCRIP"]<-"Total Industria"
  tabla1$DOMINIO_39<-as.character(tabla1$DOMINIO_39)
  tabla1<-rbind(tabla1,tabla1_1)
  tabla1$DOMINIO_39 <- ifelse(tabla1$DOMINIO_39=="T_IND",paste0(0,tabla1$DOMINIO_39),tabla1$DOMINIO_39)
  tabla1<-tabla1 %>% arrange(DOMINIO_39,ANIO,MES)
  tabla1$DOMINIO_39 <- ifelse(tabla1$DOMINIO_39=="0T_IND",substr(tabla1$DOMINIO_39,2,nchar(tabla1$DOMINIO_39)),tabla1$DOMINIO_39)
  tabla1_7 <- tabla1
  
  #Exportar
  
  sheet <- sheets[9]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1_7),
                      startRow = 12, startCol = 1,colNames=FALSE, rowNames=FALSE)
  

  addSuperSubScriptToCell_general(wb,sheet,row=7,col=1,texto = paste0("Enero 2018 - ",meses_enu[mes]," ",anio,"[p]"))
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1_7)+14), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fuente. DANE - EMMET")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1_7)+15), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("p: provisionales")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1_7)+16), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Nota: La diferencia entre el total y la suma de los dominios se debe a aproximaciones decimales.")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1_7)+17), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Nota: Se presentan cambios por actualización de información de parte de las fuentes informantes")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1_7)+18), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  
  filas_pares <- seq(12, nrow(tabla1_7)+12, by = 2)
  filas_impares <- seq(13, nrow(tabla1_7)+12, by = 2)
  addStyle(wb, sheet, style=colgr_in, rows = filas_pares, cols = 1:4, gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl_in, rows = filas_impares, cols = 1:4, gridExpand = TRUE)
  addStyle(wb, sheet, style=colgr, rows = filas_pares, cols = 5:ncol(tabla1_7), gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl, rows = filas_impares, cols = 5:ncol(tabla1_7), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcgr, rows = filas_pares, cols = ncol(tabla1_7), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = filas_impares, cols = ncol(tabla1_7), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultbl, rows = (nrow(tabla1_7)+12), cols = 1:ncol(tabla1_7), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultblfc, rows = (nrow(tabla1_7)+12), cols = ncol(tabla1_7), gridExpand = TRUE)
  
  addStyle(wb, sheet, style=rowbl, rows = (nrow(tabla1_7)+14), cols = 1:ncol(tabla1_7), gridExpand = TRUE)
  addStyle(wb, sheet, style=rowblf, rows = (nrow(tabla1_7)+14), cols = ncol(tabla1_7), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = (nrow(tabla1_7)+15):(nrow(tabla1_7)+17), cols = ncol(tabla1_7), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultrbl, rows = (nrow(tabla1_7)+18), cols = 1:ncol(tabla1_7), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultrblf, rows = (nrow(tabla1_7)+18), cols = ncol(tabla1_7), gridExpand = TRUE)
  
  # 9. Enlace legal hasta 2014 ----------------------------------------------
  
  #Creacion de una matriz con 12 filas y el mismo numero de columnas del archivo
  # indices del 2014
  N<-as.data.frame(matrix(NA,ncol = ncol(indices_14),nrow = 12),colnames=FALSE)
  colnames(N)<-colnames(indices_14)
  
  #Se une la matriz N con los índices, cada 60 datos
  in_prue<-NULL
  for(i in seq(from=60, to = nrow(indices_14), by =60)){
    if(i==60){
      in_prue<-rbind(indices_14[1:i,],N)
    }
    else{
      in_prue<-rbind(in_prue,indices_14[(i-59):i,],N)
    }
    
  }
  
  #Se acomoda de nuevo la base seleccionando algunas variables de in_prueba,
  # con otras seleccionadas de la base N
  indices_lag<-rbind(in_prue[13:nrow(in_prue),5:ncol(in_prue)],N[,5:ncol(N)])
  colnames(indices_lag) <- paste0("R", colnames(indices_lag))
  indices_lag<-cbind(in_prue,indices_lag)
  indices_lag<-indices_lag[-c((nrow(indices_lag)-11):nrow(indices_lag)),]
  
  #Creacion de las variables de la tabla anteriormente construida
  variacion<-within(indices_lag,{
    VHORASTOTALESTRABAJADAS=HORASTOTALESTRABAJADAS/RHORASTOTALESTRABAJADAS
    VSUELDOTOTALOPERARIOS  =SUELDOTOTALOPERARIOS/RSUELDOTOTALOPERARIOS
    VSUELDOTOTALEMPLEADOS  =SUELDOTOTALEMPLEADOS/RSUELDOTOTALEMPLEADOS
    VSUELDOEMPLEOTEMPORAL  =SUELDOEMPLEOTEMPORAL/RSUELDOEMPLEOTEMPORAL
    VSUELDOEMPLEOPERMANENTE=SUELDOEMPLEOPERMANENTE/RSUELDOEMPLEOPERMANENTE
    VTOTALSUELDOS      =TOTALSUELDOS/RTOTALSUELDOS
    VTOTALOPERARIOS    =TOTALOPERARIOS/RTOTALOPERARIOS
    VTOTALEMPLEADOS    =TOTALEMPLEADOS/RTOTALEMPLEADOS
    VPERSONALTEMPORAL  =PERSONALTEMPORAL/RPERSONALTEMPORAL
    VPERSONALPERMANENTE=PERSONALPERMANENTE/RPERSONALPERMANENTE
    VEMPLEOTOTAL       =EMPLEOTOTAL/REMPLEOTOTAL
    VVENTASREALES      =VENTASREALES/RVENTASREALES
    VVENTASNOMINALES   =VENTASNOMINALES/RVENTASNOMINALES
    VPRODUCCIONREAL    =PRODUCCIONREAL/RPRODUCCIONREAL
    VPRODUCCIONNOMINAL =PRODUCCIONNOMINAL/RPRODUCCIONNOMINAL
  })
  
  variacion <- variacion[!is.na(variacion$DOMINIOS),]
  
  #Se crea la base de los indices con el anio base, previamente calculada
  tabla2<-tabla1_7 %>%
    filter(ANIO==2018)
  tabla2<-as.data.frame(tabla2)
  colnames(tabla2)<-paste("I", colnames(variacion[,1:ncol(tabla2)]), sep = "")
  tabla2 <- select(tabla2, -"ICLASESINDUSTRIALES" )
  tabla2$IANO<-as.numeric(tabla2$IANO)
  
  #Empalme de las variaciones y los indices
  varia<-variacion %>% left_join(tabla2,by=c("DOMINIOS"="IDOMINIOS",
                                             "ANO"="IANO","MES"="IMES"))
  
  #Calculo de los enlaces
  varia<-varia %>%
    arrange(DOMINIOS,desc(ANO),desc(MES))
  
  vector<-unique(varia$DOMINIOS)
  variac<-NULL
  for (k in vector) {
    vari<-varia %>%
      filter(DOMINIOS==k)
    for (j in 50:ncol(varia)) {
      for(i in 1:48){
        vari[(i+12),j]=(vari[(i+12),(j-15)]*vari[i,j])
      }
    }
    variac<-rbind(variac,vari)
  }
  
  
  cols_i <- grep("^I", names(variac), value = TRUE)
  variac<- select(variac,c(DOMINIOS,ANO,MES,CLASESINDUSTRIALES,cols_i))
  colnames(variac)<-gsub("I","",names(variac))
  tabla1<-tabla1 %>% filter(ANIO!=2018)
  colnames(tabla1)<-colnames(variac)
  tabla1<-rbind(variac,tabla1)
  tabla1<-tabla1 %>% arrange(DOMNOS,ANO,MES)
  tabla1<-tabla1 %>% rename("DOMINIOS"="DOMNOS")
  tabla1$DOMINIOS <- ifelse(tabla1$DOMINIOS=="T_IND",paste0(0,tabla1$DOMINIOS),tabla1$DOMINIOS)
  tabla1<-tabla1 %>% arrange(DOMINIOS)
  tabla1$DOMINIOS <- ifelse(tabla1$DOMINIOS=="0T_IND",substr(tabla1$DOMINIOS,2,nchar(tabla1$DOMINIOS)),tabla1$DOMINIOS)
  
  tabla1$ANO <- as.character(tabla1$ANO)
  tabla1$MES <- as.character(tabla1$MES)
  
  #Exportar
  
  sheet <- sheets[12]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1),
                      startRow = 13, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  addSuperSubScriptToCell_general(wb,sheet,row=7,col=1,texto = paste0("Enero 2018 - ",meses_enu[mes]," ",anio,"[p]"))
  
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+15), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fuente. DANE - EMMET")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+16), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("p: provisionales")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+17), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Nota: La diferencia entre el total y la suma de los dominios se debe a aproximaciones decimales.")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+18), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Nota: Se presentan cambios por actualización de información de parte de las fuentes informantes")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+19), startCol = 1,colNames=FALSE, rowNames=FALSE)

  
  filas_pares <- seq(14, nrow(tabla1)+13, by = 2)
  filas_impares <- seq(13, nrow(tabla1)+13, by = 2)
  
  
  addStyle(wb, sheet, style=colgr_in, rows = filas_impares, cols = 1:4, gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl_in, rows = filas_pares, cols = 1:4, gridExpand = TRUE)
  addStyle(wb, sheet, style=colgr, rows = filas_impares, cols = 5:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl, rows = filas_pares, cols = 5:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcgr, rows = filas_impares, cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = filas_pares, cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultbl, rows = (nrow(tabla1)+13), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultblfc, rows = (nrow(tabla1)+13), cols = ncol(tabla1), gridExpand = TRUE)
  
  addStyle(wb, sheet, style=rowbl, rows = (nrow(tabla1)+15), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=rowblf, rows = (nrow(tabla1)+15), cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = (nrow(tabla1)+16):(nrow(tabla1)+18), cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultrbl, rows = (nrow(tabla1)+19), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultrblf, rows = (nrow(tabla1)+19), cols = ncol(tabla1), gridExpand = TRUE)
  
  # 10. Enlace legal Hasta 2001 ---------------------------------------------
  
  #Definicion de nuevos dominios
  data2<-data %>%
    filter(CLASE_CIIU4!=1922) %>%
    mutate(DOMINIO_39_ENLACE=
             ifelse(DOMINIO_39 %in% c(1050,1090), 1050,
                    ifelse(DOMINIO_39 %in% c(1089,1030,1082),"1089a",
                           ifelse(DOMINIO_39 %in% c(1300,1400),"1400b",
                                  ifelse(DOMINIO_39 %in% c(1900),"1900c",
                                         ifelse(DOMINIO_39 %in% c(2020,2023,2100),"2020d",DOMINIO_39)))))
    ) %>%
    mutate(DOMINIO39_DESCRIP_ENLACE=
             ifelse(
               DOMINIO_39_ENLACE==1050, "Elaboración de productos de molinería",
               ifelse(DOMINIO_39_ENLACE=="1089a" , "Resto de alimentos",
                      ifelse(DOMINIO_39_ENLACE=="1400b" , "Confección de prendas de vestir",
                             ifelse(DOMINIO_39_ENLACE=="2020d" , "Fabricación de otros productos químicos",
                                    ifelse(DOMINIO_39_ENLACE=="1900c" , "RefinaciC3n de petróleo",DOMINIO39_DESCRIP))))))
  
  data2$TOTPERS <- as.numeric(data2$TOTPERS)
  data2$TOTPERS <- ifelse(is.na(data2$TOTPERS),0,data2$TOTPERS)
  #Calculo de la contribucion mensual
  contribucion_mensual <- data2 %>%
    group_by(ANIO,MES,DOMINIO_39_ENLACE,DOMINIO39_DESCRIP_ENLACE) %>%
    summarise(produccionNom_mensual=sum(PRODUCCIONNOMPOND),
              produccion_mensual=sum(PRODUCCIONREALPOND),
              ventasNom_mensual = sum(VENTASNOMINPOND),
              ventas_mensual = sum(VENTASREALESPOND),
              personas_mensual = sum(TOTPERS),
              personal_admon=sum(TOTALEMPLEOADMON),
              personal_operario=sum(TOTALEMPLEOPRODUC))
  
  #Calculo de la contribucion con el anio base
  contribucion_base<- contribucion_mensual  %>%
    filter(ANIO==2018) %>%
    group_by(ANIO,DOMINIO_39_ENLACE) %>%
    summarise(produccionNom_total=mean(produccionNom_mensual),
              produccion_total=mean(produccion_mensual),
              ventasNom_total = mean(ventasNom_mensual),
              ventas_total = mean(ventas_mensual),
              personas_total=mean(personas_mensual),
              personal_admon_total=mean(personal_admon),
              personal_operario_total=mean(personal_operario))
  
  contribucion_base<-subset(contribucion_base, select = -ANIO)
  
  #Empalme de la contribucion base con la contribucion mensual
  contribucion<-contribucion_mensual %>%
    left_join(contribucion_base,by=c("DOMINIO_39_ENLACE"="DOMINIO_39_ENLACE"))
  
  #Creacion de variables
  tabla1<-contribucion %>%
    mutate(produccionNom =(produccionNom_mensual/produccionNom_total)*100,
           produccion    =(produccion_mensual/produccion_total)*100,
           ventasNom     =(ventasNom_mensual/ventasNom_total)*100,
           ventas        =(ventas_mensual/ventas_total)*100,
           personas      =(personas_mensual/personas_total)*100,
           admon         =(personal_admon/personal_admon_total)*100,
           operarios     =(personal_operario/personal_operario_total)*100) %>%
    select(DOMINIO_39_ENLACE,ANIO,MES,DOMINIO39_DESCRIP_ENLACE,produccionNom,
           produccion,ventasNom,ventas,personas,admon,operarios)
  
  #Calculo de variables de contribucion mensual
  
  data$TOTPERS <- as.numeric(data$TOTPERS)
  data$TOTPERS <- ifelse(is.na(data$TOTPERS),0,data$TOTPERS)
  
  contribucion_mensual <- data %>%
    group_by(ANIO,MES) %>%
    summarise(produccionNom_mensual=sum(PRODUCCIONNOMPOND),
              produccion_mensual=sum(PRODUCCIONREALPOND),
              ventasNom_mensual = sum(VENTASNOMINPOND),
              ventas_mensual = sum(VENTASREALESPOND),
              personas_mensual=sum(TOTPERS),
              personal_admon=sum(TOTALEMPLEOADMON),
              personal_operario=sum(TOTALEMPLEOPRODUC))
  
  #Calculo de variables de contribucion con el anio base
  contribucion_base <- contribucion_mensual %>%
    filter(ANIO==2018) %>%
    group_by(ANIO) %>%
    summarise(produccionNom_total=mean(produccionNom_mensual),
              produccion_total=mean(produccion_mensual),
              ventasNom_total= mean(ventasNom_mensual),
              ventas_total= mean(ventas_mensual),
              personas_total=mean(personas_mensual),
              personal_admon_total=mean(personal_admon),
              personal_operario_total=mean(personal_operario))
  
  #Creacion de variables nominales y totales
  contribucion_mensual<-contribucion_mensual %>%
    mutate(produccionNom_total=contribucion_base$produccionNom_total,
           produccion_total=contribucion_base$produccion_total,
           ventasNom_total=contribucion_base$ventasNom_total,
           ventas_total=contribucion_base$ventas_total,
           personas_total=contribucion_base$personas_total,
           personal_admon_total=contribucion_base$personal_admon_total,
           personal_operario_total=contribucion_base$personal_operario_total)
  
  #Calculo de variables nominales y totales
  tabla1_1<-contribucion_mensual %>%
    mutate(produccionNom =(produccionNom_mensual/produccionNom_total)*100,
           produccion    =(produccion_mensual/produccion_total)*100,
           ventasNom     =(ventasNom_mensual/ventasNom_total)*100,
           ventas        =(ventas_mensual/ventas_total)*100,
           personas      =(personas_mensual/personas_total)*100,
           admon         =(personal_admon/personal_admon_total)*100,
           operarios     =(personal_operario/personal_operario_total)*100) %>%
    select(ANIO,MES,produccionNom,
           produccion,ventasNom,ventas,personas,admon,operarios)
  tabla1_1["DOMINIO_39_ENLACE"]<-"T_IND"
  tabla1_1["DOMINIO39_DESCRIP_ENLACE"]<-"Total Industria"
  tabla1$DOMINIO_39_ENLACE<-as.character(tabla1$DOMINIO_39_ENLACE)
  tabla1<-rbind(tabla1,tabla1_1)
  
  #Indicides 2001 agrupados por ciius
  
  #Creacion  de las variables nominales y totales y con estas se realiza el
  #calculo de la contribucion mensual por dominio
  contribucion_mensual <- data2 %>%
    group_by(ANIO,MES,DOMINIO_39_ENLACE,DOMINIO39_DESCRIP_ENLACE) %>%
    summarise(produccionNom_mensual=sum(PRODUCCIONNOMPOND),
              produccion_mensual=sum(PRODUCCIONREALPOND),
              ventasNom_mensual = sum(VENTASNOMINPOND),
              ventas_mensual = sum(VENTASREALESPOND),
              #personas_mensual=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas_mensual=sum(TOTPERS),
              empleados_mensual=sum(TOTALEMPLEOADMON),
              operarios_mensual=sum(TOTALEMPLEOPRODUC))
  
  #Calculo de la contribucion con el anio base
  contribucion_base<- contribucion_mensual %>%
    filter(ANIO==2018) %>%
    group_by(ANIO,DOMINIO_39_ENLACE) %>%
    summarise(produccionNom_total=mean(produccionNom_mensual),
              produccion_total=mean(produccion_mensual),
              ventasNom_total = mean(ventasNom_mensual),
              ventas_total = mean(ventas_mensual),
              personas_total=mean(personas_mensual),
              empleados_total=mean(empleados_mensual),
              operarios_total=mean(operarios_mensual))
  
  
  contribucion_base<-subset(contribucion_base, select = -ANIO)
  
  
  contribucion<-contribucion_mensual %>%
    left_join(contribucion_base,by=c("DOMINIO_39_ENLACE"="DOMINIO_39_ENLACE"))
  
  #Calculo de la variacion
  tabla1a<-contribucion %>%
    mutate(produccionNom =(produccionNom_mensual/produccionNom_total)*100,
           produccion    =(produccion_mensual/produccion_total)*100,
           ventasNom     =(ventasNom_mensual/ventasNom_total)*100,
           ventas        =(ventas_mensual/ventas_total)*100,
           personas      =(personas_mensual/personas_total)*100,
           empleados     =(empleados_mensual/empleados_total)*100,
           operarios     =(operarios_mensual/operarios_total)*100) %>%
    select(DOMINIO_39_ENLACE,ANIO,MES,DOMINIO39_DESCRIP_ENLACE,produccionNom,
           produccion,ventasNom,ventas,personas,empleados,operarios)
  
  #Calculo de la contribucion mensual
  contribucion_mensual <- data %>%
    group_by(ANIO,MES) %>%
    summarise(produccionNom_mensual=sum(PRODUCCIONNOMPOND),
              produccion_mensual=sum(PRODUCCIONREALPOND),
              ventasNom_mensual = sum(VENTASNOMINPOND),
              ventas_mensual = sum(VENTASREALESPOND),
              #personas_mensual=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC),
              personas_mensual=sum(TOTPERS),
              empleados_mensual=sum(TOTALEMPLEOADMON),
              operarios_mensual=sum(TOTALEMPLEOPRODUC))
  
  #Calculo de la contribucion por el anio base
  contribucion_base <- contribucion_mensual %>%
    filter(ANIO==2018) %>%
    group_by(ANIO) %>%
    summarise(produccionNom_total=mean(produccionNom_mensual),
              produccion_total=mean(produccion_mensual),
              ventasNom_total= mean(ventasNom_mensual),
              ventas_total= mean(ventas_mensual),
              personas_total=mean(personas_mensual),
              empleados_total=mean(empleados_mensual),
              operarios_total=mean(operarios_mensual))
  
  contribucion_base<-subset(contribucion_base, select = -ANIO)
  
  contribucion_mensual<-contribucion_mensual %>%
    mutate(produccionNom_total=contribucion_base$produccionNom_total,
           produccion_total=contribucion_base$produccion_total,
           ventasNom_total=contribucion_base$ventasNom_total,
           ventas_total=contribucion_base$ventas_total,
           personas_total=contribucion_base$personas_total,
           empleados_total=contribucion_base$empleados_total,
           operarios_total=contribucion_base$operarios_total)
  
  
  #Calculo de la variables nominales y totales
  tabla1_1<-contribucion_mensual %>%
    mutate(produccionNom =(produccionNom_mensual/produccionNom_total)*100,
           produccion    =(produccion_mensual/produccion_total)*100,
           ventasNom     =(ventasNom_mensual/ventasNom_total)*100,
           ventas        =(ventas_mensual/ventas_total)*100,
           personas      =(personas_mensual/personas_total)*100,
           empleados     =(empleados_mensual/empleados_total)*100,
           operarios     =(operarios_mensual/operarios_total)*100) %>%
    select(ANIO,MES,produccionNom,produccion,ventasNom,ventas,personas,
           empleados,operarios)
  tabla1_1["DOMINIO_39_ENLACE"]<-"T_IND"
  tabla1_1["DOMINIO39_DESCRIP_ENLACE"]<-"Total Industria"
  tabla1a$DOMINIO_39_ENLACE<-as.character(tabla1a$DOMINIO_39_ENLACE)
  tabla1b<-rbind(tabla1a,tabla1_1)
  tabla1_7<-tabla1b %>% arrange(DOMINIO_39_ENLACE,ANIO,MES)
  
  #Creacion de una matriz con 12 filas y el mismo numero de columnas del archivo
  # enlaces del 2001
  N<-as.data.frame(matrix(NA,ncol = ncol(indices_01),nrow = 12),colnames=FALSE)
  colnames(N)<-colnames(indices_01)
  
  #Se une la matriz N con los índices, cada 216 datos
  in_prue<-NULL
  for(i in seq(from=216, to = nrow(indices_01), by =216)){
    if(i==216){
      in_prue<-rbind(indices_01[1:i,],N)
    }
    else{
      in_prue<-rbind(in_prue,indices_01[(i-215):i,],N)
    }
    
  }
  
  #Se acomoda de nuevo la base seleccionando algunas variables de in_prueba,
  # con otras seleccionadas de la base N
  indices_lag<-rbind(in_prue[13:nrow(in_prue),5:ncol(in_prue)],N[,5:ncol(N)])
  colnames(indices_lag) <- paste0("R", colnames(indices_lag))
  indices_lag<-cbind(in_prue,indices_lag)
  indices_lag<-indices_lag[-c((nrow(indices_lag)-11):nrow(indices_lag)),]
  
  #Creacion de las variables de la tabla anteriormente construida
  variacion<-within(indices_lag,{
    VPERSONALDEPRODUCCION  =PERSONALDEPRODUCCION/RPERSONALDEPRODUCCION
    VPERSONALDEADMINISTRACION=PERSONALDEADMINISTRACION/RPERSONALDEADMINISTRACION
    VEMPLEOTOTAL       =EMPLEOTOTAL/REMPLEOTOTAL
    VVENTASREALES      =VENTASREALES/RVENTASREALES
    VVENTASNOMINALES   =VENTASNOMINALES/RVENTASNOMINALES
    VPRODUCCIONREAL    =PRODUCCIONREAL/RPRODUCCIONREAL
    VPRODUCCIONNOMINAL =PRODUCCIONNOMINAL/RPRODUCCIONNOMINAL
  })
  
  variacion <- variacion[!is.na(variacion$DOMINIOS),]
  
  #Se crea la base de los indices con el anio base, previamente calculada
  tabla2<-tabla1_7 %>%
    filter(ANIO ==2018)
  #tabla2<-as.data.frame(tabla2[,1:ncol(indices_01)])
  tabla2<-tabla2 %>%
    select(DOMINIO_39_ENLACE,ANIO,MES,
           DOMINIO39_DESCRIP_ENLACE,
           produccionNom,produccion,ventasNom,
           ventas,personas,empleados,operarios) %>%
    as.data.frame()
  colnames(tabla2)<-paste0("I", colnames(variacion[,1:ncol(tabla2)]))
  tabla2 <- select(tabla2, -"ICLASESINDUSTRIALES" )
  varia<-variacion %>% left_join(tabla2,by=c("DOMINIOS"="IDOMINIOS",
                                             "ANO"="IANO","MES"="IMES"))
  
  varia<-varia %>%
    arrange(DOMINIOS,desc(ANO),desc(MES))
  
  #Calculo de los enlaces
  vector<-unique(varia$DOMINIOS)
  variac<-NULL
  for (k in vector) {
    vari<-varia %>%
      filter(DOMINIOS==k)
    for (j in 26:ncol(varia)) {
      for(i in 1:204){
        vari[(i+12),j]=(vari[(i+12),(j-7)]*vari[i,j])
      }
    }
    variac<-rbind(variac,vari)
  }
  
  
  cols_i <- grep("^I", names(variac), value = TRUE)
  variac<- select(variac,c(DOMINIOS,ANO,MES,CLASESINDUSTRIALES,cols_i))
  colnames(variac)<-gsub("I","",names(variac))
  tabla1<-tabla1 %>% filter(ANIO!=2018)
  #tabla1<-tabla1[,1:ncol(variac)]
  colnames(tabla1)<-colnames(variac)
  tabla1<-rbind(variac,tabla1)
  tabla1<-tabla1 %>% arrange(DOMNOS,ANO,MES)
  tabla1 <- tabla1 %>% rename("DOMINIOS"="DOMNOS")
  tabla1$DOMINIOS <- ifelse(tabla1$DOMINIOS=="T_IND",paste0(0,tabla1$DOMINIOS),tabla1$DOMINIOS)
  tabla1<-tabla1 %>% arrange(DOMINIOS,ANO,MES)
  tabla1$DOMINIOS <- ifelse(tabla1$DOMINIOS=="0T_IND",substr(tabla1$DOMINIOS,2,nchar(tabla1$DOMINIOS)),tabla1$DOMINIOS)
  
  tabla1$ANO <- as.character(tabla1$ANO)
  tabla1$MES <- as.character(tabla1$MES)
  
  #Exportar
  
  #names(sheets)
  sheet <- sheets[13]
  openxlsx::writeData(wb,sheet,as.data.frame(tabla1),
                      startRow = 13, startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  addSuperSubScriptToCell_general(wb,sheet,row=7,col=1,texto = paste0("Enero 2001 - ",meses_enu[mes]," ",anio,"[p]"))
  
  
  Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+15), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Fuente. DANE - EMMET")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+16), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("p: provisionales")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+17), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("a) agrupa los dominios 1030, 1082 y 1089 ")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+18), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("b) agrupa los dominios 1300 y 1400")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+19), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("c) No incluye la mezcla de combustibles")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+20), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("d) agrupa los dominios 2020, 2023 y 2100")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+21), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("e) agrupa los dominios 1050 y 1090")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+22), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Nota: La diferencia entre el total y la suma de los dominios se debe a aproximaciones decimales.")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+23), startCol = 1,colNames=FALSE, rowNames=FALSE)
  
  Enunciado<-paste0("Nota: Se presentan cambios por actualización de información de parte de las fuentes informantes")
  openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
                      startRow = (nrow(tabla1)+24), startCol = 1,colNames=FALSE, rowNames=FALSE)
  

  
  
  filas_pares <- seq(14, nrow(tabla1)+13, by = 2)
  filas_impares <- seq(13, nrow(tabla1)+13, by = 2)
  
  addStyle(wb, sheet, style=colgr_in, rows = filas_impares, cols = 1:4, gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl_in, rows = filas_pares, cols = 1:4, gridExpand = TRUE)
  addStyle(wb, sheet, style=colgr, rows = filas_impares, cols = 5:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=colbl, rows = filas_pares, cols = 5:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcgr, rows = filas_impares, cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = filas_pares, cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultbl, rows = (nrow(tabla1)+13), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultblfc, rows = (nrow(tabla1)+13), cols = ncol(tabla1), gridExpand = TRUE)
  
  addStyle(wb, sheet, style=rowbl, rows = (nrow(tabla1)+15), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=rowblf, rows = (nrow(tabla1)+15), cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultcbl, rows = (nrow(tabla1)+16):(nrow(tabla1)+23), cols = ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultrbl, rows = (nrow(tabla1)+24), cols = 1:ncol(tabla1), gridExpand = TRUE)
  addStyle(wb, sheet, style=ultrblf, rows = (nrow(tabla1)+24), cols = ncol(tabla1), gridExpand = TRUE)
  
  
  # # 8. Desestacionalizacion -------------------------------------------------
  # 
  # CALENDAR.FN <- function(From_year,To_year){
  #   festivos <- read_xlsx(paste0(directorio,"/data/festivos/festivos.xlsx"))
  #   #read_xlsx(paste0(directorio,"/data/festivos/festivos.xlsx"))
  #   days <- c("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
  #   #days <- c("Domingo","Lunes","Martes","Miercoles","Jueves","Viernes","Sabado")
  #   calendar   <- data.frame(dates=seq(as.POSIXct(paste0(From_year,"-01-01"),tz="GMT"),as.POSIXct(paste0(To_year,"-12-31"),tz="GMT"),"days"))
  #   calendar$year  <- year(calendar$dates)
  #   calendar$month <- month(calendar$dates)
  #   calendar$day   <- day(calendar$dates)
  #   calendar$wday  <- days[wday(calendar$dates)]
  #   
  #   # Incluir los dias festivos que no se mueven cuC!ndo caen en fin de semana +
  #   # antes de 2011
  #   add_holidays <- c(paste0(From_year:To_year,"-01-01"),
  #                     paste0(From_year:To_year,"-05-01"),
  #                     paste0(From_year:To_year,"-07-20"),
  #                     paste0(From_year:To_year,"-08-07"),
  #                     paste0(From_year:To_year,"-12-08"),
  #                     paste0(From_year:To_year,"-12-25"))
  #   
  #   holidays <- c(as.POSIXct(add_holidays,format="%Y-%m-%d",tz="GMT"),as.POSIXct(festivos$FESTIVOS,format="%Y-%m-%d",tz="GMT"))
  #   holidays <- sort(holidays[!duplicated(holidays)])
  #   
  #   # Incluir domingo de ramos y domingo de resurreccion
  #   holidays <- c(holidays,holidays[diff(holidays)==1]-ddays(4),holidays[diff(holidays)==1]+ddays(3))
  #   holidays <- sort(holidays[!duplicated(holidays)])
  #   
  #   calendar$holiday <- ifelse(as.POSIXct(calendar$dates,format="%Y-%m-%d",tz="GMT")%in%holidays,1,0)
  #   
  #   calendar_group <- calendar %>% group_by(year,month, wday) %>% summarise(total=n(),holidays=sum(holiday))
  #   calendar_group$total_available_days <- calendar_group$total-calendar_group$holidays
  #   calendar_pivot <- calendar_group %>% select(c("year","month","wday","total_available_days")) %>% pivot_wider(id_cols = c("year","month"),names_from = c("wday"),values_from = c("total_available_days"))
  #   
  #   calendar_final <- calendar_pivot
  #   for(i in days[-1]){
  #     calendar_final[,paste0(i)] <- calendar_final[,i]-calendar_final$Sunday
  #   }
  #   calendar_final <- calendar_final[,c("year","month",days[-1])]
  #   return(calendar_final)
  # }
  # calendar<-as.data.frame(CALENDAR.FN(2001,anio))
  # 
  # 
  # # Produccion Real ---------------------------------------------------------
  # 
  # #Seleccion de variables de produccion real
  # produccionreal<-tabla1[,c("DOMINIOS","ANO","MES","PRODUCCONREAL")]
  # produccionreal<-produccionreal %>% filter(DOMINIOS=="T_IND")
  # 
  # produccionreal<-produccionreal %>%
  #   right_join(calendar,by=c("ANO"="year","MES"="month"))
  # produccionreal<-produccionreal[,c("PRODUCCONREAL","Monday","Tuesday","Wednesday",
  #                                   "Thursday","Friday","Saturday")]
  # 
  # #Convertir la tabla en una serie de tiempo
  # produccionreal_ts <- ts(produccionreal$PRODUCCONREAL, start = c(2001,1),frequency = 12)
  # calendar_ts<-ts(produccionreal[,2:ncol(produccionreal)],start = c(2001,1),frequency = 12)
  # colnames(calendar_ts)<-NULL
  # 
  # 
  # #Destacionalizacion
  # produccionreal_desest <- seasonal::seas(
  #   x = produccionreal_ts,
  #   arima.model = "(0 1 1)(0 1 1)",
  #   series.span = paste0(" 2001.1,",anio,".",mes," "),
  #   #" 2001.1,2022.11",
  #   #series.span = "2001.1,2022.11",
  #   #series.modelspan = "2001.1,2022.11",
  #   #transform.function = "auto",
  #   xreg = calendar_ts,
  #   regression.variables = c("easter[2]", "lpyear", "AO2016.Jul", "AO2016.Aug",
  #                            "AO2020.Mar", "TC2020.Apr", "AO2021.May"),
  #   regression.aictest ="easter",
  #   outlier.types = "all",
  #   outlier.lsrun = 3,
  #   automdl.savelog = "amd",
  #   forecast.maxlead = 1,
  #   #forecast.exclude = 12,
  #   #forecast.end = "2022.12",
  #   x11.seasonalma = "S3X5",
  #   history.estimates = c("fcst","aic","sadj","sadjchng", "trend","trendchng"),
  #   na.action = na.omit
  # )
  # 
  # #Tabla  con los resultados de la desestacionalizacion
  # tabla_prod<-as.data.frame(produccionreal_desest$data)
  # tabla_prod<-tabla_prod$final
  # deses<-tabla1[,c("DOMINIOS","ANO","MES","CLASESNDUSTRALES")]
  # deses<-deses %>% filter(DOMINIOS=="T_IND")
  # deses<-cbind(deses,tabla_prod)
  # 
  # # Ventas Reales -----------------------------------------------------------
  # 
  # #Seleccion de variables de ventas real
  # ventasreales<-tabla1[,c("DOMINIOS","ANO","MES","VENTASREALES")]
  # ventasreales<-ventasreales %>% filter(DOMINIOS=="T_IND")
  # 
  # ventasreales<-ventasreales %>%
  #   right_join(calendar,by=c("ANO"="year","MES"="month"))
  # ventasreales<-ventasreales[,c("VENTASREALES","Monday","Tuesday","Wednesday",
  #                               "Thursday","Friday","Saturday")]
  # 
  # 
  # #Convertir la tabla en una serie de tiempo
  # ventasreales_ts <- ts(ventasreales$VENTASREALES, start = c(2001,1),frequency = 12)
  # calendar_ts<-ts(ventasreales[,2:ncol(ventasreales)],start = c(2001,1),frequency = 12)
  # colnames(calendar_ts)<-NULL
  # 
  # 
  # #Destacionalizacion
  # ventasreales_desest <-seas(
  #   x = ventasreales_ts,
  #   arima.model = "(0 1 1)(0 1 1)",
  #   series.span = paste0(" 2001.1,",anio,".",mes," "),
  #   #series.span = "2001.1,2022.11",
  #   #series.modelspan = "2001.1,2022.11",
  #   transform.function = "auto",
  #   xreg = calendar_ts,
  #   regression.variables = c("easter[2]", "lpyear", "AO2016.Jul", "AO2016.Aug",
  #                            "AO2020.Mar", "TC2020.Apr", "AO2021.May"),
  #   regression.aictest ="easter",
  #   outlier.types = "all",
  #   outlier.lsrun = 3,
  #   automdl.savelog = "amd",
  #   forecast.maxlead = 1,
  #   #forecast.exclude = 12,
  #   #forecast.end = "2022.12",
  #   x11.seasonalma = "S3X5",
  #   history.estimates = c("fcst","aic","sadj","sadjchng", "trend","trendchng"),
  #   na.action = na.omit
  # )
  # 
  # #Tabla  con los resultados de la desestacionalizacion
  # tabla_vent<-as.data.frame(ventasreales_desest$data)
  # tabla_vent<-tabla_vent$final
  # deses<-cbind(deses,tabla_vent)
  # 
  # 
  # 
  # # Empleo Total ------------------------------------------------------------
  # 
  # #Seleccion de variables de empleo
  # empleototal<-tabla1[,c("DOMINIOS","ANO","MES","EMPLEOTOTAL")]
  # empleototal<-empleototal %>% filter(DOMINIOS=="T_IND")
  # 
  # empleototal<-empleototal %>%
  #   right_join(calendar,by=c("ANO"="year","MES"="month"))
  # empleototal<-empleototal[,c("EMPLEOTOTAL","Monday","Tuesday","Wednesday",
  #                             "Thursday","Friday","Saturday")]
  # 
  # #Convertir la tabla en una serie de tiempo
  # empleototal_ts <- ts(empleototal$EMPLEOTOTAL, start = c(2001,1),frequency = 12)
  # calendar_ts<-ts(empleototal[,2:ncol(empleototal)],start = c(2001,1),frequency = 12)
  # colnames(calendar_ts)<-NULL
  # 
  # 
  # #Destacionalizacion
  # empleototal_desest <-seas(
  #   x = empleototal_ts,
  #   arima.model = "(0 1 1)(0 1 1)",
  #   series.span = paste0(" 2001.1,",anio,".",mes," "),
  #   transform.function = "auto",
  #   xreg = calendar_ts,
  #   regression.variables = c("easter[2]", "lpyear", "AO2016.Jul", "AO2016.Aug",
  #                            "AO2020.Mar", "TC2020.Apr", "AO2021.May"),
  #   regression.aictest ="easter",
  #   outlier.types = "all",
  #   outlier.lsrun = 3,
  #   automdl.savelog = "amd",
  #   forecast.maxlead = 1,
  #   #forecast.exclude = 12,
  #   #forecast.end = "2022.12",
  #   x11.seasonalma = "S3X5",
  #   history.estimates = c("fcst","aic","sadj","sadjchng", "trend","trendchng"),
  #   na.action = na.omit
  # )
  # 
  # #Tabla  con los resultados de la desestacionalizacion
  # tabla_empl<-as.data.frame(empleototal_desest$data)
  # tabla_empl<-tabla_empl$final
  # deses<-cbind(deses,tabla_empl)
  # 
  # 
  # #Exportar
  # 
  # #names(sheets)
  # sheet <- sheets[11]
  # openxlsx::writeData(wb,sheet,as.data.frame(deses),
  #                     startRow = 14, startCol = 1,colNames=FALSE, rowNames=FALSE)
  # 
  # Enunciado<-paste0("Enero 2001 - ",meses_enu[mes]," ",anio,"p")
  # openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
  #                     startRow = 8, startCol = 1,colNames=FALSE, rowNames=FALSE)
  # 
  # Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  # openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
  #                     startRow = (nrow(deses)+16), startCol = 1,colNames=FALSE, rowNames=FALSE)
  # 
  # Enunciado<-paste0("Fuente. DANE - EMMET")
  # openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
  #                     startRow = (nrow(deses)+17), startCol = 1,colNames=FALSE, rowNames=FALSE)
  # 
  # Enunciado<-paste0("p: provisionales")
  # openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
  #                     startRow = (nrow(deses)+18), startCol = 1,colNames=FALSE, rowNames=FALSE)
  # 
  # Enunciado<-paste0("Programa utilizado: X13 ARIMA US Census Bureau")
  # openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
  #                     startRow = (nrow(deses)+19), startCol = 1,colNames=FALSE, rowNames=FALSE)
  # 
  # filas_pares <- seq(14, nrow(deses)+14, by = 2)
  # filas_impares <- seq(15, nrow(deses)+14, by = 2)
  # addStyle(wb, sheet, style=colgr, rows = filas_pares, cols = 1:ncol(deses), gridExpand = TRUE)
  # addStyle(wb, sheet, style=colbl, rows = filas_impares, cols = 1:ncol(deses), gridExpand = TRUE)
  # addStyle(wb, sheet, style=ultbl, rows = (nrow(deses)+14), cols = 1:ncol(deses), gridExpand = TRUE)
  # addStyle(wb, sheet, style=ultcgr, rows = filas_pares, cols = ncol(deses), gridExpand = TRUE)
  # addStyle(wb, sheet, style=ultcbl, rows = filas_impares, cols = ncol(deses), gridExpand = TRUE)
  # 
  # 
  # addStyle(wb, sheet, style=rowbl, rows = (nrow(deses)+16), cols = 1:ncol(deses), gridExpand = TRUE)
  # addStyle(wb, sheet, style=ultcbl, rows = (nrow(deses)+16):(nrow(deses)+19), cols = ncol(deses), gridExpand = TRUE)
  # addStyle(wb, sheet, style=ultrbl, rows = (nrow(deses)+19), cols = 1:ncol(deses), gridExpand = TRUE)
  # 
  # 
  # 
  # 
  # # 11. Var y cont_Trienal --------------------------------------------------
  # 
  # 
  # #Calculo de la contribucion total
  # contribucion_total <- data %>%
  #   filter(MES==mes & ANIO%in%c(2019)) %>%
  #   summarise(produccion_total = sum(PRODUCCIONREALPOND),
  #             ventas_total=sum(VENTASREALESPOND),
  #             personal_total=sum(TOTPERS))
  # 
  # #contribucion_total <- contr_sum_an(contribucion_total)
  # 
  # #Calculo de la contribucion mensual
  # contribucion <- data %>%
  #   filter(MES==mes & ANIO%in%c(anio,2019)) %>%
  #   #mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
  #   group_by(ANIO,MES,DOMINIO_39,DOMINIO39_DESCRIP) %>%
  #   summarise(prod = sum(PRODUCCIONREALPOND),
  #             vent=sum(VENTASREALESPOND),
  #             per=sum(TOTPERS)) %>%
  #   group_by(DOMINIO_39,DOMINIO39_DESCRIP) %>%
  #   summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
  #             ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
  #             personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
  #   arrange(produccion)
  # 
  # #contribucion <- contr_fin(11,contribucion)
  # 
  # #Calculo de la variacion
  # tabla1 <- data %>%
  #   filter(ANIO%in%c(anio,2019) & MES%in%mes) %>%
  #   group_by(ANIO,DOMINIO_39,DOMINIO39_DESCRIP) %>%
  #   summarise(produccionNom=sum(PRODUCCIONNOMPOND),
  #             produccion=sum(PRODUCCIONREALPOND),
  #             ventasNom = sum(VENTASNOMINPOND),
  #             ventas = sum(VENTASREALESPOND),
  #             #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)
  #             personas=sum(TOTPERS))
  # 
  # #tabla1 <- tabla_summarise(11,tabla1)
  # 
  # tabla1 <- tabla1 %>%
  #   pivot_wider(names_from = c("ANIO"),
  #               values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
  # 
  # #tabla1 <- tabla_pivot(11,tabla1)
  # 
  # #tabla1 <- tabla_paste_an(11,tabla1)
  # 
  # tabla1[paste0("varprodnom_",anio)] <- (tabla1[paste0("produccionNom_",anio)]-tabla1[paste0("produccionNom_",2019)])/tabla1[paste0("produccionNom_",2019)]
  # tabla1[paste0("varprod_",anio)] <- (tabla1[paste0("produccion_",anio)]-tabla1[paste0("produccion_",2019)])/tabla1[paste0("produccion_",2019)]
  # tabla1[paste0("varventasnom_",anio)]<- (tabla1[paste0("ventasNom_",anio)]-tabla1[paste0("ventasNom_",2019)])/tabla1[paste0("ventasNom_",2019)]
  # tabla1[paste0("varventas_",anio)]<- (tabla1[paste0("ventas_",anio)]-tabla1[paste0("ventas_",2019)])/tabla1[paste0("ventas_",2019)]
  # tabla1[paste0("varpersonas_",anio)] <- (tabla1[paste0("personas_",anio)]-tabla1[paste0("personas_",2019)])/tabla1[paste0("personas_",2019)]
  # 
  # 
  # 
  # 
  # tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:5)],names_to = "variables",values_to = "value" )
  # 
  # #tabla1 <- tabla_acople(tabla1)
  # tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
  # tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
  # tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
  # tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
  # 
  # 
  # #Empalme de la variacion y contribucion
  # tabla1 <- tabla1 %>%
  #   left_join(contribucion,by=c("DOMINIO_39"="DOMINIO_39",
  #                               "DOMINIO39_DESCRIP"="DOMINIO39_DESCRIP"))
  # tabla1 <- tabla1[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprodnom",
  #                     "varprod","produccion","varventasnom","varventas","ventas",
  #                     "varpersonas","personal")]
  # 
  # for( i in c("varprodnom",
  #             "varprod","produccion","varventasnom","varventas","ventas",
  #             "varpersonas","personal")){
  #   tabla1[,i] <-  tabla1[,i]*100
  # }
  # 
  # #Total Industria
  # 
  # contribucion1 <- data %>%
  #   filter(MES==mes & ANIO%in%c(anio,2019)) %>%
  #   #mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
  #   group_by(ANIO,MES) %>%
  #   summarise(prod = sum(PRODUCCIONREALPOND),
  #             vent=sum(VENTASREALESPOND),
  #             per=sum(TOTPERS))
  # contribucion1["DOMINIO_39"] <- "Total Industria"
  # contribucion1["DOMINIO39_DESCRIP"] <- ""
  # contribucion1 <- contribucion1 %>%
  #   group_by(DOMINIO_39,DOMINIO39_DESCRIP) %>%
  #   summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
  #             ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
  #             personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
  #   arrange(produccion)
  # 
  # #contribucion <- contr_fin(11,contribucion)
  # 
  # #Calculo de la variacion
  # tabla2 <- data %>%
  #   filter(ANIO%in%c(anio,2019) & MES%in%mes) %>%
  #   group_by(ANIO) %>%
  #   summarise(produccionNom=sum(PRODUCCIONNOMPOND),
  #             produccion=sum(PRODUCCIONREALPOND),
  #             ventasNom = sum(VENTASNOMINPOND),
  #             ventas = sum(VENTASREALESPOND),
  #             #personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)
  #             personas=sum(TOTPERS))
  # 
  # #tabla2 <- tabla_summarise(11,tabla2)
  # 
  # tabla2 <- tabla2 %>%
  #   pivot_wider(names_from = c("ANIO"),
  #               values_from = c("produccionNom","produccion","ventasNom","ventas","personas"))
  # 
  # #tabla2 <- tabla_pivot(11,tabla2)
  # 
  # #tabla2 <- tabla_paste_an(11,tabla2)
  # 
  # tabla2[paste0("varprodnom_",anio)] <- (tabla2[paste0("produccionNom_",anio)]-tabla2[paste0("produccionNom_",2019)])/tabla2[paste0("produccionNom_",2019)]
  # tabla2[paste0("varprod_",anio)] <- (tabla2[paste0("produccion_",anio)]-tabla2[paste0("produccion_",2019)])/tabla2[paste0("produccion_",2019)]
  # tabla2[paste0("varventasnom_",anio)]<- (tabla2[paste0("ventasNom_",anio)]-tabla2[paste0("ventasNom_",2019)])/tabla2[paste0("ventasNom_",2019)]
  # tabla2[paste0("varventas_",anio)]<- (tabla2[paste0("ventas_",anio)]-tabla2[paste0("ventas_",2019)])/tabla2[paste0("ventas_",2019)]
  # tabla2[paste0("varpersonas_",anio)] <- (tabla2[paste0("personas_",anio)]-tabla2[paste0("personas_",2019)])/tabla2[paste0("personas_",2019)]
  # 
  # 
  # 
  # 
  # tabla2 <- tabla2 %>% pivot_longer(cols = colnames(tabla2)[-c(1:5)],names_to = "variables",values_to = "value" )
  # 
  # #tabla2 <- tabla_acople(tabla2)
  # tabla2$ANIO <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 2)
  # tabla2$variables <- sapply(strsplit(as.character(tabla2$variables), "_"), `[`, 1)
  # tabla2 <- tabla2 %>% filter(gsub("var","",variables)!=variables )
  # tabla2 <- tabla2 %>% pivot_wider(names_from = variables,values_from = value)
  # 
  # tabla2["DOMINIO_39"] <- "Total Industria"
  # tabla2["DOMINIO39_DESCRIP"] <- ""
  # 
  # #Empalme de la variacion y contribucion
  # tabla2 <- tabla2 %>%
  #   left_join(contribucion1,by=c("DOMINIO_39"="DOMINIO_39",
  #                                "DOMINIO39_DESCRIP"="DOMINIO39_DESCRIP"))
  # tabla2 <- tabla2[,c("DOMINIO_39","DOMINIO39_DESCRIP","varprodnom",
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
  # tabla1 <- tabla1 %>% arrange(DOMINIO_39)
  # tabla1 <- rbind(tabla2,tabla1)
  # 
  # #Exportar
  # names(sheets)
  # sheet <- sheets[14]
  # openxlsx::writeData(wb,sheet,as.data.frame(tabla1),
  #                     startRow = 13, startCol = 1,colNames=FALSE, rowNames=FALSE)
  # 
  # Enunciado<-paste0(meses_enu[mes],"(",anio,"/","2019)p")
  # openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
  #                     startRow = 9, startCol = 1,colNames=FALSE, rowNames=FALSE)
  # 
  # Enunciado<-paste0("Fecha de publicación ",meses_enu[mes]," de ",anio)
  # openxlsx::writeData(wb,sheet,as.data.frame(Enunciado),
  #                     startRow = 57, startCol = 1,colNames=FALSE, rowNames=FALSE)
  # 
  # Guardar archivo de salida -----------------------------------------------
  
  openxlsx::saveWorkbook(wb,Salida,overwrite = TRUE)

}
