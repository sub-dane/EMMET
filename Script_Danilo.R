##Script DANILO

directorio="C:/Users/USUARIO/OneDrive - dane.gov.co/proyecto2/pruebas Danilo"
mes=9
anio=2024  


library(devtools)
devtools::install_github("sub-dane/EMMET",force = TRUE)
library(DANE.EMMET)
f0_inicial(directorio,mes,anio)

variablesinte2=c("II_PA_PP_NPERS_EP","AJU_II_PA_PP_SUELD_EP","II_PA_TD_NPERS_ET",
                 "AJU_II_PA_TD_SUELD_ET","II_PA_TI_NPERS_ETA","AJU_II_PA_TI_SUELD_ETA",
                 "II_PA_AP_AAEP","AJU_II_PA_AP_AAS_AP","II_PP_PP_NPERS_OP","AJU_II_PP_PP_SUELD_OP",
                 "II_PP_TD_NPERS_OT","AJU_II_PP_TD_SUELD_OT","II_PP_TI_NPERS_OTA","AJU_II_PP_TI_SUELD_OTA",
                 "II_PP_AP_APEP","AJU_II_PP_AP_AAS_PP","AJU_II_HORAS_HORDI_T","AJU_II_HORAS_HEXTR_T",
                 "AJU_III_PE_PRODUCCION","AJU_III_PE_VENTASIN","AJU_III_PE_VENTASEX","III_EX_VEXIS")


base_panel2=read.xlsx("C:/Users/USUARIO/OneDrive - dane.gov.co/proyecto2/pruebas Danilo/data/2024/sep/COMPLETO PRIMER CIERRE SEPTIEMBRE 2024.xlsx",sheet = "COMPLETO")
colnames(base_panel2)   <- colnames_format(base_panel2)
names(base_panel2)[names(base_panel2) %in% variablesinte2] <- variablesinte
names(base_panel2)[names(base_panel2) %in%  c("ID_NUMORD","NOMBRE_ESTAB","NOMBREDPTO","DESCRIPCIONDOMINIOEMMET39","II_TOT_TOT_PERS","AJU_III_TOT_TOTAL_VENTAS","TOTALHORAS")] <- c("NORDEST","NOMBRE_ESTABLECIMIENTO","DEPARTAMENTO","DOMINIO39_DESCRIP","TOTPERS","TOTAL_VENTAS","TOTAL_HORAS")


#ir a la función f4_tematica y correr desde la linea 132

f5_anacional(directorio,mes,anio)
f6_aterritorial(directorio,mes,anio)
f7_cdominios(directorio,mes,anio)
f8_cregiones(directorio,mes,anio)




library(openxlsx)
wb <- createWorkbook()
addWorksheet(wb, sheetName = "hoja1")

 # 8. Desestacionalizacion -------------------------------------------------
 
 CALENDAR.FN <- function(From_year,To_year){
   festivos <- read_xlsx(paste0(directorio,"/data/Archivos_necesarios/festivos.xlsx"))
   #read_xlsx(paste0(directorio,"/data/festivos/festivos.xlsx"))
   days <- c("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
   #days <- c("Domingo","Lunes","Martes","Miercoles","Jueves","Viernes","Sabado")
   calendar   <- data.frame(dates=seq(as.POSIXct(paste0(From_year,"-01-01"),tz="GMT"),as.POSIXct(paste0(To_year,"-12-31"),tz="GMT"),"days"))
   calendar$year  <- year(calendar$dates)
   calendar$month <- month(calendar$dates)
   calendar$day   <- day(calendar$dates)
   calendar$wday  <- days[wday(calendar$dates)]
   
   # Incluir los dias festivos que no se mueven cuC!ndo caen en fin de semana +
   # antes de 2011
   add_holidays <- c(paste0(From_year:To_year,"-01-01"),
                     paste0(From_year:To_year,"-05-01"),
                     paste0(From_year:To_year,"-07-20"),
                     paste0(From_year:To_year,"-08-07"),
                     paste0(From_year:To_year,"-12-08"),
                     paste0(From_year:To_year,"-12-25"))
   
   holidays <- c(as.POSIXct(add_holidays,format="%Y-%m-%d",tz="GMT"),as.POSIXct(festivos$FESTIVOS,format="%Y-%m-%d",tz="GMT"))
   holidays <- sort(holidays[!duplicated(holidays)])
   
   # Incluir domingo de ramos y domingo de resurreccion
   holidays <- c(holidays,holidays[diff(holidays)==1]-ddays(4),holidays[diff(holidays)==1]+ddays(3))
   holidays <- sort(holidays[!duplicated(holidays)])
   
   calendar$holiday <- ifelse(as.POSIXct(calendar$dates,format="%Y-%m-%d",tz="GMT")%in%holidays,1,0)
   
   calendar_group <- calendar %>% group_by(year,month, wday) %>% summarise(total=n(),holidays=sum(holiday))
   calendar_group$total_available_days <- calendar_group$total-calendar_group$holidays
   calendar_pivot <- calendar_group %>% select(c("year","month","wday","total_available_days")) %>% pivot_wider(id_cols = c("year","month"),names_from = c("wday"),values_from = c("total_available_days"))
   
   calendar_final <- calendar_pivot
   for(i in days[-1]){
     calendar_final[,paste0(i)] <- calendar_final[,i]-calendar_final$Sunday
   }
   calendar_final <- calendar_final[,c("year","month",days[-1])]
   return(calendar_final)
 }
 calendar<-as.data.frame(CALENDAR.FN(2001,anio))
 
 
 # Produccion Real ---------------------------------------------------------
 
 #Seleccion de variables de produccion real
 produccionreal<-tabla1[,c("DOMINIOS","ANO","MES","PRODUCCONREAL")]
 produccionreal<-produccionreal %>% filter(DOMINIOS=="T_IND")
 
 produccionreal<-produccionreal %>%
   right_join(calendar,by=c("ANO"="year","MES"="month"))
 produccionreal<-produccionreal[,c("PRODUCCONREAL","Monday","Tuesday","Wednesday",
                                   "Thursday","Friday","Saturday")]
 
 #Convertir la tabla en una serie de tiempo
 produccionreal_ts <- ts(produccionreal$PRODUCCONREAL, start = c(2001,1),frequency = 12)
 calendar_ts<-ts(produccionreal[,2:ncol(produccionreal)],start = c(2001,1),frequency = 12)
 colnames(calendar_ts)<-NULL
 
 
  produccionreal_desest <- seasonal::seas(
    x = produccionreal_ts,
    #arima.model = "(0 1 1)(0 1 1)",
    #series.span = paste0(" 2001.1,",anio,".",mes," "),
    #" 2001.1,2022.11",
    #series.span = "2001.1,2022.11",
    #series.modelspan = "2001.1,2022.11",
    transform.function = "auto",
    xreg = calendar_ts,
    regression.variables = c( "AO2016.Jul", "AO2016.Aug",
                             "AO2020.Mar", "TC2020.Apr", "AO2021.May"),
    regression.aictest =c("easter","lpyear"),
    force.type="denton",
    #outlier.types = "all",
    #outlier.lsrun = 3,
    #automdl.savelog = "amd",
    forecast.maxlead = 1,
    #forecast.exclude = 12,
    #forecast.end = "2022.12",
    x11= "",
    #history.estimates = c("fcst","aic","sadj","sadjchng", "trend","trendchng"),
    na.action = na.omit
  )
 
 #Tabla  con los resultados de la desestacionalizacion
 tabla_prod<-as.data.frame(produccionreal_desest$data)
 tabla_prod<-tabla_prod$final
 deses<-tabla1[,c("DOMINIOS","ANO","MES","CLASESNDUSTRALES")]
 deses<-deses %>% filter(DOMINIOS=="T_IND")
 deses<-cbind(deses,tabla_prod)
 
 # Ventas Reales -----------------------------------------------------------
 
 #Seleccion de variables de ventas real
 ventasreales<-tabla1[,c("DOMINIOS","ANO","MES","VENTASREALES")]
 ventasreales<-ventasreales %>% filter(DOMINIOS=="T_IND")
 
 ventasreales<-ventasreales %>%
   right_join(calendar,by=c("ANO"="year","MES"="month"))
 ventasreales<-ventasreales[,c("VENTASREALES","Monday","Tuesday","Wednesday",
                               "Thursday","Friday","Saturday")]
 
 
 #Convertir la tabla en una serie de tiempo
 ventasreales_ts <- ts(ventasreales$VENTASREALES, start = c(2001,1),frequency = 12)
 calendar_ts<-ts(ventasreales[,2:ncol(ventasreales)],start = c(2001,1),frequency = 12)
 colnames(calendar_ts)<-NULL
 
 
 #Destacionalizacion
 ventasreales_desest <-seas(
   x = ventasreales_ts,
   #arima.model = "(0 1 1)(0 1 1)",
   #series.span = paste0(" 2001.1,",anio,".",mes," "),
   #" 2001.1,2022.11",
   #series.span = "2001.1,2022.11",
   #series.modelspan = "2001.1,2022.11",
   transform.function = "auto",
   xreg = calendar_ts,
   regression.variables = c( "AO2016.Jul", "AO2016.Aug",
                             "AO2020.Mar", "TC2020.Apr", "AO2021.May"),
   regression.aictest =c("easter","lpyear"),
   force.type="denton",
   #outlier.types = "all",
   #outlier.lsrun = 3,
   #automdl.savelog = "amd",
   forecast.maxlead = 1,
   #forecast.exclude = 12,
   #forecast.end = "2022.12",
   x11= "",
   #history.estimates = c("fcst","aic","sadj","sadjchng", "trend","trendchng"),
   na.action = na.omit
 )
 
 #Tabla  con los resultados de la desestacionalizacion
 tabla_vent<-as.data.frame(ventasreales_desest$data)
 tabla_vent<-tabla_vent$final
 deses<-cbind(deses,tabla_vent)
 
 
 
 # Empleo Total ------------------------------------------------------------
 
 #Seleccion de variables de empleo
 empleototal<-tabla1[,c("DOMINIOS","ANO","MES","EMPLEOTOTAL")]
 empleototal<-empleototal %>% filter(DOMINIOS=="T_IND")
 
 empleototal<-empleototal %>%
   right_join(calendar,by=c("ANO"="year","MES"="month"))
 empleototal<-empleototal[,c("EMPLEOTOTAL","Monday","Tuesday","Wednesday",
                             "Thursday","Friday","Saturday")]
 
 #Convertir la tabla en una serie de tiempo
 empleototal_ts <- ts(empleototal$EMPLEOTOTAL, start = c(2001,1),frequency = 12)
 calendar_ts<-ts(empleototal[,2:ncol(empleototal)],start = c(2001,1),frequency = 12)
 colnames(calendar_ts)<-NULL
 
 
 #Destacionalizacion
 empleototal_desest <-seas(
   x = empleototal_ts,
   #arima.model = "(0 1 1)(0 1 1)",
   #series.span = paste0(" 2001.1,",anio,".",mes," "),
   #" 2001.1,2022.11",
   #series.span = "2001.1,2022.11",
   #series.modelspan = "2001.1,2022.11",
   transform.function = "auto",
   xreg = calendar_ts,
   regression.variables = c( "AO2016.Jul", "AO2016.Aug",
                             "AO2020.Mar", "TC2020.Apr", "AO2021.May"),
   regression.aictest =c("easter","lpyear"),
   force.type="denton",
   #outlier.types = "all",
   #outlier.lsrun = 3,
   #automdl.savelog = "amd",
   forecast.maxlead = 1,
   #forecast.exclude = 12,
   #forecast.end = "2022.12",
   x11= "",
   #history.estimates = c("fcst","aic","sadj","sadjchng", "trend","trendchng"),
   na.action = na.omit
 )
 
 #Tabla  con los resultados de la desestacionalizacion
 tabla_empl<-as.data.frame(empleototal_desest$data)
 tabla_empl<-tabla_empl$final
 deses<-cbind(deses,tabla_empl)
 
 
 #Exportar
 
 wb2 <- createWorkbook()
 addWorksheet(wb2, sheetName = "hoja1")
 writeData(wb2, sheet = "hoja1", x = deses)
 saveWorkbook(wb2, paste0("Pruebas_desestacinalización",mes,anio,".xlsx"),overwrite = TRUE)
