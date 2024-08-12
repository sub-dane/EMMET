
# librerias ---------------------------------------------------------------
require(dplyr)
require(tidyverse)
require(zoo)
require(xts)
require(lubridate)
require(TSstudio)
require(tseries)
require(forecast)


# codigo ------------------------------------------------------------------



capitulo3=c("Ajuste_Producción","Ajuste_Venta_productos_elaborados_país",
            "Ajuste_Venta_productos_elaborados_exterior","Valor_existencias_precio_costo" )
variablesinte <- c("NPERS_EP","AJU_SUELD_EP","NPERS_ET","AJU_SUELD_ET","NPERS_ETA","AJU_SUELD_ETA",
                   "NPERS_APREA","AJU_SUELD_APREA","NPERS_OP","AJU_SUELD_OP","NPERS_OT","AJU_SUELD_OT",
                   "NPERS_OTA","AJU_SUELD_OTA","NPERS_APREO","AJU_SUELD_APREO","AJU_HORAS_ORDI",
                   "AJU_HORAS_EXT","AJU_PRODUCCION","AJU_VENTASIN","AJU_VENTASEX","EXISTENCIAS")


for (i in variablesinte) {

  base_panel[,i] <- as.numeric(base_panel[,i])
  base_panel[,i] <- ifelse(is.na(base_panel[,i]),0,base_panel[,i])
}
base_variables_encuesta <- base_panel %>%
  select(ANIO,MES,DEPARTAMENTO,NOMBREMPIO,
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

base_variables_encuesta=as.data.frame(base_variables_encuesta)

des <- base_variables_encuesta %>%
  filter(ANIO==anio & MES%in%c(mes-1,mes))  %>%
  pivot_longer(cols=colnames(base_variables_encuesta)[9:length(colnames(base_variables_encuesta))],names_to = "Variables",values_to ="Valores") %>%
  pivot_wider(names_from = "MES",values_from = "Valores")
des=as.data.frame(des)
colnames(des)[c(9,10)] <- c(meses_c[mes-1],meses_c[mes])
des$variacion <- round(((abs(des[,paste0(meses_c[mes])]-des[,paste0(meses_c[mes-1])])/des[,paste0(meses_c[mes-1])])*100),2)
des$variacion[des[, paste0(meses_c[mes])] != 0 & des[, paste0(meses_c[mes - 1])] == 0] <- 100
des$variacion[des[, paste0(meses_c[mes])] == 0 & des[, paste0(meses_c[mes - 1])] == 0] <- 0


varinteres=paste0(variablesinte,"_caso_de_imputacion")

des2 <- alertas %>% filter(ANIO==anio & MES==mes) %>%
  pivot_longer(cols=varinteres,names_to = "Variables",values_to ="Caso_imputacion")  %>%
  select(ANIO,DEPARTAMENTO,NOMBREMPIO,NORDEST,NOMBRE_ESTABLECIMIENTO,DOMINIO_39,Variables,Caso_imputacion)
des2$Variables <- sub("\\_caso_de_imputacion", "", des2$Variables)
des2=as.data.frame(des2)
df_merge <- merge(des, des2, by = c("ANIO","DEPARTAMENTO","NOMBREMPIO",
                                    "NORDEST","NOMBRE_ESTABLECIMIENTO","DOMINIO_39" ,"Variables"))

# Seleccionar todas las columnas de la tabla B y la columna valores 2 de la tabla A
desf <- df_merge[, c(names(des), "Caso_imputacion")]
desf[, paste0(meses_c[mes - 1])]=format(round(desf[, paste0(meses_c[mes - 1])], 1), big.mark=".",decimal.mark = ",")
desf[, paste0(meses_c[mes])]=format(round(desf[, paste0(meses_c[mes])], 1), big.mark=".",decimal.mark = ",")
desf$variacion=format(round(desf$variacion, 1), big.mark=".",decimal.mark = ",")
desf=arrange(desf,NORDEST)
# Renombrar variables -----------------------------------------------------


cambio=c("Empleados_permanentes","Ajuste_sueldos_Empleados_Permanentes",
         "Empleados_temporales", "Ajuste_sueldos_Empleados_Temporales",
         "Empleados_temporales_agencias","Ajuste_sueldos_Empleados_Temporales_Agencia",
         "Aprendices_pasantes","Ajuste_apoyo_sostenimiento_Aprendices_Pasantes",
         "Obreros_permanentes","Ajuste_sueldos_Operarios_Permanentes",
         "Obreros_temporales","Ajuste_sueldos_Operarios_Temporales",
         "Obreros_temporales_agencias","Ajuste_sueldos_Operarios_Temporales_Agencia",
         "Aprendices_pasantes_producción","Ajuste_apoyo_sostenimiento_aprendices_pasantes_producción",
         "Ajuste_Horas_ordinarias_trabajadas","Ajuste_Horas_extras_trabajadas",
         "Ajuste_Producción","Ajuste_Venta_productos_elaborados_país",
         "Ajuste_Venta_productos_elaborados_exterior","Valor_existencias_precio_costo")
desf$Variables <- factor(desf$Variables, levels= variablesinte, labels = cambio)
colnames(base_variables_encuesta)[9:length(colnames(base_variables_encuesta))]=c("Empleados_permanentes","Ajuste_sueldos_Empleados_Permanentes",
                                                                                 "Empleados_temporales", "Ajuste_sueldos_Empleados_Temporales",
                                                                                 "Empleados_temporales_agencias","Ajuste_sueldos_Empleados_Temporales_Agencia",
                                                                                 "Aprendices_pasantes","Ajuste_apoyo_sostenimiento_Aprendices_Pasantes",
                                                                                 "Obreros_permanentes","Ajuste_sueldos_Operarios_Permanentes",
                                                                                 "Obreros_temporales","Ajuste_sueldos_Operarios_Temporales",
                                                                                 "Obreros_temporales_agencias","Ajuste_sueldos_Operarios_Temporales_Agencia",
                                                                                 "Aprendices_pasantes_producción","Ajuste_apoyo_sostenimiento_aprendices_pasantes_producción",
                                                                                 "Ajuste_Horas_ordinarias_trabajadas","Ajuste_Horas_extras_trabajadas",
                                                                                 "Ajuste_Producción","Ajuste_Venta_productos_elaborados_país",
                                                                                 "Ajuste_Venta_productos_elaborados_exterior","Valor_existencias_precio_costo")
desf["MES"] <- mes
desf["LLAVE"] <- as.numeric(paste0(desf$ANIO,desf$MES,desf$NORDEST))
desf$OBSERVACIONES <- NA
desf$EDITADO <- NA
