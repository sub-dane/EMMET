##Script DANILO

directorio="C:/Users/USUARIO/OneDrive - dane.gov.co/proyecto2/pruebas Danilo"
mes=10
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


base_panel2=read.xlsx(paste0(directorio,"/data/",anio,"/",meses[mes],"/COMPLETO PRIMER CIERRE OCTUBRE 2024.xlsx"),sheet = "COMPLETO")
colnames(base_panel2)   <- colnames_format(base_panel2)
names(base_panel2)[names(base_panel2) %in% variablesinte2] <- variablesinte
names(base_panel2)[names(base_panel2) %in%  c("ID_NUMORD","NOMBRE_ESTAB","NOMBREDPTO","DESCRIPCIONDOMINIOEMMET39","II_TOT_TOT_PERS","AJU_III_TOT_TOTAL_VENTAS","TOTALHORAS")] <- c("NORDEST","NOMBRE_ESTABLECIMIENTO","DEPARTAMENTO","DOMINIO39_DESCRIP","TOTPERS","TOTAL_VENTAS","TOTAL_HORAS")


#ir a la función f4_tematica y correr desde la linea 132

f5_anacional(directorio,mes,anio)
f6_aterritorial(directorio,mes,anio)
f7_cdominios(directorio,mes,anio)
f8_cregiones(directorio,mes,anio)

