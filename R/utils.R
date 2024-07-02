#utils


#crea un vector con los meses
meses <- c("ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic")

# Función formato nombres columnas
colnames_format <- function(base){
  colnames(base) <- toupper(colnames(base))
  colnames(base) <- gsub(" ","",colnames(base))
  colnames(base) <- gsub("__","_",colnames(base))
  colnames(base) <- stringi::stri_trans_general(colnames(base), id = "Latin-ASCII")
  return(colnames(base))
}

#crea un vector con las variables fecha
cols_to_date <- c("II_PA_PP_IEP","II_PA_PP_FEP","II_PA_TD_IET","II_PA_TD_FET",
                  "II_PA_TI_IETA","II_PA_TI_FETA","II_PA_AP_AI_AP","II_PA_AP_AF_AP",
                  "II_PP_PP_IOP","II_PP_PP_FOP","II_PP_TD_IOT","II_PP_TD_FOT",
                  "II_PP_TI_IOTA","II_PP_TI_FOTA","II_PP_AP_AI_PP","II_PP_AP_AF_PP",
                  "II_HORAS_HORDI_D","II_HORAS_HORDI_H",
                  "III_PE_IV","III_PE_FV")


#crea un vector con las variables de interes


variablesinte <- c("NPERS_EP","AJU_SUELD_EP","NPERS_ET","AJU_SUELD_ET","NPERS_ETA","AJU_SUELD_ETA",
"NPERS_APREA","AJU_SUELD_APREA","NPERS_OP","AJU_SUELD_OP","NPERS_OT","AJU_SUELD_OT",
"NPERS_OTA","AJU_SUELD_OTA","NPERS_APREO","AJU_SUELD_APREO","AJU_HORAS_ORDI",
"AJU_HORAS_EXT","AJU_PRODUCCION","AJU_VENTASIN","AJU_VENTASEX","EXISTENCIAS")

#crea un vector con las variables del capitulo 2
cap2=c("NPERS_EP","AJU_SUELD_EP","NPERS_ET","AJU_SUELD_ET","NPERS_ETA","AJU_SUELD_ETA",
       "NPERS_APREA","AJU_SUELD_APREA","NPERS_OP","AJU_SUELD_OP","NPERS_OT","AJU_SUELD_OT",
       "NPERS_OTA","AJU_SUELD_OTA","NPERS_APREO","AJU_SUELD_APREO","AJU_HORAS_ORDI",
       "AJU_HORAS_EXT")

#crea un vector con las variables del capitulo 3
cap3=c("AJU_PRODUCCION","AJU_VENTASIN","AJU_VENTASEX","EXISTENCIAS")

#crea un vector con las variables relacionadas a personas
personas=c("NPERS_EP","NPERS_ET","NPERS_ETA",
           "NPERS_APREA","NPERS_OP","NPERS_OT",
           "NPERS_OTA","NPERS_APREO")

#crea un vector con las variables relacionadas a sueldos
sueldos=c("AJU_SUELD_EP","AJU_SUELD_ET","AJU_SUELD_ETA",
          "AJU_SUELD_APREA","AJU_SUELD_OP","AJU_SUELD_OT",
          "AJU_SUELD_OTA","AJU_SUELD_APREO")
