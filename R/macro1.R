#' Macro 1
#'
#' @description
#'
#' Con esta función realizaremos la primera parte del proceso, en donde se correran las 4
#' primeras funciones.
#'
#'
#' @param mes Definir el mes a ejecutar, ej: 11
#' @param anio Definir el año a ejecutar, ej: 2022
#' @param directorio definir el directorio donde se encuentran ubicado los datos de entrada
#' @param avance Denifir que porcentaje de la base va a ser procesada para detectar alertas, por defecto el valor
#' esta en 100, lo cual indica que estan los más de 3000 establecimientos para el procesamiento, este genera un
#' archivo llamado por ejemplo: "EMMET_PANEL_alertas_nov2022.csv", si el avance es diferente de 100 el nombre será
#' "EMMET_PANEL_alertas_nov2022_50.csv" indicando que es el 50\% de la base
#'
#' @details  Esta funcion tiene como objetivo correr desde la función inicial hasta la función
#' de identificacion de alertas, con el fin de que el usuario pueda obtener rapidamente el
#' archivo de salida de la función de alertas y poder realizar el proceso de critica.
#' Recuerde que si el archivo de alertas ya está creado, al correr esta función le pedira que
#' responda S si quiere sobreescribirlo, de lo contrario otro valor para cancelar la función
#'
#' Ver:\code{\link{f0_inicial}},\code{\link{f1_integracion}},
#' \code{\link{f2_identificacion_alertas}}
#'
#'
#' @examples
#'  macro1(directorio="Documents/DANE/Procesos DIMPE /PilotoEMMET",
#'                        mes=11,anio=2022,avance=100)


macro1<-function(directorio,mes,anio,avance=100){
f0_inicial(directorio)
  print("Se ejecuto la funcion inicial")
f1_integracion(directorio,mes,anio)
print("Se ejecuto la funcion integracion")
f2_identificacion_alertas(directorio,mes,anio,avance)
print(paste0("Se ejecuto la funcion identificacion outliers, por favor verifique el archivo en la
      carpeta",directorio,"results/S2_identificacion_alertas y realice el proceso de critica."))
print("Recuerde seguir las instrucciones para guardar el archivo y continue con el
      proceso llamando la funcion macro2(directorio,mes,anio) ")
}










