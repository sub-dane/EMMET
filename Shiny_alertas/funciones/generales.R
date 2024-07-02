
# Librerias ---------------------------------------------------------------

require(plotly)
require(readxl)
require(data.table)

# valores -----------------------------------------------------------------


meses_c <- c("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")

meses <- c("ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic")
# Funciones  --------------------------------------------------------------

colnames_format <- function(base){
  colnames(base) <- toupper(colnames(base))
  colnames(base) <- gsub(" ","",colnames(base))
  colnames(base) <- gsub("__","_",colnames(base))
  colnames(base) <- stringi::stri_trans_general(colnames(base), id = "Latin-ASCII")
  return(colnames(base))
}


# Bases -------------------------------------------------------------------

base_panel<- read.csv(paste0(directorio,"/results/S2_estandarizacion/EMMET_PANEL_estandarizado",meses[mes],anio,".csv"),fileEncoding = "latin1")

alertas <- fread(paste0(directorio,"/results/S3_identificacion_alertas/EMMET_PANEL_alertas_",meses[mes],anio,".csv"))

base_panel<- as.data.frame(base_panel)

tematica<-fread(paste0(directorio,"/results/S5_tematica/EMMET_PANEL_tematica_",meses[mes],anio,".csv"))

# Tratamiento bases -------------------------------------------------------


colnames(tematica) <- colnames_format(tematica)
tematica<-  tematica%>% mutate_at(vars(contains("OBSE")),~str_replace_all(.,pattern="[^[:alnum:]]",replacement=" "))
colnames(base_panel)<-colnames_format(base_panel)

