
# Librerias ---------------------------------------------------------------

require(ggplot2)
require(dplyr)
require(plotly)
require(readxl)

# Función general ---------------------------------------------------------

#Funcion que realiza el calculo de la variacion dependiendo el periodo asigndo
#La variacion se calcula por todo el periodo de la base de datos es decir, 
# desde 2019, hasta el anio y mes que tenga la base ingresada. 


estructura<-function(periodo,base=dataInput(),mes=input$mes,anio=input$anio){
  #Periodo 1: Mes - Anio anterior
  if(periodo==1){
    #Se crean las variables de interés: produccion, ventas y personal
    etapa <- base  %>% 
      group_by(ANIO,MES) %>% 
      summarise(produccion=sum(PRODUCCIONREALPOND),
                ventas = sum(VENTASREALESPOND),
                personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    etapa <- etapa %>% 
      pivot_wider(names_from = c("ANIO"),values_from = c("produccion","ventas","personas"))
    #Se calculan la variaciones, sobre las variables, en el periodo de interes
    for(i in unique(base$ANIO)){
      if(i>2018){
        etapa[paste0("varprod_",i)] <- (etapa[paste0("produccion_",i)]-etapa[paste0("produccion_",i-1)])/etapa[paste0("produccion_",i-1)]
        etapa[paste0("varventas_",i)]<- (etapa[paste0("ventas_",i)]-etapa[paste0("ventas_",i-1)])/etapa[paste0("ventas_",i-1)]
        etapa[paste0("varpersonas_",i)] <- (etapa[paste0("personas_",i)]-etapa[paste0("personas_",i-1)])/etapa[paste0("personas_",i-1)]
      }}
    etapa <- etapa %>% pivot_longer(cols = colnames(etapa)[-1],names_to = "variables",values_to = "variacion" )
    
    etapa$ANIO <- sapply(strsplit(as.character(etapa$variables), "_"), `[`, 2)
    etapa$variables <- sapply(strsplit(as.character(etapa$variables), "_"), `[`, 1)
    
    
    etapa <- etapa %>% filter(gsub("var","",variables)!=variables)
    etapa$variables <- factor(etapa$variables,levels=c("varprod","varventas","varpersonas"),labels=c("Producción Real","Ventas Reales","Personal Ocupado"))
  }
  
  #Periodo 2: Año corrido
  if(periodo==2){
    #Se crean las variables de interés: produccion, ventas y personal
    etapa <- base  %>% group_by(ANIO,MES) %>% summarise(produccion=sum(PRODUCCIONREALPOND),
                                                                      ventas = sum(VENTASREALESPOND),
                                                                      personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)) %>% arrange(ANIO,MES)
    etapa <- etapa %>% group_by(ANIO) %>% mutate(suma_produccion=cumsum(produccion),
                                                                               suma_ventas   =cumsum(ventas),
                                                                               suma_personas =cumsum(personas))
    
    etapa <- etapa %>% 
      select(ANIO,MES,suma_produccion,suma_ventas,suma_personas) %>% 
      pivot_wider(names_from = c("ANIO"),values_from = c("suma_produccion","suma_ventas","suma_personas"))
    #Se calculan la variaciones, sobre las variables, en el periodo de interes
    for(i in unique(base$ANIO)){
      if(i>2018){
        etapa[paste0("varventas_",i)]<- (etapa[paste0("suma_ventas_",i)]-etapa[paste0("suma_ventas_",i-1)])/etapa[paste0("suma_ventas_",i-1)]
        etapa[paste0("varpersonas_",i)] <- (etapa[paste0("suma_personas_",i)]-etapa[paste0("suma_personas_",i-1)])/etapa[paste0("suma_personas_",i-1)]
        etapa[paste0("varprod_",i)] <- (etapa[paste0("suma_produccion_",i)]-etapa[paste0("suma_produccion_",i-1)])/etapa[paste0("suma_produccion_",i-1)]
      }}
    etapa <- etapa %>%
      pivot_longer(cols = colnames(etapa)[-1],names_to = "variables",values_to = "variacion" )
    
    etapa$ANIO <- sapply(strsplit(as.character(etapa$variables), "_"), `[`, 2)
    etapa$variables <- sapply(strsplit(as.character(etapa$variables), "_"), `[`, 1)
    
    
    etapa <- etapa %>% filter(gsub("var","",variables)!=variables)
    etapa <- etapa %>% filter(ANIO>=2019)
    etapa$variables <- factor(etapa$variables,levels=c("varpersonas","varprod","varventas"),labels=c("Personal Ocupado","Producción Real","Ventas Reales"))
    
  }
  #Periodo 3: Año acumulado
  if(periodo==3){
    #Se crean las variables de interés: produccion, ventas y personal
    etapa <- base %>% 
      group_by(ANIO,MES) %>% 
      summarise(produccion=sum(PRODUCCIONREALPOND),
                ventas = sum(VENTASREALESPOND),
                personas = sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)) %>%
      arrange(ANIO,MES)
    
    etapa$suma_produccion <- NA
    etapa$suma_ventas <- NA
    etapa$suma_personas <- NA
    for(i in 1:dim(etapa)[1]){
      if(i>=12){
        etapa$suma_produccion[i] <- sum(etapa$produccion[(i-11):i])
        etapa$suma_ventas[i] <- sum(etapa$ventas[(i-11):i])
        etapa$suma_personas[i] <- sum(etapa$personas[(i-11):i])
      }
    }
    #Se calculan la variaciones, sobre las variables, en el periodo de interes
    etapa <- etapa %>% select(ANIO,MES,suma_produccion,suma_ventas,suma_personas) %>% pivot_wider(names_from = c("ANIO"),values_from = c("suma_produccion","suma_ventas","suma_personas"))
    for(i in unique(base$ANIO)){
      if(i>2018){
        etapa[paste0("varprod_",i)] <- (etapa[paste0("suma_produccion_",i)]-etapa[paste0("suma_produccion_",i-1)])/etapa[paste0("suma_produccion_",i-1)]
        etapa[paste0("varventas_",i)]<- (etapa[paste0("suma_ventas_",i)]-etapa[paste0("suma_ventas_",i-1)])/etapa[paste0("suma_ventas_",i-1)]
        etapa[paste0("varpersonas_",i)] <- (etapa[paste0("suma_personas_",i)]-etapa[paste0("suma_personas_",i-1)])/etapa[paste0("suma_personas_",i-1)]
      }}
    etapa <- etapa %>% pivot_longer(cols = colnames(etapa)[-1],names_to = "variables",values_to = "variacion" )
    
    etapa$ANIO <- sapply(strsplit(as.character(etapa$variables), "_"), `[`, 2)
    etapa$variables <- sapply(strsplit(as.character(etapa$variables), "_"), `[`, 1)
    
    
    etapa <- etapa %>% filter(gsub("var","",variables)!=variables)
    etapa <- etapa %>% filter((ANIO>=2019 & MES==12) | ANIO>=2020)
    etapa$variables <- factor(etapa$variables,levels=c("varprod","varventas","varpersonas"),labels=c("Producción Real","Ventas Reales","Personal Ocupado"))
    
  }
  #Periodo 4: precovid
  if(periodo==4){
    #Se crean las variables de interés: produccion, ventas y personal 
    variacion <- base %>% 
      filter(ANIO>=2019) %>% 
      group_by(ANIO,MES) %>% 
      summarise(produccion=sum(PRODUCCIONREALPOND),
                ventas = sum(VENTASREALESPOND),
                personas = sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    #Se calculan la variaciones, sobre las variables, en el periodo de interes
    variacion$varprod2019     <- (((variacion$produccion-rep(variacion$produccion[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$produccion)])/rep(variacion$produccion[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$produccion)]))
    variacion$varventa2019    <- (((variacion$ventas-rep(variacion$ventas[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$ventas)])/rep(variacion$ventas[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$ventas)]))
    variacion$varpersonas2019 <- (((variacion$personas-rep(variacion$personas[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$personas)])/rep(variacion$personas[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$personas)]))
    etapa <- variacion  %>% select(c(starts_with("var"),"MES","ANIO")) %>% pivot_longer(cols = starts_with("var"),names_to="variables",values_to="variacion")
    
    etapa$variables <- gsub("2019","",etapa$variables)
    etapa$variables <- factor(etapa$variables,levels=c("varprod","varventa","varpersonas"),labels=c("Producción\nReal","Ventas\nReales","Personal\nOcupado"))
    
    
    etapa<-etapa %>% filter(ANIO>2019)
  }
  
  return(etapa)
}



# Funcion completar barras ------------------------------------------------

#Esta funcion permite filtrar de la tabla completa,el año y el mes que 
#seleccione la persona desde la aplicacion, para, posteriormente graficar 
# el diagrama de barras.(la funcion de grafica aparece mas abajo)
estructura_barra<-function(data,base=dataInput(),mes=input$mes,anio=input$anio){
  data<-data%>% filter(ANIO==as.numeric(anio),MES==as.numeric(mes))
  return(data)
}


#  Funcion completar series -----------------------------------------------

#Esta funcion añade columnas a la tabla para poder graficar la serie de tiempo 
#La tabla a la que se hace mención es la que se genera en la funcion estructura


estructura_serie<-function(data,base=dataInput(),mes=input$mes,anio=input$anio){

    data$periodo <- paste0(substr(meses,1,3)[as.numeric(data$MES)],"-",substr(data$ANIO,3,4))
    data$periodo <- factor(data$periodo,levels=c(paste0(rep(substr(meses,1,3),length(c(2018:max(unique(base$ANIO))))),"-",rep(18:substr(max(unique(base$ANIO)),3,4),each=12))))
    data<-data %>% mutate(Punto=as.numeric(ifelse(periodo==paste0(substr(meses,1,3)[as.numeric(mes)],"-",substr(as.numeric(anio),3,4)),variacion,NA)))
  
  return(data)
}





# Funcion barras ----------------------------------------------------------

#Funcion encargada de realizar las graficas de diagrama de barras

barras<-function(data){
  
  colores <- c("#f83796","#0875a5","#80c175")
  
  p<-ggplot(data,aes(x=variables,y=variacion,fill=variables,text = paste('</br>', variables,
                                                                         '</br> Variación: ', paste0(round(variacion*100,1)," %"))))+geom_bar(stat="identity")+
    theme_classic()+geom_text(aes(x=variables,y=variacion,label=paste0(round(variacion*100,1),"%")),vjust=-0.5)+
    scale_y_continuous(breaks = function(x)pretty_breaks(n=15)(c(as.numeric(x)[1],as.numeric(x)[2])),labels=percent)+
    scale_fill_manual(values = colores)+theme(legend.position = "none")+labs(y="Variación (%)",x="")+
    theme(strip.background.x=element_rect(color = NA,  fill=NA))
  ggplotly(p,tooltip=c("text"))

}


# Funcion series ----------------------------------------------------------

#Funcion encargada de realizar las graficas de series de tiempo

series<-function(periodo, data){
  colores <- c("#f83796","#0875a5","#80c175")  
  
  serie <- ggplot(data,aes(x=periodo,y=variacion,group=variables,color=variables))+geom_line(size=1)+
    geom_point(aes(x=max(variacion)+1,y=0),color="transparent")+
    geom_point(aes(x=periodo,y=Punto,group=variables),shape = "circle",size = 4)+
    theme_classic()+theme(axis.text.x=element_text(angle=90),legend.direction = "horizontal",legend.position="bottom")+
    scale_color_manual(values = colores)+labs(x="",y="Variación (%)",color="")+
    scale_y_continuous(breaks=function(x)pretty_breaks(n=20)(c(as.numeric(x)[1],as.numeric(x)[2])),labels = percent)+
    geom_hline(aes(yintercept=0),color="gray",size=0.5,size = 5)+
    theme(strip.background.x=element_rect(color = NA,  fill=NA),strip.text = element_text(color="transparent"))
  
  
  if(periodo==2){
    serie <- serie + facet_grid(~ANIO,scales = "free_x")
  }
  
  
  return(serie)
  
}
