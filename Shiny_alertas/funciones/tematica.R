
require(ggplot2)
require(dplyr)
require(plotly)


tematica       <-read_xlsx("D:/Documentos/CAMILA/DANE/APP/APP/bases/EMMET Base Temática definitiva Nov2022....xlsx")


colnames(tematica) <- colnames_format(tematica)
tematica<-  tematica%>% mutate_at(vars(contains("OBSE")),~str_replace_all(.,pattern="[^[:alnum:]]",replacement=" "))




# Funcion barras_tabla ----------------------------------------------------

tabla_barras<-function(periodo,base=dataInput(),mes=input$mes,anio=input$anio){
  
  if(periodo==1){
    # barra1 ------------------------------------------------------------------
    
    variacion <- base %>%
      group_by(ANIO,MES) %>%
      summarise(produccion=sum(PRODUCCIONREALPOND),
                ventas = sum(VENTASREALESPOND),
                personas = sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    variacion <- variacion %>% filter(MES==mes)
    variacion$varprod     <- ((c(NA,c(diff(variacion$produccion),NA)/variacion$produccion)[1:dim(variacion)[1]]))
    variacion$varventa    <- ((c(NA,c(diff(variacion$ventas),NA)/variacion$ventas)[1:dim(variacion)[1]]))
    variacion$varpersonas <- ((c(NA,c(diff(variacion$personas),NA)/variacion$personas)[1:dim(variacion)[1]]))
    
    variaciones <- variacion %>%
      filter(ANIO==anio) %>%
      select(starts_with("var")) %>%
      pivot_longer(cols = starts_with("var"),names_to="variables",values_to="variacion")
    variaciones$variables <- factor(variaciones$variables,levels=c("varprod","varventa","varpersonas"),labels=c("Producción\nReal","Ventas\nReales","Personal\nOcupado"))
    variaciones_totales<-variaciones
    
  }
  
  if(periodo==2){
    # barra2 ------------------------------------------------------------------
    
    
    variaciones_totales1 <- base  %>% group_by(ANIO,MES) %>% summarise(produccion=sum(PRODUCCIONREALPOND),
                                                                       ventas = sum(VENTASREALESPOND),
                                                                       personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)) %>% arrange(ANIO,MES)
    variaciones_totales1 <- variaciones_totales1 %>% group_by(ANIO) %>% mutate(suma_produccion=cumsum(produccion),
                                                                               suma_ventas   =cumsum(ventas),
                                                                               suma_personas =cumsum(personas))
    
    variaciones_totales1 <- variaciones_totales1 %>% 
      select(ANIO,MES,suma_produccion,suma_ventas,suma_personas) %>% 
      pivot_wider(names_from = c("ANIO"),values_from = c("suma_produccion","suma_ventas","suma_personas"))
    for(i in unique(base$ANIO)){
      if(i>2018){
        variaciones_totales1[paste0("varventas_",i)]<- (variaciones_totales1[paste0("suma_ventas_",i)]-variaciones_totales1[paste0("suma_ventas_",i-1)])/variaciones_totales1[paste0("suma_ventas_",i-1)]
        variaciones_totales1[paste0("varpersonas_",i)] <- (variaciones_totales1[paste0("suma_personas_",i)]-variaciones_totales1[paste0("suma_personas_",i-1)])/variaciones_totales1[paste0("suma_personas_",i-1)]
        variaciones_totales1[paste0("varprod_",i)] <- (variaciones_totales1[paste0("suma_produccion_",i)]-variaciones_totales1[paste0("suma_produccion_",i-1)])/variaciones_totales1[paste0("suma_produccion_",i-1)]
      }}
    variaciones_totales1 <- variaciones_totales1 %>%
      pivot_longer(cols = colnames(variaciones_totales1)[-1],names_to = "variables",values_to = "variacion" )
    
    variaciones_totales1$ANIO <- sapply(strsplit(as.character(variaciones_totales1$variables), "_"), `[`, 2)
    variaciones_totales1$variables <- sapply(strsplit(as.character(variaciones_totales1$variables), "_"), `[`, 1)
    
    
    variaciones_totales1 <- variaciones_totales1 %>% filter(gsub("var","",variables)!=variables)
    variaciones_totales1 <- variaciones_totales1 %>% filter(ANIO>=2019)
    variaciones_totales1$variables <- factor(variaciones_totales1$variables,levels=c("varpersonas","varprod","varventas"),labels=c("Personal Ocupado","Producción Real","Ventas Reales"))
    
    variaciones_totales<- variaciones_totales1 %>% filter(ANIO==as.numeric(anio),MES==as.numeric(mes))
    


  }
  
  if(periodo==3){
    # barra3 ------------------------------------------------------------------
    
    variaciones_totales <- base %>% 
      group_by(ANIO,MES) %>% 
      summarise(produccion=sum(PRODUCCIONREALPOND),
                ventas = sum(VENTASREALESPOND),
                personas = sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)) %>%
      arrange(ANIO,MES)
    
    variaciones_totales$suma_produccion <- NA
    variaciones_totales$suma_ventas <- NA
    variaciones_totales$suma_personas <- NA
    for(i in 1:dim(variaciones_totales)[1]){
      if(i>=12){
        variaciones_totales$suma_produccion[i] <- sum(variaciones_totales$produccion[(i-11):i])
        variaciones_totales$suma_ventas[i] <- sum(variaciones_totales$ventas[(i-11):i])
        variaciones_totales$suma_personas[i] <- sum(variaciones_totales$personas[(i-11):i])
      }
    }
    variaciones_totales <- variaciones_totales %>% select(ANIO,MES,suma_produccion,suma_ventas,suma_personas) %>% pivot_wider(names_from = c("ANIO"),values_from = c("suma_produccion","suma_ventas","suma_personas"))
    for(i in unique(base$ANIO)){
      if(i>2018){
        variaciones_totales[paste0("varprod_",i)] <- (variaciones_totales[paste0("suma_produccion_",i)]-variaciones_totales[paste0("suma_produccion_",i-1)])/variaciones_totales[paste0("suma_produccion_",i-1)]
        variaciones_totales[paste0("varventas_",i)]<- (variaciones_totales[paste0("suma_ventas_",i)]-variaciones_totales[paste0("suma_ventas_",i-1)])/variaciones_totales[paste0("suma_ventas_",i-1)]
        variaciones_totales[paste0("varpersonas_",i)] <- (variaciones_totales[paste0("suma_personas_",i)]-variaciones_totales[paste0("suma_personas_",i-1)])/variaciones_totales[paste0("suma_personas_",i-1)]
      }}
    variaciones_totales <- variaciones_totales %>% pivot_longer(cols = colnames(variaciones_totales)[-1],names_to = "variables",values_to = "variacion" )
    
    variaciones_totales$ANIO <- sapply(strsplit(as.character(variaciones_totales$variables), "_"), `[`, 2)
    variaciones_totales$variables <- sapply(strsplit(as.character(variaciones_totales$variables), "_"), `[`, 1)
    
    
    variaciones_totales <- variaciones_totales %>% filter(gsub("var","",variables)!=variables)
    variaciones_totales <- variaciones_totales %>% filter((ANIO>=2019 & MES==12) | ANIO>=2020)
    variaciones_totales$variables <- factor(variaciones_totales$variables,levels=c("varprod","varventas","varpersonas"),labels=c("Producción Real","Ventas Reales","Personal Ocupado"))
    variaciones_totales$periodo <- paste0(substr(meses,1,3)[as.numeric(variaciones_totales$MES)],"-",substr(variaciones_totales$ANIO,3,4))
    variaciones_totales$periodo <- factor(variaciones_totales$periodo,levels=c(paste0(rep(substr(meses,1,3),length(c(2018:max(unique(base$ANIO))))),"-",rep(18:substr(max(unique(base$ANIO)),3,4),each=12))))
    
    variaciones_totales<-variaciones_totales %>% filter(ANIO==as.numeric(anio),MES==as.numeric(mes))
    
  }
  
  
  if(periodo==4){
    # barra4 ------------------------------------------------------------------
    
    variacion <- base %>%
      group_by(ANIO,MES) %>%
      summarise(produccion=sum(PRODUCCIONREALPOND),
                ventas = sum(VENTASREALESPOND),
                personas = sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    variacion <- variacion %>% filter(MES==mes)
    variacion$varprod     <- ((c(NA,c(diff(variacion$produccion),NA)/variacion$produccion)[1:dim(variacion)[1]]))
    variacion$varventa    <- ((c(NA,c(diff(variacion$ventas),NA)/variacion$ventas)[1:dim(variacion)[1]]))
    variacion$varpersonas <- ((c(NA,c(diff(variacion$personas),NA)/variacion$personas)[1:dim(variacion)[1]]))

    variacion$varprod2019     <- (((variacion$produccion-variacion$produccion[variacion$ANIO==2019])/variacion$produccion[variacion$ANIO==2019]))
    variacion$varventa2019    <- (((variacion$ventas-variacion$ventas[variacion$ANIO==2019])/variacion$ventas[variacion$ANIO==2019]))
    variacion$varpersonas2019 <- (((variacion$personas-variacion$personas[variacion$ANIO==2019])/variacion$personas[variacion$ANIO==2019]))
    
    variacion<-variacion %>% mutate(varprod2019=case_when(
      ANIO==2018 ~ 0,
      TRUE ~ varprod2019
    ))
    variacion<-variacion %>% mutate(varpersonas2019=case_when(
      ANIO==2018 ~ 0,
      TRUE ~ varpersonas2019
    ))
    variacion<-variacion %>% mutate(varventa2019=case_when(
      ANIO==2018 ~ 0,
      TRUE ~ varventa2019
    ))
    
    variacion<-variacion %>% filter(ANIO==as.numeric(anio),MES==as.numeric(mes))
    variaciones <- variacion %>% select(starts_with("var")) %>% pivot_longer(cols = starts_with("var"),names_to="variable",values_to="variacion")
    
    variaciones<-variaciones[4:nrow(variaciones),]
    
    variaciones$tipo <- ifelse(gsub("2019","",variaciones$variable)!=variaciones$variable,
                               paste0(meses[as.numeric(mes)], " (",anio," / 2019)"),paste0(meses[as.numeric(mes)]," (",anio," / ",as.numeric(anio)-1,")"))
    
    
    
    variaciones$variable <- gsub("2019","",variaciones$variable)
    variaciones$variable <- factor(variaciones$variable,levels=c("varprod","varventa","varpersonas"),labels=c("Producción\nReal","Ventas\nReales","Personal\nOcupado"))
    variaciones$tipo <- factor(variaciones$tipo,levels=c(paste0(meses[as.numeric(mes)], " (",anio," / 2019)")))
    
    variaciones_totales<-variaciones
  }
  
  return(variaciones_totales)
}




tabla_series<-function(periodo,base=dataInput(),mes=input$mes,anio=input$anio){
  
  if(periodo == 1){
    # serie1 ------------------------------------------------------------------
    
    variaciones_totales <- base  %>% 
      group_by(ANIO,MES) %>% 
      summarise(produccion=sum(PRODUCCIONREALPOND),
                ventas = sum(VENTASREALESPOND),
                personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    variaciones_totales <- variaciones_totales %>% 
      pivot_wider(names_from = c("ANIO"),values_from = c("produccion","ventas","personas"))
    for(i in unique(base$ANIO)){
      if(i>2018){
        variaciones_totales[paste0("varprod_",i)] <- (variaciones_totales[paste0("produccion_",i)]-variaciones_totales[paste0("produccion_",i-1)])/variaciones_totales[paste0("produccion_",i-1)]
        variaciones_totales[paste0("varventas_",i)]<- (variaciones_totales[paste0("ventas_",i)]-variaciones_totales[paste0("ventas_",i-1)])/variaciones_totales[paste0("ventas_",i-1)]
        variaciones_totales[paste0("varpersonas_",i)] <- (variaciones_totales[paste0("personas_",i)]-variaciones_totales[paste0("personas_",i-1)])/variaciones_totales[paste0("personas_",i-1)]
      }}
    variaciones_totales <- variaciones_totales %>% pivot_longer(cols = colnames(variaciones_totales)[-1],names_to = "variables",values_to = "value" )
    
    variaciones_totales$ANIO <- sapply(strsplit(as.character(variaciones_totales$variables), "_"), `[`, 2)
    variaciones_totales$variables <- sapply(strsplit(as.character(variaciones_totales$variables), "_"), `[`, 1)
    
    
    variaciones_totales <- variaciones_totales %>% filter(gsub("var","",variables)!=variables)
    variaciones_totales$variables <- factor(variaciones_totales$variables,levels=c("varprod","varventas","varpersonas"),labels=c("Producción Real","Ventas Reales","Personal Ocupado"))
    variaciones_totales$periodo   <- paste0(substr(meses,1,3)[as.numeric(variaciones_totales$MES)],"-",substr(variaciones_totales$ANIO,3,4))
    variaciones_totales$periodo   <- factor(variaciones_totales$periodo,levels=c(paste0(rep(substr(meses,1,3),length(c(2018:max(unique(base$ANIO))))),"-",rep(18:substr(max(unique(base$ANIO)),3,4),each=12))))
    variaciones_totales<-variaciones_totales %>% mutate(Punto=as.numeric(ifelse(periodo==paste0(substr(meses,1,3)[as.numeric(mes)],"-",substr(as.numeric(anio),3,4)),value,NA)))
    
  }
  
  if(periodo == 2){
    # serie2 ------------------------------------------------------------------
    
    variaciones_totales <- base  %>% group_by(ANIO,MES) %>% summarise(produccion=sum(PRODUCCIONREALPOND),
                                                                             ventas = sum(VENTASREALESPOND),
                                                                             personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)) %>% arrange(ANIO,MES)
    variaciones_totales <- variaciones_totales %>% group_by(ANIO) %>% mutate(suma_produccion=cumsum(produccion),
                                                                             suma_ventas   =cumsum(ventas),
                                                                             suma_personas =cumsum(personas))
    
    variaciones_totales <- variaciones_totales %>% select(ANIO,MES,suma_produccion,suma_ventas,suma_personas) %>% pivot_wider(names_from = c("ANIO"),values_from = c("suma_produccion","suma_ventas","suma_personas"))
    for(i in unique(base$ANIO)){
      if(i>2018){
        variaciones_totales[paste0("varventas_",i)]<- (variaciones_totales[paste0("suma_ventas_",i)]-variaciones_totales[paste0("suma_ventas_",i-1)])/variaciones_totales[paste0("suma_ventas_",i-1)]
        variaciones_totales[paste0("varpersonas_",i)] <- (variaciones_totales[paste0("suma_personas_",i)]-variaciones_totales[paste0("suma_personas_",i-1)])/variaciones_totales[paste0("suma_personas_",i-1)]
        variaciones_totales[paste0("varprod_",i)] <- (variaciones_totales[paste0("suma_produccion_",i)]-variaciones_totales[paste0("suma_produccion_",i-1)])/variaciones_totales[paste0("suma_produccion_",i-1)]
      }}
    variaciones_totales <- variaciones_totales %>% pivot_longer(cols = colnames(variaciones_totales)[-1],names_to = "variables",values_to = "value" )
    
    variaciones_totales$ANIO <- sapply(strsplit(as.character(variaciones_totales$variables), "_"), `[`, 2)
    variaciones_totales$variables <- sapply(strsplit(as.character(variaciones_totales$variables), "_"), `[`, 1)
    
    
    variaciones_totales <- variaciones_totales %>% filter(gsub("var","",variables)!=variables)
    variaciones_totales <- variaciones_totales %>% filter(ANIO>=2019)
    variaciones_totales$variables <- factor(variaciones_totales$variables,levels=c("varpersonas","varprod","varventas"),labels=c("Personal Ocupado","Producción Real","Ventas Reales"))
    variaciones_totales$periodo <- paste0(substr(meses,1,3)[as.numeric(variaciones_totales$MES)],"-",substr(variaciones_totales$ANIO,3,4))
    variaciones_totales$periodo <- factor(variaciones_totales$periodo,levels=c(paste0(rep(substr(meses,1,3),length(c(2018:max(unique(base$ANIO))))),"-",rep(18:substr(max(unique(base$ANIO)),3,4),each=12))))
    variaciones_totales<-variaciones_totales %>% mutate(Punto=as.numeric(ifelse(periodo==paste0(substr(meses,1,3)[as.numeric(mes)],"-",substr(as.numeric(anio),3,4)),value,NA)))
    
  }
  
  if(periodo==3){
    # serie3 ------------------------------------------------------------------
    
    variaciones_totales <- base  %>% group_by(ANIO,MES) %>% summarise(produccion=sum(PRODUCCIONREALPOND),
                                                                             ventas = sum(VENTASREALESPOND),
                                                                             personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC)) %>% arrange(ANIO,MES)
    
    variaciones_totales$suma_produccion <- NA
    variaciones_totales$suma_ventas <- NA
    variaciones_totales$suma_personas <- NA
    for(i in 1:dim(variaciones_totales)[1]){
      if(i>=12){
        variaciones_totales$suma_produccion[i] <- sum(variaciones_totales$produccion[(i-11):i])
        variaciones_totales$suma_ventas[i] <- sum(variaciones_totales$ventas[(i-11):i])
        variaciones_totales$suma_personas[i] <- sum(variaciones_totales$personas[(i-11):i])
      }
    }
    variaciones_totales <- variaciones_totales %>% select(ANIO,MES,suma_produccion,suma_ventas,suma_personas) %>% pivot_wider(names_from = c("ANIO"),values_from = c("suma_produccion","suma_ventas","suma_personas"))
    for(i in unique(base$ANIO)){
      if(i>2018){
        variaciones_totales[paste0("varprod_",i)] <- (variaciones_totales[paste0("suma_produccion_",i)]-variaciones_totales[paste0("suma_produccion_",i-1)])/variaciones_totales[paste0("suma_produccion_",i-1)]
        variaciones_totales[paste0("varventas_",i)]<- (variaciones_totales[paste0("suma_ventas_",i)]-variaciones_totales[paste0("suma_ventas_",i-1)])/variaciones_totales[paste0("suma_ventas_",i-1)]
        variaciones_totales[paste0("varpersonas_",i)] <- (variaciones_totales[paste0("suma_personas_",i)]-variaciones_totales[paste0("suma_personas_",i-1)])/variaciones_totales[paste0("suma_personas_",i-1)]
      }}
    variaciones_totales <- variaciones_totales %>% pivot_longer(cols = colnames(variaciones_totales)[-1],names_to = "variables",values_to = "value" )
    
    variaciones_totales$ANIO <- sapply(strsplit(as.character(variaciones_totales$variables), "_"), `[`, 2)
    variaciones_totales$variables <- sapply(strsplit(as.character(variaciones_totales$variables), "_"), `[`, 1)
    
    
    variaciones_totales <- variaciones_totales %>% filter(gsub("var","",variables)!=variables)
    variaciones_totales <- variaciones_totales %>% filter((ANIO>=2019 & MES==12) | ANIO>=2020)
    variaciones_totales$variables <- factor(variaciones_totales$variables,levels=c("varprod","varventas","varpersonas"),labels=c("Producción Real","Ventas Reales","Personal Ocupado"))
    variaciones_totales$periodo <- paste0(substr(meses,1,3)[as.numeric(variaciones_totales$MES)],"-",substr(variaciones_totales$ANIO,3,4))
    variaciones_totales$periodo <- factor(variaciones_totales$periodo,levels=c(paste0(rep(substr(meses,1,3),length(c(2018:max(unique(base$ANIO))))),"-",rep(18:substr(max(unique(base$ANIO)),3,4),each=12))))
    variaciones_totales<-variaciones_totales %>% mutate(Punto=as.numeric(ifelse(periodo==paste0(substr(meses,1,3)[as.numeric(mes)],"-",substr(as.numeric(anio),3,4)),value,NA)))
    
  }
  
  
  if(periodo==4){
    # serie4 ------------------------------------------------------------------
    
    variacion <- base  %>% filter(ANIO>=2019) %>% group_by(ANIO,MES) %>% summarise(produccion=sum(PRODUCCIONREALPOND),
                                                                                          ventas = sum(VENTASREALESPOND),
                                                                                          personas = sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    variacion$varprod2019     <- (((variacion$produccion-rep(variacion$produccion[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$produccion)])/rep(variacion$produccion[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$produccion)]))
    variacion$varventa2019    <- (((variacion$ventas-rep(variacion$ventas[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$ventas)])/rep(variacion$ventas[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$ventas)]))
    variacion$varpersonas2019 <- (((variacion$personas-rep(variacion$personas[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$personas)])/rep(variacion$personas[variacion$ANIO==2019],length(unique(base$ANIO)))[1:length(variacion$personas)]))
    variaciones <- variacion  %>% select(c(starts_with("var"),"MES","ANIO")) %>% pivot_longer(cols = starts_with("var"),names_to="variable",values_to="variacion")
    
    variaciones$periodo <- paste0(substr(meses,1,3)[as.numeric(variaciones$MES)],"-",substr(variaciones$ANIO,3,4))
    variaciones$periodo <- factor(variaciones$periodo,levels=c(paste0(rep(substr(meses,1,3),length(c(2018:max(unique(base$ANIO))))),"-",rep(18:substr(max(unique(base$ANIO)),3,4),each=12))))
    
    variaciones$variable <- gsub("2019","",variaciones$variable)
    variaciones$variable <- factor(variaciones$variable,levels=c("varprod","varventa","varpersonas"),labels=c("Producción\nReal","Ventas\nReales","Personal\nOcupado"))
    
    
    variaciones<-variaciones %>% filter(ANIO>2019)
    
    variaciones_totales <- variaciones %>% mutate(Punto=as.numeric(ifelse(periodo==paste0(substr(meses,1,3)[as.numeric(mes)],"-",substr(as.numeric(anio),3,4)),variacion,NA)))
    
  }
  return(variaciones_totales)
}



# Funcion barras ----------------------------------------------------------


barras<-function(periodo,data){
  
  colores <- c("#f83796","#0875a5","#80c175")
  
  if(periodo ==1 | periodo ==2 |periodo ==3){
    p<-ggplot(data,aes(x=variables,y=variacion,fill=variables,text = paste('</br>', variables,
                                                                           '</br> Variación: ', paste0(round(variacion*100,1)," %"))))+geom_bar(stat="identity")+
      theme_classic()+geom_text(aes(x=variables,y=variacion,label=paste0(round(variacion*100,1),"%")),vjust=-0.5)+
      scale_y_continuous(breaks = function(x)pretty_breaks(n=15)(c(as.numeric(x)[1],as.numeric(x)[2])),labels=percent)+
      scale_fill_manual(values = colores)+theme(legend.position = "none")+labs(y="Variación (%)",x="")+
      theme(strip.background.x=element_rect(color = NA,  fill=NA))
    ggplotly(p,tooltip=c("text"))
    
  }else{
    p<-ggplot(data,aes(x=variable,y=variacion,fill=variable,text = paste('</br>', variable,
                                                                         '</br> Variación: ', paste0(round(variacion*100,1)," %"))))+
      geom_bar(stat="identity")+
      theme_classic()+facet_grid(~tipo)+geom_text(aes(x=variable,y=variacion,label=paste0(round(variacion*100,1),"%")),vjust=-0.5)+
      scale_y_continuous(breaks = function(x)pretty_breaks(n=15)(c(as.numeric(x)[1],as.numeric(x)[2])),labels=percent)+
      scale_fill_manual(values = colores)+theme(legend.position = "none")+labs(y="Variación (%)",x="")+
      theme(strip.background.x=element_rect(color = NA,  fill=NA))
    ggplotly(p,tooltip=c("text"))
  }
  
  
}


# Funcion series ----------------------------------------------------------

grafica_series<-function(periodo, data){
  colores <- c("#f83796","#0875a5","#80c175")  
  
  serie <- ggplot(data,aes(x=periodo,y=value,group=variables,color=variables))+geom_line(size=1)+
    geom_point(aes(x=max(value)+1,y=0),color="transparent")+
    geom_point(aes(x=periodo,y=Punto,group=variables),shape = "circle",size = 4)+
    theme_classic()+theme(axis.text.x=element_text(angle=90),legend.direction = "horizontal",legend.position="bottom")+
    scale_color_manual(values = colores)+labs(x="",y="Variación (%)",color="")+
    scale_y_continuous(breaks=function(x)pretty_breaks(n=20)(c(as.numeric(x)[1],as.numeric(x)[2])),labels = percent)+
    geom_hline(aes(yintercept=0),color="gray",size=0.5,size = 5)+
    theme(strip.background.x=element_rect(color = NA,  fill=NA),strip.text = element_text(color="transparent"))
  
  
  if(periodo==2){
    serie <- serie + facet_grid(~ANIO,scales = "free_x")
  }
  
  if(periodo==4){
    serie <- ggplot(data,aes(x=periodo,y=variacion,group=variable,color=variable))+geom_line(size=1)+
      geom_point(aes(x=max(variacion)+1,y=0,group=variable,color=variable),color="transparent")+
      geom_point(aes(x=periodo,y=Punto,group=variable,color=variable),shape = "circle",size = 4)+
      theme_classic()+theme(axis.text.x=element_text(angle=90),legend.direction = "horizontal",legend.position="bottom")+
      scale_color_manual(values = colores)+labs(x="",y="Variación (%)",color="")+
      scale_y_continuous(breaks=function(x)pretty_breaks(n=20)(c(as.numeric(x)[1],as.numeric(x)[2])),labels = percent)+
      geom_hline(aes(yintercept=0),color="gray",size=0.5,size = 5)+
      theme(strip.background.x=element_rect(color = NA,  fill=NA),strip.text = element_text(color="transparent"))
  }
  
  
  return(serie)
  
}






