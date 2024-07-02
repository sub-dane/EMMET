
# Librerias ---------------------------------------------------------------

require(ggplot2)
require(dplyr)
require(plotly)
require(readxl)


# Funcion -----------------------------------------------------------------

#Esta funcion crea las tablas que contienen las contibuciones y las variaciones 
#dependiendo los periodos consultados

tabla_CIIU<-function(periodo,base=dataCompa(),anio2=input$anio2,mes2=input$mes2){
  #Periodo 1: Mes - Anio anterior
  if(periodo==1){
    # CIIU_1 ------------------------------------------------------------------
    
    
    #Se halla la contribucion total
    contribucion_total <- base %>% 
      filter(MES==as.numeric(mes2) & ANIO%in%c(as.numeric(anio2)-1)) %>% 
      summarise(produccion_total = sum(PRODUCCIONREALPOND),
                ventas_total=sum(VENTASREALESPOND),
                personal_total=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    #Se halla la contribución específica
    contribucion <- base %>% 
      filter(MES==as.numeric(mes2) & ANIO%in%c(as.numeric(anio2),as.numeric(anio2)-1)) %>% 
      mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
      group_by(ANIO,MES,desa_num,desa_cat) %>%
      summarise(prod = sum(PRODUCCIONREALPOND),vent=sum(VENTASREALESPOND),per=sum(PERSONAL)) %>%
      group_by(desa_num,desa_cat) %>% 
      summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
                ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
                personal=(per[2]-per[1])/contribucion_total$personal_total) %>% 
      arrange(produccion)
    
    
    #Se crean la variables de interes: Produccion, ventas y personal 
    tabla1 <- base %>% 
      filter(ANIO%in%c(as.numeric(anio2),as.numeric(anio2)-1) & MES%in%c(as.numeric(mes2))) %>%
      group_by(ANIO,MES,desa_num,desa_cat) %>%
      summarise(produccion=sum(PRODUCCIONREALPOND),
                ventas = sum(VENTASREALESPOND),
                personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    #Se hallan los valores de variacion para las variables de interes
    tabla1 <- tabla1 %>% 
      pivot_wider(names_from = c("ANIO"),values_from = c("produccion","ventas","personas"))
    tabla1[paste0("varprod_",as.numeric(anio2))] <- (tabla1[paste0("produccion_",as.numeric(anio2))]-tabla1[paste0("produccion_",as.numeric(anio2)-1)])/tabla1[paste0("produccion_",as.numeric(anio2)-1)]
    tabla1[paste0("varventas_",as.numeric(anio2))]<- (tabla1[paste0("ventas_",as.numeric(anio2))]-tabla1[paste0("ventas_",as.numeric(anio2)-1)])/tabla1[paste0("ventas_",as.numeric(anio2)-1)]
    tabla1[paste0("varpersonas_",as.numeric(anio2))] <- (tabla1[paste0("personas_",as.numeric(anio2))]-tabla1[paste0("personas_",as.numeric(anio2)-1)])/tabla1[paste0("personas_",as.numeric(anio2)-1)]
    tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:3)],names_to = "variables",values_to = "value" )
    
    tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
    tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
    tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
    
    
    tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
    
    tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("desa_num","desa_cat"))
    tabla1 <- tabla1[,c("desa_num","desa_cat","varprod","produccion","varventas","ventas","varpersonas","personal")]
    
    for( i in c("varprod","produccion","varventas","ventas","varpersonas","personal")){
      tabla1[,i] <-  round(tabla1[,i]*100,2)
    }
    
    tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[c(3,5,7)],names_to = "variables",values_to = "value" )
    tabla1$desa_num <- as.character(tabla1$desa_num)
    tabla1$desa_num <- factor(tabla1$desa_num,levels=tabla1$desa_num[tabla1$variables=="varprod"][order(tabla1$produccion[tabla1$variables=="varprod"],decreasing = F)])
    tabla1$variables <- factor(tabla1$variables,levels=c("varprod","varventas","varpersonas"),labels=c("Producción","Ventas","Personal"))
    
  }
  if(periodo==2){
    #Periodo 1: Mes - Anio anterior
    # CIIU_2 ------------------------------------------------------------------
    
    #Se halla la contribucion total
    contribucion_total <- base %>%
      filter(MES%in%c(1:as.numeric(mes2)) & ANIO%in%c(as.numeric(anio2)-1)) %>%
      summarise(produccion_total = sum(PRODUCCIONREALPOND),ventas_total=sum(VENTASREALESPOND),personal_total=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    #Se halla la contribución específica
    contribucion <- base%>%
      filter(MES%in%c(1:as.numeric(mes2)) & ANIO%in%c(as.numeric(anio2),as.numeric(anio2)-1)) %>%
      mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
      group_by(ANIO,desa_num,desa_cat) %>%
      summarise(prod = sum(PRODUCCIONREALPOND),vent=sum(VENTASREALESPOND),per=sum(PERSONAL)) %>%
      group_by(desa_num,desa_cat) %>%
      summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,personal=(per[2]-per[1])/contribucion_total$personal_total) %>%
      arrange(produccion)
    
    
    #Se crean la variables de interes: Produccion, ventas y personal 
    tabla1 <- base %>%
      filter(ANIO%in%c(as.numeric(anio2),as.numeric(anio2)-1) & MES%in%c(1:as.numeric(mes2))) %>%
      group_by(ANIO,MES,desa_num,desa_cat) %>%
      summarise(produccion=sum(PRODUCCIONREALPOND),ventas = sum(VENTASREALESPOND),personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
    
    #Se hallan los valores de ccariacion para las variables de interes
    tabla1 <- tabla1 %>% pivot_wider(names_from = c("ANIO"),values_from = c("produccion","ventas","personas"))
    tabla1[paste0("varprod_",as.numeric(anio2))] <- (tabla1[paste0("produccion_",as.numeric(anio2))]-tabla1[paste0("produccion_",as.numeric(anio2)-1)])/tabla1[paste0("produccion_",as.numeric(anio2)-1)]
    tabla1[paste0("varventas_",as.numeric(anio2))]<- (tabla1[paste0("ventas_",as.numeric(anio2))]-tabla1[paste0("ventas_",as.numeric(anio2)-1)])/tabla1[paste0("ventas_",as.numeric(anio2)-1)]
    tabla1[paste0("varpersonas_",as.numeric(anio2))] <- (tabla1[paste0("personas_",as.numeric(anio2))]-tabla1[paste0("personas_",as.numeric(anio2)-1)])/tabla1[paste0("personas_",as.numeric(anio2)-1)]
    tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:3)],names_to = "variables",values_to = "value" )
    
    tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
    tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
    tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
    
    
    tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
    
    tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("desa_num","desa_cat"))
    tabla1 <- tabla1[,c("desa_num","desa_cat","varprod","produccion","varventas","ventas","varpersonas","personal")]
    
    
    for( i in c("varprod","produccion","varventas","ventas","varpersonas","personal")){
      tabla1[,i] <-  round(tabla1[,i]*100,1)
    }
    
    tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[c(3,5,7)],names_to = "variables",values_to = "value" )
    tabla1$desa_num <- as.character(tabla1$desa_num)
    tabla1$desa_num <- factor(tabla1$desa_num,levels=tabla1$desa_num[tabla1$variables=="varprod"][order(tabla1$produccion[tabla1$variables=="varprod"],decreasing = F)])
    tabla1$variables <- factor(tabla1$variables,levels=c("varprod","varventas","varpersonas"),labels=c("Producción","Ventas","Personal"))
    
    
  }
  
  if(periodo ==3){
    #Periodo 1: Mes - Anio anterior
    # CIIU_3 ------------------------------------------------------------------
    
    if(as.numeric(anio2)!=2019){
      #Se halla la contribucion total
      contribucion_total <- base %>% 
        filter(ANIO2%in%(as.numeric(anio2)-1)) %>% 
        summarise(produccion_total = sum(PRODUCCIONREALPOND),
                  ventas_total=sum(VENTASREALESPOND),
                  personal_total=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
      
      #Se halla la contribución específica
      contribucion <- base %>% filter(ANIO2%in%c(as.numeric(anio2)-1,as.numeric(anio2))) %>%
        mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
        dplyr::group_by(ANIO2,desa_num,desa_cat) %>% 
        dplyr::summarise(prod = sum(PRODUCCIONREALPOND),
                         vent=sum(VENTASREALESPOND),
                         per=sum(PERSONAL)) %>%
        dplyr::group_by(desa_num,desa_cat) %>%
        summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,
                  ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,
                  personal=(per[2]-per[1])/contribucion_total$personal_total) %>% 
        arrange(produccion)
      
      #Se crean la variables de interes: Produccion, ventas y personal 
      tabla1 <- base %>% filter(ANIO2%in%c(as.numeric(anio2),as.numeric(anio2)-1)) %>%
        group_by(ANIO2,desa_num,desa_cat) %>%
        dplyr::summarise(produccion=sum(PRODUCCIONREALPOND),
                         ventas = sum(VENTASREALESPOND),
                         personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
      
      #Se hallan los valores de ccariacion para las variables de interes
      tabla1 <- tabla1 %>% pivot_wider(names_from = c("ANIO2"),values_from = c("produccion","ventas","personas"))
      
      tabla1[paste0("varprod_",as.numeric(anio2))] <- (tabla1[paste0("produccion_",as.numeric(anio2))]-tabla1[paste0("produccion_",as.numeric(anio2)-1)])/tabla1[paste0("produccion_",as.numeric(anio2)-1)]
      tabla1[paste0("varventas_",as.numeric(anio2))]<- (tabla1[paste0("ventas_",as.numeric(anio2))]-tabla1[paste0("ventas_",as.numeric(anio2)-1)])/tabla1[paste0("ventas_",as.numeric(anio2)-1)]
      tabla1[paste0("varpersonas_",as.numeric(anio2))] <- (tabla1[paste0("personas_",as.numeric(anio2))]-tabla1[paste0("personas_",as.numeric(anio2)-1)])/tabla1[paste0("personas_",as.numeric(anio2)-1)]
      tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:3)],names_to = "variables",values_to = "value" )
      
      tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
      tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
      tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
      
      
      
      tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
      
      tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("desa_num","desa_cat"))
      tabla1 <- tabla1[,c("desa_num","desa_cat","varprod","produccion","varventas","ventas","varpersonas","personal")]
      
      for( i in c("varprod","produccion","varventas","ventas","varpersonas","personal")){
        tabla1[,i] <-  round(tabla1[,i]*100,1)
      }
      
      tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[c(3,5,7)],names_to = "variables",values_to = "value" )
      tabla1$desa_num <- as.character(tabla1$desa_num)
      tabla1$desa_num <- factor(tabla1$desa_num,levels=tabla1$desa_num[tabla1$variables=="varprod"][order(tabla1$produccion[tabla1$variables=="varprod"],decreasing = F)])
      tabla1$variables <- factor(tabla1$variables,levels=c("varprod","varventas","varpersonas"),labels=c("Producción","Ventas","Personal"))
      
      # tabla1<-tabla1 %>% mutate(value=case_when(
      #   ANIO==2019 ~ 0,
      #   TRUE ~ value
      # ))
      
    }else{tabla1<-"Esta tabla está disponiblie a partir del anio 2020"}
      
  }
    
    if(periodo==4){
      #Periodo 1: Mes - Anio anterior
      # CIIU_4 ------------------------------------------------------------------
      
      if(as.numeric(anio2)!=2019){
        #Se halla la contribucion total
        contribucion_total <- base %>%
          filter(MES==as.numeric(mes2) & ANIO%in%c(2019)) %>%
          summarise(produccion_total = sum(PRODUCCIONREALPOND),
                    ventas_total=sum(VENTASREALESPOND),
                    personal_total=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
        
        #Se halla la contribución específica
        contribucion <- base %>%
          filter(MES==as.numeric(mes2) & ANIO%in%c(as.numeric(anio2),2019)) %>%
          mutate(PERSONAL=TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC) %>%
          group_by(ANIO,MES,desa_num,desa_cat) %>%
          summarise(prod = sum(PRODUCCIONREALPOND),vent=sum(VENTASREALESPOND),per=sum(PERSONAL)) %>%
          group_by(desa_num,desa_cat) %>%
          summarise(produccion=(prod[2]-prod[1])/contribucion_total$produccion_total,ventas=(vent[2]-vent[1])/contribucion_total$ventas_total,personal=(per[2]-per[1])/contribucion_total$personal_total) %>% arrange(produccion)
        
        
        #Se crean la variables de interes: Produccion, ventas y personal 
        tabla1 <- base %>%
          filter(ANIO%in%c(as.numeric(anio2),2019) & MES%in%as.numeric(mes2)) %>%
          group_by(ANIO,MES,desa_num,desa_cat) %>%
          summarise(produccion=sum(PRODUCCIONREALPOND),
                    ventas = sum(VENTASREALESPOND),
                    personas=sum(TOTALEMPLEOPERMANENTE+TOTALEMPLEOTEMPORAL+TOTALEMPLEOADMON+TOTALEMPLEOPRODUC))
        
        #Se hallan los valores de ccariacion para las variables de interes
        tabla1 <- tabla1 %>% pivot_wider(names_from = c("ANIO"),values_from = c("produccion","ventas","personas"))
        tabla1[paste0("varprod_",as.numeric(anio2))] <- (tabla1[paste0("produccion_",as.numeric(anio2))]-tabla1[paste0("produccion_",2019)])/tabla1[paste0("produccion_",2019)]
        tabla1[paste0("varventas_",as.numeric(anio2))]<- (tabla1[paste0("ventas_",as.numeric(anio2))]-tabla1[paste0("ventas_",2019)])/tabla1[paste0("ventas_",2019)]
        tabla1[paste0("varpersonas_",as.numeric(anio2))] <- (tabla1[paste0("personas_",as.numeric(anio2))]-tabla1[paste0("personas_",2019)])/tabla1[paste0("personas_",2019)]
        tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[-c(1:3)],names_to = "variables",values_to = "value" )
        
        tabla1$ANIO <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 2)
        tabla1$variables <- sapply(strsplit(as.character(tabla1$variables), "_"), `[`, 1)
        tabla1 <- tabla1 %>% filter(gsub("var","",variables)!=variables )
        
        
        tabla1 <- tabla1 %>% pivot_wider(names_from = variables,values_from = value)
        
        tabla1 <- inner_join(x=tabla1,y=contribucion,by=c("desa_num","desa_cat"))
        tabla1 <- tabla1[,c("desa_num","desa_cat","varprod","produccion","varventas","ventas","varpersonas","personal")]
        
        for( i in c("varprod","produccion","varventas","ventas","varpersonas","personal")){
          tabla1[,i] <-  round(tabla1[,i]*100,2)
        }
        
        tabla1 <- tabla1 %>% pivot_longer(cols = colnames(tabla1)[c(3,5,7)],names_to = "variables",values_to = "value" )
        tabla1$desa_num <- as.character(tabla1$desa_num)
        tabla1$desa_num <- factor(tabla1$desa_num,levels=tabla1$desa_num[tabla1$variables=="varprod"][order(tabla1$produccion[tabla1$variables=="varprod"],decreasing = F)])
        tabla1$variables <- factor(tabla1$variables,levels=c("varprod","varventas","varpersonas"),labels=c("Producción","Ventas","Personal"))
        
        
      
      }else{tabla1<-"Esta tabla está disponiblie a partir del anio 2020"}
    }
  
  return(tabla1)
  
}


# Grafica contribucion ----------------------------------------------------

#Se produce el grafico en donde se muestra de menor a mayor por contribucion 
# y se refleja finalmente el valor de la variación

contribucion<-function(tabla,desa=input$desa){
  if(is.character(tabla)){
    p <- plotly_empty(type = "scatter", mode = "markers") %>%
      config(
        displayModeBar = FALSE
      ) %>%
      layout(
        title = tabla
      )
    ggplotly(p)
  }else{
  colores <- c("#f83796","#0875a5","#80c175")
    p <- ggplot(data=tabla,aes(x=desa_num,y=value,fill=variables))+geom_bar(stat="identity",position="dodge")+
      geom_text(data=tabla[which(tabla$variables=="Producción" & tabla$value>0),],aes(x=desa_num,y=value,label=paste0(round(produccion,3)),text = paste('</br>', desa_cat,
                                                                                                                                                        '</br> Variación: ', paste0(round(value,1)," %"),
                                                                                                                                                        '</br> Contibución: ', paste0(round(produccion,3)," p.p"))),size=2,nudge_y = 4)+
      geom_text(data=tabla[which(tabla$variables=="Ventas" & tabla$value>0),],aes(x=desa_num,y=value,label=paste0(round(ventas,2)),text = paste('</br>', desa_cat,
                                                                                                                                                '</br> Variación: ', paste0(round(value,1)," %"),
                                                                                                                                                '</br> Contibución: ', paste0(round(ventas,3)," p.p."))),size=2,nudge_y = 4)+
      geom_text(data=tabla[which(tabla$variables=="Personal"& tabla$value>0),],aes(x=desa_num,y=value,label=paste0(round(personal,2)),text = paste('</br>', desa_cat,
                                                                                                                                                   '</br> Variación: ', paste0(round(value,1)," %"),
                                                                                                                                                   '</br> Contibución: ', paste0(round(personal,3)," p.p."))),size=2,nudge_y = 4)+
      geom_text(data=tabla[which(tabla$variables=="Producción" & tabla$value<=0),],aes(x=desa_num,y=value,label=paste0(round(produccion,2)),text = paste('</br>', desa_cat,
                                                                                                                                                         '</br> Variación: ', paste0(round(value,1)," %"),
                                                                                                                                                         '</br> Contibución: ', paste0(round(produccion,3)," p.p"))),size=2,nudge_y = -4)+
      geom_text(data=tabla[which(tabla$variables=="Ventas" & tabla$value<=0),],aes(x=desa_num,y=value,label=paste0(round(ventas,2)),text = paste('</br>', desa_cat,
                                                                                                                                                 '</br> Variación: ', paste0(round(value,1)," %"),
                                                                                                                                                 '</br> Contibución: ', paste0(round(ventas,3)," p.p."))),size=2,nudge_y = -4)+
      geom_text(data=tabla[which(tabla$variables=="Personal"& tabla$value<=0),],aes(x=desa_num,y=value,label=paste0(round(personal,2)),text = paste('</br>', desa_cat,
                                                                                                                                                    '</br> Variación: ', paste0(round(value,1)," %"),
                                                                                                                                                    '</br> Contibución: ', paste0(round(personal,3)," p.p."))),size=2,nudge_y = -4)+
      
      theme_classic(base_size=8)+
      theme(legend.position = "none")+
      coord_flip()+facet_grid(~variables)+
      scale_fill_manual(values=colores,guide="none")+
      labs(y="Variación",x=desa)+
      scale_y_continuous(breaks = function(x)pretty_breaks(n=10)(c(as.numeric(x[1]),as.numeric(x[2]))))
    
    ggplotly(p,tooltip=c("text"))
    }
}


