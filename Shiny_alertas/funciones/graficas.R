
# Cargar librerias --------------------------------------------------------

require(ggplot2)



# Funcion barras ----------------------------------------------------------


# barras<-function(periodo,data){
#   
#   colores <- c("#f83796","#0875a5","#80c175")
#   
#   if(periodo ==1 | periodo ==2 |periodo ==3){
#     p<-ggplot(data,aes(x=variables,y=variacion,fill=variables,text = paste('</br>', variables,
#                                                                            '</br> Variación: ', paste0(round(variacion*100,1)," %"))))+geom_bar(stat="identity")+
#       theme_classic()+geom_text(aes(x=variables,y=variacion,label=paste0(round(variacion*100,1),"%")),vjust=-0.5)+
#       scale_y_continuous(breaks = function(x)pretty_breaks(n=15)(c(as.numeric(x)[1],as.numeric(x)[2])),labels=percent)+
#       scale_fill_manual(values = colores)+theme(legend.position = "none")+labs(y="Variación (%)",x="")+
#       theme(strip.background.x=element_rect(color = NA,  fill=NA))
#     ggplotly(p,tooltip=c("text"))
#     
#   }else{
#     p<-ggplot(data,aes(x=variables,y=variacion,fill=variables,text = paste('</br>', variables,
#                                                                                 '</br> Variación: ', paste0(round(variacion*100,1)," %"))))+
#       geom_bar(stat="identity")+
#       theme_classic()+facet_grid(~tipo)+geom_text(aes(x=variables,y=variacion,label=paste0(round(variacion*100,1),"%")),vjust=-0.5)+
#       scale_y_continuous(breaks = function(x)pretty_breaks(n=15)(c(as.numeric(x)[1],as.numeric(x)[2])),labels=percent)+
#       scale_fill_manual(values = colores)+theme(legend.position = "none")+labs(y="Variación (%)",x="")+
#       theme(strip.background.x=element_rect(color = NA,  fill=NA))
#     ggplotly(p,tooltip=c("text"))
#   }
#     
#  
# }


# Funcion series ----------------------------------------------------------

# series<-function(periodo, data){
#   colores <- c("#f83796","#0875a5","#80c175")  
# 
#   serie <- ggplot(data,aes(x=periodo,y=variacion,group=variables,color=variables))+geom_line(size=1)+
#     geom_point(aes(x=max(variacion)+1,y=0),color="transparent")+
#     geom_point(aes(x=periodo,y=Punto,group=variables),shape = "circle",size = 4)+
#     theme_classic()+theme(axis.text.x=element_text(angle=90),legend.direction = "horizontal",legend.position="bottom")+
#     scale_color_manual(values = colores)+labs(x="",y="Variación (%)",color="")+
#     scale_y_continuous(breaks=function(x)pretty_breaks(n=20)(c(as.numeric(x)[1],as.numeric(x)[2])),labels = percent)+
#     geom_hline(aes(yintercept=0),color="gray",size=0.5,size = 5)+
#     theme(strip.background.x=element_rect(color = NA,  fill=NA),strip.text = element_text(color="transparent"))
#   
#   
#   if(periodo==2){
#     serie <- serie + facet_grid(~ANIO,scales = "free_x")
#   }
#     
#   if(periodo==4){
#     serie <- ggplot(data,aes(x=periodo,y=variacion,group=variables,color=variables))+geom_line(size=1)+
#       geom_point(aes(x=max(variacion)+1,y=0,group=variables,color=variables),color="transparent")+
#       geom_point(aes(x=periodo,y=Punto,group=variables,color=variables),shape = "circle",size = 4)+
#       theme_classic()+theme(axis.text.x=element_text(angle=90),legend.direction = "horizontal",legend.position="bottom")+
#       scale_color_manual(values = colores)+labs(x="",y="Variación (%)",color="")+
#       scale_y_continuous(breaks=function(x)pretty_breaks(n=20)(c(as.numeric(x)[1],as.numeric(x)[2])),labels = percent)+
#       geom_hline(aes(yintercept=0),color="gray",size=0.5,size = 5)+
#       theme(strip.background.x=element_rect(color = NA,  fill=NA),strip.text = element_text(color="transparent"))
#   }
# 
# 
# return(serie)
#   
# }



# Funcion contribucion ----------------------------------------------------


contribucion<-function(tabla,desa=input$desa){
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




