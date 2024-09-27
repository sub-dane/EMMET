#utils
library(openxlsx)

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

addSuperSubScriptToCell_general <- function(wb,
                                            sheet,
                                            row,
                                            col,
                                            texto,
                                            size = '9',
                                            colour = '000000',
                                            font = "Segoe UI",
                                            family = '2',
                                            bold = TRUE,
                                            italic = FALSE,
                                            underlined = FALSE) {
  
  placeholderText <- 'This is placeholder text that should not appear anywhere in your document.'
  
  openxlsx::writeData(wb = wb,
                      sheet = sheet,
                      x = placeholderText,
                      startRow = row,
                      startCol = col)
  
  #finds the string that you want to update
  stringToUpdate <- which(sapply(wb$sharedStrings,
                                 function(x){
                                   grep(pattern = placeholderText,
                                        x)
                                 }
  )
  == 1)
  
  #splits the text into normal text, superscript and subcript
  
  normal_text <- str_split(texto, "\\[.*\\]|~.*~") %>% pluck(1) %>% purrr::discard(~ . == "")
  
  sub_sup_text <- str_extract_all(texto, "\\[.*\\]|~.*~") %>% pluck(1)
  
  if (length(normal_text) > length(sub_sup_text)) {
    sub_sup_text <- c(sub_sup_text, "")
  } else if (length(sub_sup_text) > length(normal_text)) {
    normal_text <- c(normal_text, "")
  }
  # this is the separated text which will be used next
  texto_separado <- map2(normal_text, sub_sup_text, ~ c(.x, .y)) %>% 
    reduce(c) %>% 
    purrr::discard(~ . == "")
  
  #formatting instructions
  
  sz    <- paste('<sz val =\"',size,'\"/>',
                 sep = '')
  col   <- paste('<color rgb =\"',colour,'\"/>',
                 sep = '')
  rFont <- paste('<rFont val =\"',font,'\"/>',
                 sep = '')
  fam   <- paste('<family val =\"',family,'\"/>',
                 sep = '')
  
  #if its sub or sup adds the corresponding xml code
  sub_sup_no <- function(texto) {
    
    if(str_detect(texto, "\\[.*\\]")){
      return('<vertAlign val=\"superscript\"/>')
    } else if (str_detect(texto, "~.*~")) {
      return('<vertAlign val=\"subscript\"/>')
    } else {
      return('')
    }
  }
  
  #get text from normal text, sub and sup
  get_text_sub_sup <- function(texto) {
    str_remove_all(texto, "\\[|\\]|~")
  }
  
  #formating
  if(bold){
    bld <- '<b/>'
  } else{bld <- ''}
  
  if(italic){
    itl <- '<i/>'
  } else{itl <- ''}
  
  if(underlined){
    uld <- '<u/>'
  } else{uld <- ''}
  
  #get all properties from one element of texto_separado
  
  get_all_properties <- function(texto) {
    
    paste0('<r><rPr>',
           sub_sup_no(texto),
           sz,
           col,
           rFont,
           fam,
           bld,
           itl,
           uld,
           '</rPr><t xml:space="preserve">',
           get_text_sub_sup(texto),
           '</t></r>')
  }
  
  
  # use above function in texto_separado
  newString <- map(texto_separado, ~ get_all_properties(.)) %>% 
    reduce(paste, sep = "") %>% 
    {c("<si>", ., "</si>")} %>% 
    reduce(paste, sep = "")
  
  # replace initial text
  wb$sharedStrings[stringToUpdate] <- newString
}


filas_blanco<- function(df) {
  blank_row <- df[1, ]
  blank_row[,] <- NA
  
  df_with_blank <- df %>%
    group_by(ORDEN_DEPTO) %>%
    do(rbind(., blank_row)) 
  return(df_with_blank)
}

# Formatos ----------------------------------------------------------------
conte <- createStyle(
  fontName = "Segoe UI",
  fontSize = 14,
  fontColour = "#000000",
  halign = "center",
  valign = "center",
  textDecoration = "bold"
)

colgr <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  fgFill = "#F2F2F2",
  bgFill = "#FFFFFF",
  halign = "center",
  valign = "center",
  numFmt = "0.0"
)


colbl <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  fgFill = "#FFFFFF",
  halign = "center",
  valign = "center",
  numFmt = "0.0"
)

colgr_in <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  fgFill = "#F2F2F2",
  bgFill = "#FFFFFF",
  numFmt = "0"
)


colbl_in <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  fgFill = "#FFFFFF",
  numFmt = "0"
)    
ultbl <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  border = "Bottom: thin",
  borderColour = "#000000",
  fgFill = "#FFFFFF",
  bgFill = "#FFFFFF",
  halign = "center",
  valign = "center"
)

ultblfc <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  border = "Bottom: thin, Right: thin",
  borderColour = "#000000",
  fgFill = "#FFFFFF",
  bgFill = "#FFFFFF",
  halign = "center",
  valign = "center"
)

ultcgr <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  border = "Right",
  borderColour = "#000000",
  halign = "center",
  valign = "center",
  fgFill = "#F2F2F2",
  bgFill = "#FFFFFF",
  numFmt = "0.0"
)


ultcbl <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  border = "Right",
  borderColour = "#000000",
  halign = "center",
  valign = "center",
  fgFill = "#FFFFFF",
  bgFill = "#FFFFFF",
  numFmt = "0.0"
)

rowbl <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  border = "Top",
  borderColour = "#000000",
  halign = "left",
  fgFill = "#FFFFFF",
  bgFill = "#FFFFFF"
)

rowblf <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  border ="Top: thin, Right: thin",
  borderColour = "#000000",
  halign = "left",
  fgFill = "#FFFFFF",
  bgFill = "#FFFFFF"
)  

ultrbl <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  border = "Bottom: thin",
  borderColour = "#000000",
  fgFill = "#FFFFFF",
  bgFill = "#FFFFFF",
)
ultrblf <- createStyle(
  fontName = "Segoe UI",
  fontSize = 9,
  fontColour = "#000000",
  border = "Bottom: thin, Right: thin",
  borderColour = "#000000",
  fgFill = "#FFFFFF",
  bgFill = "#FFFFFF",
)  
