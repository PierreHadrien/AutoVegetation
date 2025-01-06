library(openxlsx2)

#Chargement du relevé du tableur de saisie 
wb1 <- wb_load("Tableur saisie.xlsx", sheet = "Saisie relevé")

#ligne de code utile pour écrire dans une seul cellule ou une plage de cellules:
# write_data(wb1, sheet = "Feuil1", x=data.frame(value ="My Value"),
#            startCol = "H", startRow = 12, colNames = FALSE )

#Transformation du wb openxlsx en data frame
#(pour coeff abondance un x gris a été ajouté ligne 8)
df1 <- wb_to_df(wb1, sheet = "Saisie relevé") 
df1<-df1[-7,]
df1<-df1[-7,]

#Recherche des numéros de lignes importants pour la suite
na_vec <- which(is.na(df1[1])) 
first_na <- min(na_vec)
lnum_bryo <- which(df1$`Référence DIGITALE` == "Strate muscinale")
last_row <- nrow(df1)

## Enregistrement sp vasc et coeffs : 
  #1. Si dernière ligne NA avant la ligne de séparation avec mousses n'est pas remplie
  if (first_na != lnum_bryo){
  
    #1.1 Enregistrement et merging de la colonne d'sp vasculaires
    sp_vasc1 <- df1[8:(first_na-1),1]
    sp_vasc2 <- as.data.frame(sp_vasc1)
    ref_PHYTO <- openxlsx2::read_xlsx("ref_PHYTO.xlsx")
    colnames(sp_vasc2)[[1]] <- "Nom_espece"
    sp_vasc_merged <- merge(x = sp_vasc2, y = ref_PHYTO, all.x = TRUE)
    
    #1.2 Enregistrement de la partie abondance correspondant aux sp vasculaires
    sp_vasc_coeff <- df1[8:(first_na-1),8]
    sp_vasc_coeff <- as.data.frame(sp_vasc_coeff)
  
  #2. Si dernière ligne NA avant la ligne de séparation avec mousses est remplie
  }else{ 
  
    #2.1 Enregistrement et merging de la colonne d'sp vasculaires
    sp_vasc1 <- df1[8:(lnum_bryo-1),1]
    sp_vasc2 <- as.data.frame(sp_vasc1)
    ref_PHYTO <- openxlsx2::read_xlsx("ref_PHYTO.xlsx")
    colnames(sp_vasc2)[[1]] <- "Nom_espece"
    sp_vasc_merged <- merge(x = sp_vasc2, y = ref_PHYTO, all.x = TRUE)
    
    #2.2 Enregistrement de la partie abondance correspondant aux sp vasculaires
    sp_vasc_coeff <- df1[8:(lnum_bryo-1),8]
    sp_vasc_coeff <- as.data.frame(sp_vasc_coeff)
  
}

## Suite des enregistrements qui ne dépenent pas d'une ligne NA en fin d'sp vasc
  # Enregistrement des espèces de mousses et leurs coeffs dans deux listes
  sp_bryo <- df1[(lnum_bryo+1):last_row,1]
  sp_bryo_coeff <- df1[(lnum_bryo+1):last_row,8]
  
  # Enregistrement des annexes:l'aire échantillon, des hauteurs et recouvrements dans un objet
  annexs <- df1[1:5,8] 
  annexs <- as.data.frame(annexs)


  #Ajout de deux colonnes pour la partie sp vasculaires
  sp_vasc_merged <- cbind(sp_vasc_merged, "")
  sp_vasc_merged <- cbind(sp_vasc_merged, "")
  sp_vasc_merged <- sp_vasc_merged[,c(1,4,5,3,2)]
  sp_vasc_merged <- sp_vasc_merged[order(sp_vasc_merged$Classe),]
  
  #remplacement des NA par des blank spaces pour chaque partie enregistrée
  annexs[is.na(annexs)] <- ""
  sp_bryo[is.na(sp_bryo)] <- ""
  sp_bryo_coeff[is.na(sp_bryo_coeff)] <- ""
  sp_vasc_merged[is.na(sp_vasc_merged)] <- ""
  sp_vasc_coeff[is.na(sp_vasc_coeff)] <- ""

## Réécriture sur wb dans l'espace R et sauvegarde en excel sur PC
  
  #Copie du template vide sur le tableur de résultat avant introduction des résultats
  wb2 <- wb_load("Template.xlsx", sheet = "Feuil1")
  wb_save(wb2, "Relevé_complété.xlsx", overwrite = TRUE)
  
  #Réécriture des mousses sur le wb
  write_data(wb1, sheet = "Saisie relevé", x=data.frame(sp_bryo),
             startCol = "A", startRow = 52, colNames = FALSE )
  
  #Réécriture des coeffs de mousses sur le wb
  write_data(wb1, sheet = "Saisie relevé", x=data.frame(sp_bryo_coeff),
             startCol = "H", startRow = 52, colNames = FALSE )
  
  #Réécriture des sp vasculaires sur le wb
  write_data(wb1, sheet = "Saisie relevé", x=data.frame(sp_vasc_merged),
             startCol = "A", startRow = 11, colNames = FALSE )
  
  #Réécriture des coeffs sp vasculaires sur le wb
  write_data(wb1, sheet = "Saisie relevé", x=data.frame(sp_vasc_coeff),
             startCol = "H", startRow = 11, colNames = FALSE )
  
  #Réécriture des annexes (recouvrement, hauteur strates, etc.) sur le wb
  write_data(wb1, sheet = "Saisie relevé", x=data.frame(annexs),
             startCol = "H", startRow = 2, colNames = FALSE )
  
  #Coloration des lignes selon leurs coeffs
  list_coeff <- c("1","2","3","4","5")
  
  for (i in 1:length(list_coeff)){
    assign(paste0("sp_vasc_coeff_",list_coeff[i], sep=""), ((which(grepl(list_coeff[i], sp_vasc_coeff$sp_vasc_coeff)))+10) )
  }
  
  sp_bryo_coeff2 <- as.data.frame(sp_bryo_coeff)
  
  for (i in 1:length(list_coeff)){
    assign(paste0("sp_bryo_coeff2_",list_coeff[i], sep=""), ((which(grepl(list_coeff[i], sp_bryo_coeff2$sp_bryo_coeff)))+lnum_bryo+2) )
  }
  
  for (i in sp_vasc_coeff_1) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "D9F2D0"))
  } 
  
  for (i in sp_vasc_coeff_2) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "B3E5A1"))
  } 
  
  for (i in sp_vasc_coeff_3) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "8DD873"))
  }  
  
  for (i in sp_vasc_coeff_4) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "3A7C22"))
  } 
  
  for (i in sp_vasc_coeff_4) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_font(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "FFFFFF"))
  } 
  
  for (i in sp_vasc_coeff_5) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "275317"))
  } 
  
  for (i in sp_vasc_coeff_5) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_font(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "FFFFFF"))
  } 
  
  for (i in sp_bryo_coeff2_1) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "D9F2D0"))
  } 
  
  for (i in sp_bryo_coeff2_2) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "B3E5A1"))
  } 
  
  for (i in sp_bryo_coeff2_3) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "8DD873"))
  }  
  
  for (i in sp_bryo_coeff2_4) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "3A7C22"))
  } 
  
  for (i in sp_bryo_coeff2_4) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_font(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "FFFFFF"))
  } 
  
  for (i in sp_bryo_coeff2_5) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "275317"))
  } 
  
  for (i in sp_bryo_coeff2_5) {
    if ((is.na(i))) next
    wb1 <- wb1 %>% wb_add_font(sheet = "Saisie relevé", dims = paste0("A",i,":H",i, sep=""), color = wb_color(hex =  "FFFFFF"))
  } 
  
  
  wb1 <- wb1 %>% wb_add_fill(sheet = "Saisie relevé", dims = "G1:G1000", color = wb_color(hex =  "D9D9D9"))

  
  #Enregistrer le relevé complété sur le PC en écrasant le précédent
  wb_save(wb1, "Relevé_complété.xlsx", overwrite = TRUE)

# Ajouter ligne supplémentaire en début de liste comme dans Automate_ts de Autorapport?