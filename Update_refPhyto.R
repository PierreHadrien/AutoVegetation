library(openxlsx2)

#MORCEAU DE CODE POUVANT SERVIR A ENREGISTRER SOUS SON VRAI NOM DU CBNBL LE FICHIER BBE ou BSS BIV
# #récupère les noms
# wb_BBE_sheets <- wb_get_sheet_names(wb_BBE)
# wb_BSSBIV_sheets <- wb_get_sheet_names(wb_BSSBIV)
# wb_BBE_name <- wb_BBE_name[3]
# wb_BSSBIV_name <- wb_BSSBIV_name[3]
# #récupère les dates sous la bonne forme
# dfBBE <- read_xlsx(wb_BBE, sheet = 1)
# dfBSSBIV <- read_xlsx(wb_BSSBIV, sheet = 1)
# dateBBE <- dfBBE[2,2]
# dateBSSBIV <- dfBSSBIV[2,2]
# library(lubridate)
# if (month(dateBBE)<10){monthBBE <- paste0("0",month(dateBBE))}else{monthBBE <- month(dateBBE)}
# if (month(dateBSSBIV)<10){monthBSSBIV <- paste0("0",month(dateBSSBIV))}else{monthBSSBIV <- month(dateBSSBIV)}
# if (day(dateBBE)<10){dayBBE <- paste0("0",day(dateBBE))}else{dayBBE <- day(dateBBE)}
# if (day(dateBSSBIV)<10){dayBSSBIV <- paste0("0",day(dateBSSBIV))}else{dayBSSBIV <- day(dateBSSBIV)}
# dateBBE2 <- paste0(year(dateBBE),monthBBE,dayBBE,sep="")
# dateBSSBIV2 <- paste0(year(dateBSSBIV),monthBSSBIV,dayBSSBIV,sep="")
#Réécrit par dessus les fichiers existants les nouveaux
#wb_save(wb_BBE, paste0(wb_BBE_name,"_",dateBBE,.xlsx", sep = ""), overwrite = TRUE)
#wb_save(wb_BSSBIV, paste0("wb_BSSBIV_name,"_",dateBSSBIV,.xlsx", sep = ""), overwrite = TRUE)

#transforme en data frames les wb
dfBBE_s3 <- openxlsx2::read_xlsx(wb_BBE, sheet = 3, start_row = 2)
dfBSSBIV_s3 <- openxlsx2::read_xlsx(wb_BSSBIV, sheet = 3, start_row = 2)



##récupère la colonne avec synatxons et ajoute colonne Classe phyto
#isole ligne d'intérêt avec affinité phytos et où se trouvent Classes en uppercase
dfBSSBIV_s3_colclasse <- as.data.frame(dfBSSBIV_s3[,5], nm = "Classe")

#Créé une fonction pour savoir si premier mot est uppercase 
library(stringr)
firstword_uppercase <- function (x) { 
  if (grepl("^[[:upper:]]+$", word(x, start = 1L, end = 1L, sep = fixed(" "))) == TRUE) { dfBSSBIV_s3_colclasse[x,1] <- x
  }else{ dfBSSBIV_s3_colclasse[x,1] <- NA } #fonction disant si premier mot est uppercase selon TRUE FALSE
}

#code permettant respectivement de savoir si la chaine de chr est uppercase puis extraction du premier mot:
# grepl("^[A-Z]+(?:[ -][A-Z]+)*$", dfBSSBIV_s3_colclasse[1,1])
# word(dfBSSBIV_s3_colclasse[1,1], start = 1L, end = 1L, sep = fixed(" "))

#lie la colonne classe à celle avec affinités phytos et remplit celle des classes quand NA
dfBSSBIV_s3_colclasse2 <- apply( X = dfBSSBIV_s3_colclasse, MARGIN = 1, FUN = firstword_uppercase)
dfBSSBIV_s3_colclasse2 <- as.data.frame(dfBSSBIV_s3_colclasse2)
library(zoo)
dfBSSBIV_s3_colclasse2 <- na.locf(dfBSSBIV_s3_colclasse2)
dfBSSBIV_s3_classe <- as.data.frame(dfBSSBIV_s3[,5], nm = "Classe")
dfBSSBIV_s3_classe <- cbind(dfBSSBIV_s3_classe, dfBSSBIV_s3_colclasse2)

#merge espèces du réferentiel BBE avec affinités phytos et classes en uppercases
dfBBE_s3_sp <- as.data.frame(dfBBE_s3[,c(3,25)])
colnames(dfBBE_s3_sp)<-c("Nom_espece","Nom_vegetation")
colnames(dfBSSBIV_s3_classe) <- c("Nom_vegetation", "Classe")
ref_PHYTO <- merge(dfBBE_s3_sp, dfBSSBIV_s3_classe, all.x = TRUE)
ref_PHYTO <- ref_PHYTO[,c(2,1,3)]
ref_PHYTO <- ref_PHYTO[order(ref_PHYTO$Nom_espece, decreasing = FALSE), ]

#Réécrit par dessus Ref_PHYTO existant
openxlsx2::write_xlsx(ref_PHYTO, "ref_PHYTO.xlsx")
rm(list=ls())#vide l'environnement sur l'app
