library(shiny)
library(shinythemes)
library(bslib)
library(bsicons)
library(shinyWidgets)
library(rmarkdown)


ui <- page_navbar( 
  title = strong("AutoVégétation", style = "font-family: 'Pristina'; font-size:45px"),
  underline = TRUE, 
  
  theme = bs_theme(
    
    bg = "#EEE4D8", fg = "#3E624B", primary = "#945C40", secondary = "#BD8467",
    info = "#04798D", warning = "#D29F06", danger = "#B72A37", success = "#1D6F3A",
    base_font = "Noto sans", font_scale = NULL, 
    `enable-gradients` = TRUE, preset = "bootstrap", version = version_default()
  ),
  #588969 
  
  #################################################################
  # Onglet completion des relevés avec affinité phytosociologique 
  #################################################################

  nav_panel(title = div("Relevé phyto", style = "font-family: 'Noto sans'; font-size:20px"),
           
            #1.Sélection du type d'entrée
            card( max_height = 120,
                  card_header( #class = "bg-primary",
                    "1. Sélection du type d'entrée"
                  ),
                  card_body( #class = "text-primary",
                    radioGroupButtons(inputId = "type_entree_releve",
                                      label = NULL, #supprime ligne du label
                                      c("Tableur saisie" = "ts", "Fichier depuis Qfield" = "qf", "Fichier depuis Mes Relevés" = "mr"),
                                      direction = "horizontal",
                                      justified = TRUE,
                                      size = "sm",
                                      status = "secondary"
                    )
                  ), 
            ),
            
            #2. Deux cartes pour deux types d'entrée (fichier ou tableur saisie)
            navset_hidden( id = "container",
                           
                           #2.1 première carte qui apparait quand tableur saisie sélectionné
                           nav_panel_hidden("ts",
                                            card( max_height = 160,
                                                  card_header( #class = "bg-primary",
                                                    "2. Saisie du tableur avec onglets déroulants"
                                                  ),
                                                  card_body( class = "align-items-center",
                                                             actionButton("button_ts", class = "btn btn-secondary text-white", "Tableur de saisie", width = 200)
                                                  )
                                            )
                           ),
                           
                           #2.2 deuxième carte qui apparait quand qfield sélectionné
                           nav_panel_hidden("qf",      
                                            card( max_height = 160,
                                                  card_header( #class = "bg-primary",
                                                    "2. Chargement d'un fichier"
                                                  ),
                                                  card_body( class = "align-items-center",
                                                             "A venir..."
                                                             # fileInput("upload", 
                                                             #           label = NULL,
                                                             #           buttonLabel = "Charger un fichier extérieur",
                                                             #           placeholder = "Aucun fichier sélectionné",
                                                             #           width = 450)  
                                                  )
                                            )
                           ),
                           
                           #2.3 Troisième carte qui apparait quand Mesrelevés sélectionné
                           nav_panel_hidden("mr",      
                                            card( max_height = 160,
                                                  card_header( #class = "bg-primary",
                                                    "2. Chargement d'un fichier"
                                                  ),
                                                  card_body( class = "align-items-center",
                                                             "A venir..."
                                                             # fileInput("upload",
                                                             #           label = NULL,
                                                             #           buttonLabel = "Charger un fichier extérieur",
                                                             #           placeholder = "Aucun fichier sélectionné",
                                                             #           width = 450)
                                                  )
                                            )
                           )    
                           
            ),#navset_hidden
            
            
            #3.Bouton pour lancer le script de complétion du relevé phytosociologique
            card( max_height = 130,
                  card_header( #class = "bg-primary",
                    "3. Obtention du relevé complété"
                  ),
                  card_body( class = "text-white",
                             fluidRow(
                               column(12, align = "center",
                                      actionButton("button_launch_rp", class = "btn btn-success btn-lg", "Lancer", width = 300) ))       
                  ), 
            )
            
  ),#fermeture nav_panel_hidden (onglet completion affinité phytosociologique)
  
  
  ########################################
  # Onglet complétion statuts végétations
  ########################################
  
  nav_panel(title = div("Statut Veg", style = "font-family: 'Noto sans'; font-size:20px"),
            
         p("A venir")
            
            ),
  
  
  ##################################
  # Onglet dark mode et paramètres
  ##################################  
  
  nav_spacer(),#éloigne vers la droite le titre-à-cliquer
  nav_panel(title =NULL, icon=icon("cog", style = "font-size: 20px;"), 
      layout_column_wrap(      
            card( id = "card_refPHYTO", max_height = 600,
                  card_header( class = "bg-primary",
                               "Mise à jour des référentiels pour la complétion des relevés phytosociologiques"),
                  card_body( class = "align-items-center",
                             "Chargez les référentiels les plus récents du Conservatoire Botanique National de Bailleul",
#Référentiel biologique et écologique sur la flore du nord de la France (BBE) & Référentiels syntaxonomiques régionaux de la végétation du nord de la France (BSS-BIV)                            
                             fileInput(inputId = "Ref_BBE_CBNBL", 
                                       label = strong("Chargez le tableur BBE (avec colonne affinité phytosociologique) au format .xlsx"),
                                       buttonLabel = "Charger",
                                       placeholder = "Aucun fichier sélectionné",
                                       accept = ".xlsx"),
                             fileInput(inputId = "Ref_BSSBIV_CBNBL", 
                                       label = strong("Chargez le tableur BSS-BIV (syntaxons) au format .xlsx"),
                                       buttonLabel = "Charger",
                                       placeholder = "Aucun fichier sélectionné",
                                       accept = ".xlsx"),
                              actionButton("mj_refPHYTO", class = "btn btn-success", "Mettre à jour", width = 200)
                  )
                ),
            card( id = "card_PV", max_height = 600,
                  card_header( class = "bg-primary",
                               "Mise à jour du référentiel de la flore vasculaire"),
                  card_body( class = "align-items-center",
                             "Chargez le référentiel le plus récent du Conservatoire Botanique National de Bailleul",
                             #Référentiel biologique et écologique sur la flore du nord de la France (BBE) & Référentiels syntaxonomiques régionaux de la végétation du nord de la France (BSS-BIV)                            
                             fileInput(inputId = "Ref_PV_CBNBL", 
                                       label = strong("Chargez le tableur PV (plantes vasculaires) au format .xlsx"),
                                       buttonLabel = "Charger",
                                       placeholder = "Aucun fichier sélectionné",
                                       accept = ".xlsx"),
                             # actionButton("mj_bryo", class = "btn btn-success", "Mettre à jour", width = 200)
                             "A venir..."
                     )
                  ),
            card( id = "card_MH", max_height = 600,
                  card_header( class = "bg-primary",
                               "Mise à jour du référentiel des Bryophytes"),
                  card_body( class = "align-items-center",
                             "Chargez le référentiel le plus récent du Conservatoire Botanique National de Bailleul",
                             #Référentiel biologique et écologique sur la flore du nord de la France (BBE) & Référentiels syntaxonomiques régionaux de la végétation du nord de la France (BSS-BIV)                            
                             fileInput(inputId = "Ref_MH_CBNBL", 
                                       label = strong("Chargez le tableur MH (mousses) au format .xlsx"),
                                       buttonLabel = "Charger",
                                       placeholder = "Aucun fichier sélectionné",
                                       accept = ".xlsx"),
                             # actionButton("mj_bryo", class = "btn btn-success", "Mettre à jour", width = 200)
                             "A venir..."
                  )
            )
      )
 ),
  nav_panel_hidden(" "),
  nav_item(input_dark_mode(id = "dark_mode", mode = "light")), #mode nuit en switch pour l'interface
  nav_panel_hidden(" ")
  
)#fermeture page_navbar





server <- function(input, output, session) {
  
  ##1.Relevé Phyto
  #bouton multi-choix dans affinite pyto
  observe(nav_select("container", input$type_entree_releve))
  
  #Tableur saisie pour les relevés: ouverture
  observeEvent(input$button_ts,{
    shell.exec("Tableur saisie.xlsx")
  })
  
  #Lancement complétion
  observeEvent(input$button_launch_rp,{
    if (input$type_entree_releve == "ts"){
      source("Relevé_phyto_ts.R")
      shell.exec("Relevé_complété.xlsx")
    }
    if (input$type_entree_releve == "qf"){
      print("non disponible")
    }
    if (input$type_entree_releve == "mr"){
      print("non disponible")
    }
  })
  
  ##2. Paramètres
  #Remplacement de ref_Phyto
  observeEvent(input$mj_refPHYTO,{
    req(input$Ref_BBE_CBNBL)
    req(input$Ref_BSSBIV_CBNBL) 
    library(openxlsx2)
    wb_BBE <<- wb_load(input$Ref_BBE_CBNBL$datapath)
    wb_BSSBIV <<- wb_load(input$Ref_BSSBIV_CBNBL$datapath)
    source("Update_refPhyto.R")
    rm(list=ls())#vide l'environnement sur le PC (pas l'app)
  })
  
  
}#fermeture function server

# shinyApp(ui,server)
shinyApp(ui = ui, server = server)

