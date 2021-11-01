# *-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-**-*
#==  Programa para el Desarrollo de factores de expansión de la encuesta turismo 2021 ===== 
# *-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-**-*
#
#
# Héctor Garrido Henríquez
# Analista Cuantitativo. Observatorio Laboral Ñuble
# Universidad del Bío-Bío
# Avenida Andrés Bello 720, Casilla 447, Chillán
# Teléfono: +56-942353973
# http://www.observatoriolaboralnuble.cl


rm(list = ls())

## Se crea una funcón para cargar y/o instalar los paquetes necesarios para la rutina

load_pkg <- function(pack){
  create.pkg <- pack[!(pack %in% installed.packages()[, "Package"])]
  if (length(create.pkg))
    install.packages(create.pkg, dependencies = TRUE)
  sapply(pack, require, character.only = TRUE)
}


packages = c("tidyverse", "stringi", "lubridate", 
             "data.table", "pbapply", 
             "rgdal", "readxl", "webshot", "htmlwidgets", 
             "dataMeta", "sjPlot", "openxlsx", "surveytoolbox")


load_pkg(packages)

sample = haven::read_sav("20211025_muestra_turismo.sav")


directorio = sample %>% varl_tb()

field_work_data = readxl::read_excel("turismo_verificacion_TODAS.xlsx", range = "A2:AQ462")


glimpse(field_work_data)

table(field_work_data$estado_encuesta)

# Las categorías de estado de una encustas son los siguientes: 


# 1: En proceso de agendar
# 2: Agendada
# 3: Aplicada parcialmente
# 4: Aplicada y finalizada
# 5: Fuera de muestra
# 6: Rechazada
# 7: Inactiva
# 8: Sin contacto

# Las que para efectos del cálculo de los factores de expansión se recodificarán según la siguiente lista de correspencia


field_work_data %>% 
  group_by(estado_encuesta) %>% 
  count()


field_work_data %>% 
  group_by(estado_contacto) %>% 
  count()

# Proceso a clasificar la muestra en unidades elegibles y no elegibles 


sample_ = sample %>% 
  filter(Q29R1A1!="") %>% 
  mutate(Q1 = as.character(Q1))

field_work_data = field_work_data %>% 
  mutate(Folio = as.character(Folio),
         elegibilidad = ifelse(estado_encuesta %in% c("1: En proceso de agendar", "2: Agendada",
                                                      "6: Rechazada", "8: Sin contacto"), "Elegibilidad desconocida", 
                               ifelse(estado_encuesta %in% c("4: Aplicada y finalizada"), "Elegible", 
                               ifelse(estado_encuesta %in% c("5: Fuera de muestra", "7: Inactiva"), "No elegible", "Elegibilidad desconocida"))))




# Importadora y Comercializadora Karin Valck Toro E.I.R.L.
# Francisca Gatica Ferrada



ausente =   with(sample_, Q1) %in% with(field_work_data %>% filter(elegibilidad== "Elegible"), Folio)


no_coinciden = sample_$Q1[!ausente]

sample_ %>%
  filter(Q1 %in% no_coinciden) %>% 
  select(Q1,Q7, Q2)

# Casos en los que los folios no coinciden: 

# [1] "Alba Rosa Lopez Ferrada"                                               "Maria luisa San martin Bustamante"                                    
# [3] "ANA MARIA SEPULVEDA HIDD"                                              "Comercila & inversiones Ruiz y Delgado ltda"                          
# [5] "Centro turistico Sol de Quillon SPA"                                   "Ecoturismo tierra verde limitada"                                     
# [7] "Sociedad Vitivinícola viña lomas de Quillón"                           "RUTA LIMITE TURISMO LIMITADA"                                         
# [9] "Agricultor y productor de vinos María Loreto Alarcón Muffeler E.I.R.L" "Liliams Jara Castro"                                                  
# [11] "Cervecera y Comercial Toropaire SpA"

# Se procede a corregir los folios con errores de acuerdo a la siguiente rutina una vez cotejada la base de verificación con la muestra

sample_ = sample_ %>% 
  mutate(Q1 = ifelse(Q7 == "Alba Rosa Lopez Ferrada", "18139169142015294", 
              ifelse(Q7 == "Maria luisa San martin Bustamante", "18139169142015110", 
              ifelse(Q7 == "ANA MARIA SEPULVEDA HIDD", "11813925211471361", 
              ifelse(Q7 == "Comercila & inversiones Ruiz y Delgado ltda", "11845169142015453", 
              ifelse(Q7 == "Centro turistico Sol de Quillon SPA", "118139172191212324", 
              ifelse(Q7 == "Ecoturismo tierra verde limitada", "22139169142015397", 
              ifelse(Q7 == "Sociedad Vitivinícola viña lomas de Quillón", "22139172191212519", 
              ifelse(Q7 == "RUTA LIMITE TURISMO LIMITADA", "18224514914821366", 
              ifelse(Q7 == "Agricultor y productor de vinos María Loreto Alarcón Muffeler E.I.R.L", "22139181141721573", 
              ifelse(Q7 == "Liliams Jara Castro", "11845172191212341", 
              ifelse(Q7 == "Cervecera y Comercial Toropaire SpA", "22139172191212481", Q1)))))))))))) %>% 
  filter(!(Response_ID %in% c("128791127", "124703710")))


ausente = with(field_work_data %>% filter(elegibilidad== "Elegible"), Folio)  %in% with(sample_, Q1)


with(field_work_data %>% filter(elegibilidad == "Elegible"), Folio[!ausente])

# Finalmente la encuesta duplicada corresponde a el caso de Karin Valck

data = field_work_data %>%
  filter(Folio != "22139169142015484") %>% 
  rename(folio = Folio) %>% 
  left_join(sample_ %>% rename(folio = Q1) %>% filter(!duplicated(folio)))


data %>% 
  group_by(estado_encuesta) %>% 
  count()

data %>% 
  group_by(elegibilidad) %>% 
  count()




# s = Initial set of all sample units 
# s_IN = Set of units in s that are known to be ineligible 
# s_ER = Set of units that are eligible respondents 
# s_ENR = Set of units that are eligible nonrespondents 
# s_KN = Set of units whose elegibility is known 
# s_UNK = Set whose elegibility is unknown 

