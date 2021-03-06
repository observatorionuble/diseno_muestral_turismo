# *-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-**-*
#==  Programa para el C�lculo del tama�o muestral Encuesta Sector Turismo 
# *-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-*-**-*-**-*
#
#
# H�ctor Garrido Henr�quez
# Analista Cuantitativo. Observatorio Laboral �uble
# Universidad del B�o-B�o
# Avenida Andr�s Bello 720, Casilla 447, Chill�n
# Tel�fono: +56-942353973
# http://www.observatoriolaboralnuble.cl




rm(list = ls())

## Se crea una func�n para cargar y/o instalar los paquetes necesarios para la rutina

load_pkg <- function(pack){
  create.pkg <- pack[!(pack %in% installed.packages()[, "Package"])]
  if (length(create.pkg))
    install.packages(create.pkg, dependencies = TRUE)
  sapply(pack, require, character.only = TRUE)
}


packages = c("tidyverse", "stringi", "lubridate", 
             "data.table", "pbapply", 
             "rgdal", "readxl", "webshot", "htmlwidgets", 
             "dataMeta", "sjPlot", "openxlsx")


load_pkg(packages)

# Se determina la ruta del directorio de trabajo en que se encuentran los datos y se procede a cargarlos en el Environment

path =  "G:/Mi unidad/OLR �uble - Observatorio laboral de �uble/An�lisis Cuantitativo/"

prestadores = read_excel(paste0(path, "SERVICIOS REGISTRADOS 04_03_21_OBL.xlsx"), 
                         sheet = 1, col_names = TRUE) 


# Se procede a depurar la base de datos para que pueda servir de marco muestral para el estudio, los pasos aqu� ejecutados est�n 
# descritos a grandes rasgos en el documento metodol�gico. 

data = prestadores %>% 
  mutate(Comuna = ifelse(Comuna == "Chillan Viejo", "ChViejo", Comuna)) %>% 
  filter(`Tipo de Servicio` %in% c("Alojamiento tur�stico", "Restaurantes y similares",
                                   "Gu�as de Turismo", "Tour operador",
                                   "Agencia de viajes", "Turismo aventura",
                                   "Arriendo de Veh�culos", "Servicios deportivos",
                                   "Servicios de esparcimiento", "Servicios Culturales")) %>% 
  mutate(codigo_tres = stri_sub(`C�digo Actividad econ�mica principal`, from = 1, to = 3), 
         `Tipo de Servicio` = ifelse(`Tipo de Servicio` %in% c("Gu�as de Turismo", "Tour operador",
                                                               "Agencia de viajes", "Turismo aventura", 
                                                               "Arriendo de Veh�culos", "Servicios deportivos",
                                                               "Servicios de esparcimiento", "Servicios Culturales"),
                                     "Varios", `Tipo de Servicio`)) %>% 
  group_by(`Tipo de Servicio`, `Rut Empresa`) %>% #(1)
  mutate(clases = paste0(unique(`Clase`), collapse = " | "),
         n_ = n(), 
         ord = 1:n()) %>% 
  filter(ord == 1)  %>% 
  arrange(`Tipo de Servicio`) %>% 
  group_by(`Rut Empresa`) %>% #(2) 
  mutate(`Tipo de Servicio` = stri_sub(`Tipo de Servicio`,1,1), 
         servicios = paste0(unique(`Tipo de Servicio`), collapse = " | "),
         m_ = 1:n()) %>% 
  mutate(servicios = ifelse(`Rut Empresa` %in% c("5839540-4", "5429064-0", "76870694-8", "9300626-7"), "R", 
                     ifelse(`Rut Empresa` %in% c("50837020-2", "76225190-6", "6470335-8", "6635027-4", 
                                                 "14746889-k", "76138624-7", "7145642-0", "5254436-k",
                                                 "76301254-9", "4887565-3", "76580737-9","9237007-0"), "A",
                            servicios))) %>% 
  filter(m_ == 1) %>% 
  mutate(complicado = ifelse(grepl("|", servicios, fixed = TRUE),1,0), 
         `N�mero de trabajadores dependientes informados` = ifelse(`N�mero de trabajadores dependientes informados HOMBRES`>0 |
                                                                     `N�mero de trabajadores dependientes informados MUJERES`>0, 
                                                                   rowSums(across(contains("N�mero de trabajadores")), na.rm = TRUE), 
                                                                   NA), 
         tramo_tama�o = ifelse(is.na(`N�mero de trabajadores dependientes informados`),"desconocido", 
                        ifelse(`N�mero de trabajadores dependientes informados` == 4340, "desconocido", 
                        ifelse(`N�mero de trabajadores dependientes informados` %in% 0:9, "Micro",
                        ifelse(`N�mero de trabajadores dependientes informados` %in% 10:49, "Peque�a", 
                        ifelse(`N�mero de trabajadores dependientes informados` %in% 50:199, "Mediana", 
                        ifelse(`N�mero de trabajadores dependientes informados`>=200, "Grande","desconocido"))))))) %>% 
  mutate(`Categor�a Empresa` = ifelse(is.na(`Categor�a Empresa`),"Sin categor�a", `Categor�a Empresa`))



# Se procede a calcular el tama�o muestral teniendo en cuenta los siguientes par�metros 

error = .30 #M�ximo error relativo a tolerar en cada estrato de actividad econ�mica 
prop = .5 # Varianza m�xima 
t_ = 2 # valor Z aproximado suponiendo un 95% de confianza


# Se procede a calcular el tama�o muestral objetivo y con sobremuestra de acuerdo a un muestreo estratificado

no_a = 0.19
no_r = 0.24
no_v = 0.29

# Tasas de logro efectiva. Obtenidas con el trabajo de campo al 75% (24/09/2021)

no_a_e = 0.69
no_r_e = 0.57
no_v_e = 0.67


sample_size = data %>% 
  mutate(`Tipo de Servicio` = fct_explicit_na(`Tipo de Servicio`, na_level = "(Sin Informaci�n)")) %>% 
  #filter(Vigencia != "No Vigente") %>% 
  group_by(servicios) %>% 
  mutate(n = n()) %>% 
  mutate(muestra = ((n*(prop-prop*prop+(((error*prop)/t_))^2))/(((((error*prop))/t_)^2)*n+prop-prop*prop)) %>%  ceiling()) %>% 
  group_by(servicios) %>% 
  summarise(total = n[1], 
            muestra = muestra[1]) %>% 
  mutate(muestra = ifelse(total<=6, total, muestra)) %>% 
  arrange(desc(muestra)) %>% 
  mutate(var = ((total-muestra)/total)*((prop*(1-prop))/(muestra-1)), 
         error_ = sqrt(var)*4, 
         sobremuestra = ifelse(`servicios` == "A", muestra*(1/(1-no_a)), 
                        ifelse(`servicios` == "R", muestra*(1/(1-no_r)), 
                        ifelse(`servicios` == "V", muestra*(1/(1-no_v)), muestra))), 
         sobremuestra = sobremuestra %>% round(0)) %>% 
  mutate(reemplazo = ifelse(`servicios` == "A", sobremuestra*((1/(1-no_a_e))-1), 
                            ifelse(`servicios` == "R", sobremuestra*((1/(1-no_r_e))-1), 
                                   ifelse(`servicios` == "V", sobremuestra*((1/(1-no_v_e))-1), 0))), 
         reemplazo = round(reemplazo,0))


# Se procede a distribuir la muestraa por comuna, actividad econ�mica y tama�o de la empresa 
# (El proceso de redondeo hace crecer levemente el tama�o estimado de la muestra)

# Se definieron como empresas de dif�cil acceso a aquellas con una probabilidad de selecci�n
# en el marco menores a 1% de acuerdo al cruce de sector econ�mico, tama�o y comuna

# Distribuci�n del tama�o muestral por afijaci�n proporcional a los estratos 

dist_sample = data %>% 
  left_join(sample_size) %>% 
  group_by(servicios, tramo_tama�o, Comuna) %>% 
  summarise(frecuencia = n(), 
            sobremuestra = sobremuestra[1], 
            muestra = muestra[1], 
            reemplazo = reemplazo[1]) %>% 
  group_by(servicios) %>% 
  mutate(freq_s = sum(frecuencia), 
         pct = frecuencia/freq_s, 
         muestra_efec = round(pct*sobremuestra), 
         muestra = round(pct*muestra)) %>% 
  mutate(pct_pqn = frecuencia/25, # 25 es el total de empresas peque�as 
         muestra = ifelse(tramo_tama�o == "Peque�a", round(15*pct_pqn,0), # 15 es el tama�o necesario para obtener representatividad de empresas peque�as   
                   ifelse(tramo_tama�o == "Mediana", 1, muestra)), 
         muestra_efec = ifelse(tramo_tama�o == "Peque�a", round(15*pct_pqn,0), 
                               ifelse(tramo_tama�o == "Mediana", 1, muestra_efec)), 
         replace = round(pct*reemplazo)) %>% 
  arrange(across(c("servicios", "tramo_tama�o", "Comuna"))) %>% 
  ungroup() %>% 
  mutate(estratos = gsub(" ","", paste0(servicios, stri_sub(tramo_tama�o,1,2), stri_sub(Comuna,1,5))) %>%
           stri_trans_tolower %>% stri_trans_general("Latin-ASCII"),
         estrato = map(estratos, ~paste0(sapply(strsplit(gsub(" |\\|", "",.x), "")[[1]], function(x) grep(x, letters)), collapse = "")), 
         estrato = as.character(estrato)) %>% 
  group_by(servicios) %>% 
  mutate(frec_el = ifelse(muestra_efec==0, 0, frecuencia), 
         tot_el = sum(frec_el), #total empresas accesibles 
         tot = sum(frecuencia), 
         r_el = tot/tot_el)  # raz�n de ajuste por empresas de d�ficil acceso 




# muestra_efec : muestra original 


# Se fija una semilla para la generaci�n de una secuencia pseudo-aleatoria que permita que la muestra aqu� tomada sea replicable 

set.seed(12345)

# Se crea una variable estrato y folio para cada empresa basada en la actividad econ�mica, el tama�o, la comuna y el Rut

# sum(sapply(strsplit(gsub(" |\\|", "",muestra_dist_$estratos[50]),"")[[1]], function(x) grep(x, letters)))


data_ = data %>% 
  mutate(estratos = gsub(" ","", paste0(servicios, stri_sub(tramo_tama�o,1,2), stri_sub(Comuna,1,5))) %>%
           stri_trans_tolower %>% stri_trans_general("Latin-ASCII"),
         estrato = map(estratos, ~paste0(sapply(strsplit(gsub(" |\\|", "",.x), "")[[1]], function(x) grep(x, letters)), collapse = "")), 
         estrato = as.character(estrato)) %>% 
  arrange(estrato) %>% 
  left_join(dist_sample) %>% 
  ungroup() %>% 
  arrange(`Rut Empresa`) %>% 
  mutate(rep = 1:n()) %>% 
  mutate(folio = paste0(estrato, rep)) %>% 
  dplyr::select(folio, estrato, everything(), -`N�mero de trabajadores dependientes informados HOMBRES`, 
                -`N�mero de trabajadores dependientes informados MUJERES`,
                -`Fecha inicio de actividades ante el Servicio de Impuestos Internos`, 
                -`Vigencia`, 
                -`codigo_tres`, 
                -`n_`, 
                -ord, 
                -m_,
                -complicado, 
                -`N�mero de trabajadores dependientes informados`,
                -frecuencia, 
                -sobremuestra, 
                -muestra, 
                -freq_s,
                -pct,
                -muestra_efec,
                -pct_pqn,
                -estratos,
                -`Tama�o empresa`, 
                -rep, 
                -reemplazo, 
                -replace, 
                -frec_el, -tot_el, -tot, -r_el, -muestra)


estratos_ = dist_sample$estrato

n_ = dist_sample$muestra_efec

muestra_objetivo = pblapply(1:length(estratos_), function(x) data_ %>% 
                              ungroup() %>% 
                              filter(estrato == estratos_[x]) %>%   
                              slice_sample(n = n_[x])) %>% 
  bind_rows() %>% 
  mutate(muestra = "Muestra original")


# Datos para la muestra de reemplazo 
m_ = dist_sample$replace

muestra_reemplazo = pblapply(1:length(estratos_), function(x) data_ %>% 
  filter(!(folio %in% muestra_objetivo$folio)) %>% 
    ungroup() %>% 
    filter(estrato == estratos_[x]) %>% 
    slice_sample(n = m_[x])) %>% 
  bind_rows() %>% 
  mutate(muestra = "Muestra de reemplazo")



# Se verifica el tama�o de la muestra seg�n actividad econ�mica y tama�o de la empresa
(muestra_tam = muestra_dist %>% 
    group_by(servicios, tramo_tama�o) %>% 
    summarise(sobremuestra = sum(muestra_efec), 
              total = sum(frecuencia), 
              muestra = sum(muestra)))


# Se verifica el tama�o de la muestra por actividad econ�mica y comuna 

(muestra_com = muestra_dist %>% 
    group_by(servicios, Comuna) %>% 
    summarise(sobremuestra = sum(muestra_efec), 
              total = sum(frecuencia), 
              muestra = sum(muestra)) %>% 
    reshape2::dcast(Comuna~servicios, value.var = "sobremuestra"))


# Se verifica el tama�o de la muestra por actividad econ�mica 
muestra_def = muestra_dist %>% 
  group_by(servicios) %>% 
  summarise(muestra = sum(muestra),
            sobremuestra = sum(muestra_efec), 
            total = sum(frecuencia), 
            var = ((total-muestra)/total)*((prop*(1-prop))/(muestra-1)), 
            error_ = sqrt(var)*4)




# Se verifica el tama�o de muestra definitivo por tramo de tama�o de la empresa seg�n el n�mero de trabajadores

muestra_def_2 = muestra_dist %>% 
  group_by(tramo_tama�o) %>% 
  summarise(muestra = sum(muestra),
            sobremuestra = sum(muestra_efec), 
            total = sum(frecuencia), 
            var = ((total-muestra)/total)*((prop*(1-prop))/(muestra-1)), 
            error_ = sqrt(var)*4)

# Se procede a impprimir la tabla del tama�o definitivo en un documento word.

tab_df(muestra_def_2,
       show.rownames = FALSE,
       file = paste0(path, "20201123 tama�o muestral por tama�o de la empresa.doc"),
       footnote = c("Fuente: Base de datos de prestadores tur�sticos de Sernatur"),
       title = "Figura 1. Distribuci�n del Marco Muestral seg�n Actividad econ�mica",
       col.header = c("Tipo de Servicio", "Muestra", "Sobremuestra", "Total", "varianza", "error relativo"),
       encoding = "Windows-1252")

# Se procede a imprimir la tabla con los tama�os muestrales definitivos en un documento word. 
tab_df(muestra_def,
       show.rownames = FALSE,
       file = paste0(path, "20201123 tama�o muestral por actividad.doc"),
       footnote = c("Fuente: Base de datos de prestadores tur�sticos de Sernatur"),
       title = "Figura 1. Distribuci�n del Marco Muestral seg�n Actividad econ�mica",
       col.header = c("Tipo de Servicio", "Muestra", "Sobremuestra", "Total", "varianza", "error relativo"),
       encoding = "Windows-1252")



wb = createWorkbook()

addWorksheet(wb, sheetName = "Marco Muestral", gridLines = FALSE)

writeData(wb, sheet = "Marco Muestral",
          x = data_,
          startRow = 3, startCol = 2, 
          colNames = TRUE, rowNames = FALSE,
          keepNA = FALSE, 
          withFilter = FALSE)

addWorksheet(wb, sheetName = "Muestra", gridLines = FALSE)

writeData(wb, sheet = "Muestra",
          x = muestra_objetivo,
          startRow = 3, startCol = 2, 
          colNames = TRUE, rowNames = FALSE,
          keepNA = FALSE, 
          withFilter = FALSE)

addWorksheet(wb, sheetName = "Muestra", gridLines = FALSE)
writeData(wb, sheet = "Muestra reemplazo",
          x = muestra_reemplazo,
          startRow = 3, startCol = 2, 
          colNames = TRUE, rowNames = FALSE,
          keepNA = FALSE, 
          withFilter = FALSE)


saveWorkbook(wb, file = paste0(path, "20210924 estudio SENCE SERNATUR.xlsx"), overwrite = TRUE)


# 
# actividad_2 = data %>% 
#   group_by(Comuna) %>% count() %>% arrange(desc(n))
# 
# 
# tab_df(actividad_2,
#        show.rownames = FALSE,
#        file = paste0(path, "20201123 Cuadro prestadores seg�n comuna.doc"),
#        footnote = c("Fuente: Base de datos de prestadores tur�sticos de Sernatur"),
#        title = "Figura 1. Distribuci�n del Marco Muestral seg�n Tipo de Servicio y Vigencia",
#        col.header = c("Tipo de Servicio", "No Vigente", "Provisorio", "Vigente"),
#        encoding = "Windows-1252")
# 
# 
# tama�o = data %>% 
#   group_by(tramo_tama�o) %>% count() %>% arrange(desc(n))
# 
# 
# tab_df(tama�o,
#        show.rownames = FALSE,
#        file = paste0(path, "20201123 Cuadro prestadores seg�n tama�o.doc"),
#        footnote = c("Fuente: Base de datos de prestadores tur�sticos de Sernatur"),
#        title = "Figura 1. Distribuci�n del Marco Muestral seg�n Tipo de Servicio y Vigencia",
#        col.header = c("Tramo tama�o", "Empresas"),
#        encoding = "Windows-1252")
# 
# 
# 
# 
# tabla_actividades = prestadores %>% 
#   mutate(Comuna = ifelse(Comuna == "Chillan Viejo", "ChViejo", Comuna)) %>% 
#   filter(`Tipo de Servicio` %in% c("Alojamiento tur�stico", "Restaurantes y similares",
#                                    "Gu�as de Turismo", "Tour operador",
#                                    "Agencia de viajes", "Turismo aventura",
#                                    "Arriendo de Veh�culos", "Servicios deportivos",
#                                    "Servicios de esparcimiento", "Servicios Culturales")) %>% 
#   mutate(codigo_tres = stri_sub(`C�digo Actividad econ�mica principal`, from = 1, to = 3), 
#          `Tipo de Servicio` = ifelse(`Tipo de Servicio` %in% c("Gu�as de Turismo", "Tour operador",
#                                                                "Agencia de viajes", "Turismo aventura", 
#                                                                "Arriendo de Veh�culos", "Servicios deportivos",
#                                                                "Servicios de esparcimiento", "Servicios Culturales"),
#                                      "Varios", `Tipo de Servicio`)) %>% 
#   group_by(`Tipo de Servicio`) %>% 
#   summarise(actividades = paste0(unique(Clase), collapse = ", "))
# 
# 
# tab_df(tabla_actividades, 
#        show.rownames = FALSE, 
#        file = paste0(path, "20201123 Cuadro prestadores seg�n actividad.doc"),
#        footnote = c("Fuente: Base de datos de prestadores tur�sticos de Sernatur"),
#        title = "Figura 1. Distribuci�n del Marco Muestral seg�n actividades",
#        encoding = "Windows-1252")
# 
# 
# 
