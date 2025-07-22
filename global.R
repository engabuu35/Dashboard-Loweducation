# ================== Library ===================
library(readxl)
library(dplyr)
library(writexl)
library(leaflet)
library(sf)
library(htmltools)
library(stringr)
library(plm)
library(car) 
library(lmtest)
library(robustbase)
library(spdep)
library(openxlsx)
library(classInt)
library(shinyjs)
library(htmlwidgets)
library(webshot2)

# ================== Cek File ====================
if (!file.exists("Data/sovi_data.xlsx")) stop("File sovi_data.xlsx tidak ditemukan.")
if (!file.exists("Data/indonesia511.geojson")) stop("File indonesia511.geojson tidak ditemukan.")
if (!file.exists("Data/indonesia_simplified.geojson")) stop("File indonesia_simplified.geojson tidak ditemukan.")
if (!file.exists("Data/distance.csv")) stop("File distance.csv tidak ditemukan.")

# ================== Load Data ===================
data <- read_excel("Data/sovi_data.xlsx")
matrixdata <- read.csv("Data/distance.csv")

# ================== Load Geo Data ==============
geo <- st_read("Data/indonesia511.geojson", quiet = TRUE)
geo2 <- st_read("Data/indonesia_simplified.geojson", quiet = TRUE)

# ================== Persiapan Kolom Kode ===================
geo$DISTRICTCODE <- as.numeric(geo$kodeprkab)
geo2$DISTRICTCODE <- as.numeric(geo2$kodeprkab)
data$DISTRICTCODE <- as.numeric(data$DISTRICTCODE)

# ================== Gabung Nama Provinsi & Kabupaten ===================
data <- left_join(
  data,
  geo %>% select(DISTRICTCODE, nmkab, nmprov),
  by = "DISTRICTCODE"
)

# ================== Gabungkan ke Geo untuk Peta ===================
geo_map <- left_join(
  geo2,
  data %>% select(DISTRICTCODE, LOWEDU, POVERTY, NOELECTRIC, ELDERLY, FHEAD, FAMILYSIZE, RENTED, nmkab, nmprov),
  by = "DISTRICTCODE"
)

# ================== Perbaikan Nama Wilayah ===================
geo_map$nmkab <- geo_map$nmkab.y
geo_map$nmprov <- geo_map$nmprov.y

# ================== Definisikan Variabel Analisis: LOWEDU ===================
Y <- data$LOWEDU
X1 <- data$ELDERLY
X2 <- data$FHEAD
X3 <- data$FAMILYSIZE
X4 <- data$POVERTY
X5 <- data$NOELECTRIC
X6 <- data$RENTED
