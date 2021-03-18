# Load lib
library(lubridate)
library(plyr)
library("xlsx")
library(openxlsx)
library(tidyverse)

# load dataset
dataset <- read.csv("D:\\# Mengajar\\Bu Nur\\Webinar\\dataset\\dataset mentah\\Statistik_Perkembangan_COVID19_Indonesia_New.csv", header = TRUE)

# seleksi fitur dan transformasi data
covid <- data.frame(
  tanggal = as.Date(dataset$Tanggal),
  positif_kumulatif = dataset$Jumlah_Kasus_Kumulatif,
  sembuh_kumulatif = dataset$Jumlah_Pasien_Sembuh,
  meninggal_kumulatif = dataset$Jumlah_Pasien_Meninggal,
  positif_harian = dataset$Jumlah_Kasus_Baru_per_Hari,
  sembuh_harian = dataset$Jumlah_Kasus_Sembuh_per_Hari,
  meninggal_harian = dataset$Jumlah_Kasus_Meninggal_per_Hari,
  perawatan_kumulatif = dataset$Jumlah_Pasien_Dalam_Perawatan,
  perawatan_harian = dataset$Jumlah_Kasus_Dirawat_per_Hari,
  persentase_sembuh = dataset$Persentase_Pasien_Sembuh,
  persentase_meninggal = dataset$Persentase_Pasien_Meninggal
)

# order data berdasarkan tanggal
covid <- covid[order(covid$tanggal),]

# potong data
covid <- covid[covid$tanggal <= "2020-10-19",]

# write.xlsx(covid, file = "D:\\# Mengajar\\Bu Nur\\Webinar\\dataset\\dataset_covid.xlsx", sheetName="data covid indonesia", row.names = FALSE, append=TRUE)

# ----------------------------------------------------------------------------------------

dataset2 <- read.csv("D:\\# Mengajar\\Bu Nur\\Webinar\\dataset\\dataset mentah\\Data_Harian_Kasus_per_Provinsi_COVID-19_Indonesia.csv", header = TRUE)

covid2 <- data.frame(
  provinsi = dataset2$Provinsi,
  jumlah_positif = dataset2$Kasus_Posi,
  jumlah_sembuh = dataset2$Kasus_Semb,
  jumlah_meninggal = dataset2$Kasus_Meni
)

# order data berdasarkan jumlah positif
covid2 <- covid2[order(-covid2$jumlah_positif),]

# potong data
covid2 <- covid2[covid2$provinsi != "Indonesia",]

# write.xlsx(covid2, file = "D:\\# Mengajar\\Bu Nur\\Webinar\\dataset\\dataset_covid.xlsx", sheetName="data covid provinsi", row.names = FALSE, append=TRUE)


# -------------------------------------------------------------------
# Create a blank workbook
OUT <- createWorkbook()

# Add some sheets to the workbook
addWorksheet(OUT, "data covid indonesia")
addWorksheet(OUT, "data covid provinsi")

# Write the data to the sheets
writeData(OUT, sheet = "data covid indonesia", x = covid)
writeData(OUT, sheet = "data covid provinsi", x = covid2)

# Reorder worksheets
worksheetOrder(OUT) <- c(2,1)

# Export the file
saveWorkbook(OUT, "D:\\# Mengajar\\Bu Nur\\Webinar\\dataset\\dataset_covid.xlsx")
