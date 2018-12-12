#Author : Navya Mote
#Date   : 10/22/2018
#Title  : Pre-processing for markets
#---------------------------------------------------------------------------------------------
#Libraries
library(expss)
library(dplyr)
library(data.table)
library(openxlsx)
library(knitr)
library(DT)
#To set the working directory
setwd("C:/Users/Navya/Desktop/PGI")
`CBRE-EA`<-data.frame()
# ----------------------------------------
# To get the dataframes from EBA excel
library(readxl)    
read_excel_allsheets <- function(filename, tibble = FALSE) {
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  if(!tibble) x <- lapply(x, as.data.frame)
  names(x) <- sheets
  x
}
mysheets <- read_excel_allsheets("2018Q2 EBA data.xlsx")
EA_Off_Annual<-mysheets$EA_Off_Annual
EA_Ind_Annual<-mysheets$EA_Ind_Annual
EA_Off_Qtrly<-mysheets$EA_Off_Qtrly
EA_Ind_Qtrly<-mysheets$EA_Ind_Qtrly
`RCA Apartment`<-mysheets$`RCA Apartment`
`RCA Industrial`<-mysheets$`RCA Industrial`
`RCA Office`<-mysheets$`RCA Office`
`RCA Retail`<-mysheets$`RCA Retail`
moodys_data<-mysheets$moodys_data
Rent_Fcst<-mysheets$Rent_Fcst
PPR_All<-mysheets$PPR_All
`Reis Apartment`<-mysheets$`Reis Apartment`
`Reis Retail`<-mysheets$`Reis Retail`
`UC Ind`<-mysheets$`UC Ind`
`UC Apt`<-mysheets$`UC Apt`
`UC Off`<-mysheets$`UC Off`
`UC Ret`<-mysheets$`UC Ret`
`Equil_rate Off`<-mysheets$`Equil_rate Off`
`Equil_rate Ind`<-mysheets$`Equil_rate Ind`
`Equil_rate Apt`<-mysheets$`Equil_rate Apt`
`Equil_rate Ret`<-mysheets$`Equil_rate Ret`
educdata<-mysheets$educdata
emplyrdata<-mysheets$emplyrdata
# To get the cities
mysheets1 <- read_excel_allsheets("EBA Dashboard Target Markets.xlsx")
`CBRE-EA`<-data.frame("Cities"= mysheets1$Sheet1$`CBRE-EA`,stringsAsFactors=FALSE)
`CBRE-EA`$Cities[`CBRE-EA`$Cities=="Sum of Markets"]<-"USA"
# --------------------------------------------------------
# City change at PPR
PPR_All$`Geography Name`[PPR_All$`Geography Name`=="Virginia Beach"]<-"Norfolk"
PPR_All$`Geography Name`[PPR_All$`Geography Name`=="Northern New Jersey"]<-"Newark"
PPR_All$`Geography Name`[PPR_All$`Geography Name`=="Oakland-East Bay"]<-"Oakland"
PPR_All$`Geography Name`[PPR_All$`Geography Name`=="Palm Beach"]<-"West Palm Beach"
PPR_All$`Geography Name`[PPR_All$`Geography Name`=="Raleigh-Durham"]<-"Raleigh"
PPR_All$`Geography Name`[PPR_All$`Geography Name`=="San Bernardino/Riverside"]<-"Riverside"
PPR_All$`Geography Name`[PPR_All$`Geography Name`=="Tampa-St. Petersburg"]<-"Tampa"
PPR_All$`Geography Name`[PPR_All$`Geography Name`=="Ventura County"]<-"Ventura"
# ----------------------------------------------------------
# City change at REIS
`Reis Apartment`$Metro[`Reis Apartment`$Metro == "New York Metro"]<-"New York"
`Reis Retail`$Metro[`Reis Retail`$Metro == "New York Metro"]<-"New York"
`Reis Apartment`$Metro[`Reis Apartment`$Metro == "Norfolk/Hampton Roads"]<-"Norfolk"
`Reis Retail`$Metro[`Reis Retail`$Metro == "Norfolk/Hampton Roads"]<-"Norfolk"
`Reis Apartment`$Metro[`Reis Apartment`$Metro == "Northern New Jersey"]<-"Newark"
`Reis Retail`$Metro[`Reis Retail`$Metro == "Northern New Jersey"]<-"Newark"
`Reis Apartment`$Metro[`Reis Apartment`$Metro == "Oakland-East Bay"]<-"Oakland"
`Reis Retail`$Metro[`Reis Retail`$Metro == "Oakland-East Bay"]<-"Oakland"
`Reis Apartment`$Metro[`Reis Apartment`$Metro == "Palm Beach"]<-"West Palm Beach"
`Reis Retail`$Metro[`Reis Retail`$Metro == "Palm Beach"]<-"West Palm Beach"
`Reis Apartment`$Metro[`Reis Apartment`$Metro == "Raleigh-Durham"]<-"Raleigh"
`Reis Retail`$Metro[`Reis Retail`$Metro == "Raleigh-Durham"]<-"Raleigh"
`Reis Apartment`$Metro[`Reis Apartment`$Metro == "San Bernardino/Riverside"]<-"Riverside"
`Reis Retail`$Metro[`Reis Retail`$Metro == "San Bernardino/Riverside"]<-"Riverside"
`Reis Apartment`$Metro[`Reis Apartment`$Metro == "Tampa-St. Petersburg"]<-"Tampa"
`Reis Retail`$Metro[`Reis Retail`$Metro == "Tampa-St. Petersburg"]<-"Tampa"
`Reis Apartment`$Metro[`Reis Apartment`$Metro == "Ventura County"]<-"Ventura"
`Reis Retail`$Metro[`Reis Retail`$Metro == "Ventura County"]<-"Ventura"
`Reis Apartment`$Metro[`Reis Apartment`$Metro == "District of Columbia"]<-"Washington, DC"
`Reis Retail`$Metro[`Reis Retail`$Metro == "District of Columbia"]<-"Washington, DC"
# --------------------------------------------------------------
# City change at Moody's
moodys_data$Population[moodys_data$Population =="Norfolk/Hampton Roads"]<-"Norfolk"
moodys_data$Population[moodys_data$Population =="Northern New Jersey"]<-"Newark"
moodys_data$Population[moodys_data$Population =="Oakland-East Bay"]<-"Oakland"
moodys_data$Population[moodys_data$Population =="Palm Beach"]<-"West Palm Beach"
moodys_data$Population[moodys_data$Population =="Raleigh-Durham"]<-"Raleigh"
moodys_data$Population[moodys_data$Population =="San Bernardino/Riverside"]<-"Riverside"
moodys_data$Population[moodys_data$Population =="Tampa-St. Petersburg"]<-"Tampa"
moodys_data$Population[moodys_data$Population =="Ventura County"]<-"Ventura"
moodys_data$Population[moodys_data$Population =="Washington D.C. "]<-"Washington, DC"
# -----------------------------------------------------------------
# To insert a row for population
# moodys_data<-rbind(c("Population",NA,NA,NA,NA,NA,NA,NA,NA,
#                      NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,
#                      NA,NA,NA,NA), moodys_data)
# -----------------------------------------------------------------
# City change at RCA
`RCA Apartment`$Market[`RCA Apartment`$Market == "Broward"]<-"Fort Lauderdale"
`RCA Industrial`$Market[`RCA Industrial`$Market == "Broward"]<-"Fort Lauderdale"
`RCA Office`$Market[`RCA Office`$Market == "Broward"]<-"Fort Lauderdale"
`RCA Retail`$Market[`RCA Retail`$Market == "Broward"]<-"Fort Lauderdale"

`RCA Apartment`$Market[`RCA Apartment`$Market == "Miami/Dade Co"]<-"Miami"
`RCA Industrial`$Market[`RCA Industrial`$Market == "Miami/Dade Co"]<-"Miami"
`RCA Office`$Market[`RCA Office`$Market == "Miami/Dade Co"]<-"Miami"
`RCA Retail`$Market[`RCA Retail`$Market == "Miami/Dade Co"]<-"Miami"

`RCA Apartment`$Market[`RCA Apartment`$Market == "NYC Boroughs"]<-"New York"
`RCA Industrial`$Market[`RCA Industrial`$Market == "NYC Boroughs"]<-"New York"
`RCA Office`$Market[`RCA Office`$Market == "NYC Boroughs"]<-"New York"
`RCA Retail`$Market[`RCA Retail`$Market == "NYC Boroughs"]<-"New York"

`RCA Apartment`$Market[`RCA Apartment`$Market == "No NJ"]<-"Newark"
`RCA Industrial`$Market[`RCA Industrial`$Market == "No NJ"]<-"Newark"
`RCA Office`$Market[`RCA Office`$Market == "No NJ"]<-"Newark"
`RCA Retail`$Market[`RCA Retail`$Market == "No NJ"]<-"Newark"

`RCA Apartment`$Market[`RCA Apartment`$Market == "East Bay"]<-"Oakland"
`RCA Industrial`$Market[`RCA Industrial`$Market == "East Bay"]<-"Oakland"
`RCA Office`$Market[`RCA Office`$Market == "East Bay"]<-"Oakland"
`RCA Retail`$Market[`RCA Retail`$Market == "East Bay"]<-"Oakland"

`RCA Apartment`$Market[`RCA Apartment`$Market == "Orange Co"]<-"Orange County"
`RCA Industrial`$Market[`RCA Industrial`$Market == "Orange Co"]<-"Orange County"
`RCA Office`$Market[`RCA Office`$Market == "Orange Co"]<-"Orange County"
`RCA Retail`$Market[`RCA Retail`$Market == "Orange Co"]<-"Orange County"

`RCA Apartment`$Market[`RCA Apartment`$Market == "Palm Beach Co"]<-"West Palm Beach"
`RCA Industrial`$Market[`RCA Industrial`$Market == "Palm Beach Co"]<-"West Palm Beach"
`RCA Office`$Market[`RCA Office`$Market == "Palm Beach Co"]<-"West Palm Beach"
`RCA Retail`$Market[`RCA Retail`$Market == "Palm Beach Co"]<-"West Palm Beach"

`RCA Apartment`$Market[`RCA Apartment`$Market == "Raleigh/Durham"]<-"Raleigh"
`RCA Industrial`$Market[`RCA Industrial`$Market == "Raleigh/Durham"]<-"Raleigh"
`RCA Office`$Market[`RCA Office`$Market == "Raleigh/Durham"]<-"Raleigh"
`RCA Retail`$Market[`RCA Retail`$Market == "Raleigh/Durham"]<-"Raleigh"

`RCA Apartment`$Market[`RCA Apartment`$Market == "Inland Empire"]<-"Riverside"
`RCA Industrial`$Market[`RCA Industrial`$Market == "Inland Empire"]<-"Riverside"
`RCA Office`$Market[`RCA Office`$Market == "Inland Empire"]<-"Riverside"
`RCA Retail`$Market[`RCA Retail`$Market == "Inland Empire"]<-"Riverside"

`RCA Apartment`$Market[`RCA Apartment`$Market == "St Louis"]<-"St. Louis"
`RCA Industrial`$Market[`RCA Industrial`$Market == "St Louis"]<-"St. Louis"
`RCA Office`$Market[`RCA Office`$Market == "St Louis"]<-"St. Louis"
`RCA Retail`$Market[`RCA Retail`$Market == "St Louis"]<-"St. Louis"

`RCA Apartment`$Market[`RCA Apartment`$Market == "Tulsa"]<-"Tucson"
`RCA Industrial`$Market[`RCA Industrial`$Market == "Tulsa"]<-"Tucson"
`RCA Office`$Market[`RCA Office`$Market == "Tulsa"]<-"Tucson"
`RCA Retail`$Market[`RCA Retail`$Market == "Tulsa"]<-"Tucson"

`RCA Apartment`$Market[`RCA Apartment`$Market == "Ventura Co"]<-"Ventura"
`RCA Industrial`$Market[`RCA Industrial`$Market == "Ventura Co"]<-"Ventura"
`RCA Office`$Market[`RCA Office`$Market == "Ventura Co"]<-"Ventura"
`RCA Retail`$Market[`RCA Retail`$Market == "Ventura Co"]<-"Ventura"

`RCA Apartment`$Market[`RCA Apartment`$Market == "DC"]<-"Washington, DC"
`RCA Industrial`$Market[`RCA Industrial`$Market == "DC"]<-"Washington, DC"
`RCA Office`$Market[`RCA Office`$Market == "DC"]<-"Washington, DC"
`RCA Retail`$Market[`RCA Retail`$Market == "DC"]<-"Washington, DC"
# ----------------------------------------------------------------
# Rent
Rent_Fcst$MarketName[Rent_Fcst$MarketName == "Northern New Jersey"] <- "Norfolk"
# ---------------------------------------------------------------
# To place the tables into excel file
l<-list("CBRE-EA" = `CBRE-EA`,"EA_Off_Annual" = `EA_Off_Annual`, "EA_Ind_Annual" = `EA_Ind_Annual`,
        "EA_Off_Qtrly" = `EA_Off_Qtrly`, "EA_Ind_Qtrly" = `EA_Ind_Qtrly`,
        "RCA Apartment" = `RCA Apartment`, "RCA Industrial" = `RCA Industrial`,
        "RCA Office" = `RCA Office`, "RCA Retail" = `RCA Retail`,
        "moodys_data" = `moodys_data`, "Rent_Fcst" = `Rent_Fcst`,
        "PPR_All" = `PPR_All`, "Reis Apartment" = `Reis Apartment`,
        "Reis Retail" = `Reis Retail`, "UC Ind" = `UC Ind`,
        "UC Apt" = `UC Apt`, "UC Off" = `UC Off`, "UC Ret" = `UC Ret`,
        "Equil_rate Off" = `Equil_rate Off`, "Equil_rate Ind" = `Equil_rate Ind`,
        "Equil_rate Apt" = `Equil_rate Apt`, "Equil_rate Ret" = `Equil_rate Ret`,
        "educdata" = `educdata`, "emplyrdata" = `emplyrdata`)
write.xlsx(l,"2018Q2 EBA data.xlsx")







