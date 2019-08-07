rm(list = ls())

library(tidyverse)
library(janitor)
library(readxl)
library(openxlsx)


#---------------------------------------------
# GENIE IMPORT AND SET UP OF BASE DATAFRAME. NOTE THAT "FILTER" CODE LINE ARE ANALYSIS SPECIFIC 

df <- read.delim("~/R/datasets/Genie_Daily_82e86f5aaba344fa8e890f3ff91fa5ea.txt") %>% ####### NEEDS TO BE MANUALLY IMPORTED THROUGH MENU WITH EACH DATA GENIE UPDATE
  filter(Fiscal_Year == "2019") %>%
  rename(target = TARGETS, 
         site_id = orgUnitUID, 
         disag = standardizedDisaggregate, 
         agency = FundingAgency, 
         site = SiteName, 
         partner = mech_name, 
         Q1 = Qtr1, 
         Q2 = Qtr2, 
         Q3 = Qtr3, 
         Q4 = Qtr4)

df$partner <- as.character(df$partner)

df <- df %>%
  mutate(partner = 
           if_else(partner == "FADM HIV Treatment Scale-Up Program", "FADM",
                   if_else(partner == "Clinical Services System Strenghening (CHASS)", "CHASS",
                           if_else(partner == "Friends in Global Health", "FGH", partner)
                   )
           )
  ) 
