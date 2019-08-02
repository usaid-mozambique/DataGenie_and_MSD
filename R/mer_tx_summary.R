rm(list = ls())

library(tidyverse)
library(janitor)
library(readxl)
library(openxlsx)


#---------------------------------------------
# GENIE IMPORT AND SET UP OF BASE DATAFRAME. NOTE THAT "FILTER" AND "SELECT" CODE LINES ARE ANALYSIS SPECIFIC 

df <- read.delim("~/R/datasets/Genie_Daily_82e86f5aaba344fa8e890f3ff91fa5ea.txt") %>% ####### NEEDS TO BE MANUALLY IMPORTED THROUGH MENU WITH EACH DATA GENIE UPDATE
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


#---------------------------------------------
## PREPARE TX_CURR DATASET
tx_curr_base <- df %>%
  filter(indicator == "TX_CURR" & Fiscal_Year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target)


#---------------------------------------------
##AGENCY LEVEL TX_CURR PERFORMANCE
tx_curr_agency <- tx_curr_base %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()


#---------------------------------------------
##PROVINCE LEVEL TX_CURR PERFORMANCE
tx_curr_snu1 <- tx_curr_base %>%
  group_by(agency, partner, SNU1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()


#---------------------------------------------
##DISTRICT LEVEL TX_CURR PERFORMANCE
tx_curr_psnu <- tx_curr_base %>%
  group_by(agency, partner, SNU1, PSNU) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()


#---------------------------------------------
##SITE LEVEL TX_CURR PERFORMANCE
tx_curr_site <- tx_curr_base %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()

rm(tx_curr_base)


#---------------------------------------------
##AGENCY LEVEL TX_NEW PERFORMANCE

tx_new_agency <- df %>%
  filter(indicator == "TX_NEW" & Fiscal_Year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q_vs_Target = (Q3)/(target / 4)) %>% ####### QUARTER SPECIFIC CODE
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()


#---------------------------------------------
##PROVINCE LEVEL TX_NEW PERFORMANCE
tx_new_snu1 <- df %>%
  filter(indicator == "TX_NEW" & Fiscal_Year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(agency, partner, SNU1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q_vs_Target = (Q3)/(target / 4)) %>% ####### QUARTER SPECIFIC CODE
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()


#---------------------------------------------
##DISTRICT LEVEL TX_NEW PERFORMANCE
tx_new_psnu <- df %>%
  filter(indicator == "TX_NEW" & Fiscal_Year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(agency, partner, SNU1, PSNU) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q_vs_Target = (Q3)/(target / 4)) %>% ####### QUARTER SPECIFIC CODE
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()


#---------------------------------------------
##SITE LEVEL TX_NEW PERFORMANCE
tx_new_site <- df %>%
  filter(indicator == "TX_NEW" & Fiscal_Year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q_vs_Target = (Q3)/(target / 4)) %>% ####### QUARTER SPECIFIC CODE
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()


#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_curr OUTPUT

wb <- createWorkbook()
addWorksheet(wb, "tx_curr_agency",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeData(wb, "tx_curr_agency", tx_curr_agency)
addWorksheet(wb, "tx_curr_snu1",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeData(wb, "tx_curr_snu1", tx_curr_snu1)
addWorksheet(wb, "tx_curr_psnu",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeData(wb, "tx_curr_psnu", tx_curr_psnu)
addWorksheet(wb, "tx_curr_site",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeData(wb, "tx_curr_site", tx_curr_site)
addWorksheet(wb, "tx_new_agency",
             tabColour = "#FFCC00",
             gridLines = FALSE)
writeData(wb, "tx_new_agency", tx_new_agency)
addWorksheet(wb, "tx_new_snu1",
             tabColour = "#FFCC00",
             gridLines = FALSE)
writeData(wb, "tx_new_snu1", tx_new_snu1)
addWorksheet(wb, "tx_new_psnu",
             tabColour = "#FFCC00",
             gridLines = FALSE)
writeData(wb, "tx_new_psnu", tx_new_psnu)
addWorksheet(wb, "tx_new_site",
             tabColour = "#FFCC00",
             gridLines = FALSE)
writeData(wb, "tx_new_site", tx_new_site)

saveWorkbook(wb, file = "C:/Users/josep/Documents/R/r_projects/mer/output/ct/tx_summary.xlsx", overwrite = TRUE)

