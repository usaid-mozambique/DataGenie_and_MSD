rm(list = ls())

library(tidyverse)
library(janitor)
library(readxl)
library(openxlsx)


#---------------------------------------------
# GENIE IMPORT AND SET UP OF BASE DATAFRAME. NOTE THAT "FILTER" AND "SELECT" CODE LINES ARE ANALYSIS SPECIFIC 

df <- read.delim("~/R/datasets/Genie_Daily_80bc795352084a32adb1299ebaa7a132.txt") %>% ####### NEEDS TO BE MANUALLY IMPORTED THROUGH MENU WITH EACH DATA GENIE UPDATE
  rename(target = TARGETS, 
         site_id = FacilityUID, 
         disag = standardizedDisaggregate,
         disag_p = otherDisaggregate,
         agency = FundingAgency, 
         site = SiteName, 
         partner = mech_name, 
         Q1 = Qtr1, 
         Q2 = Qtr2, 
         Q3 = Qtr3, 
         Q4 = Qtr4,
         Cum = Cumulative)

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
## PREPARE PMTCT_ART DATASET

p_art_pos <- df %>%
  filter(indicator == "PMTCT_ART", Fiscal_Year != "2020", disag == "Total Denominator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site, Fiscal_Year) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE)) %>%
  ungroup()

p_art_pos_wide <- p_art_pos %>%
  pivot_wider(names_from = "Fiscal_Year", values_from = "Q1")



p_art_pos_2018 <- df %>%
  filter(indicator == "PMTCT_ART" & Fiscal_Year == "2018", disag == "Total Denominator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE)) %>%
  ungroup()

p_art_alr_2018 <- df %>%
  filter(indicator == "PMTCT_ART" & Fiscal_Year == "2018", disag_p == "Life-long ART, Already") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE)) %>%
  ungroup()

p_art_new_2018 <- df %>%
  filter(indicator == "PMTCT_ART" & Fiscal_Year == "2018", disag_p == "Life-long ART, New") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE)) %>%
  ungroup()

p_art_pos_2019 <- df %>%
  filter(indicator == "PMTCT_ART" & Fiscal_Year == "2019", disag == "Total Denominator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE)) %>%
  ungroup()

p_art_alr_2019 <- df %>%
  filter(indicator == "PMTCT_ART" & Fiscal_Year == "2019", disag_p == "Life-long ART, Already") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE)) %>%
  ungroup()

p_art_new_2019 <- df %>%
  filter(indicator == "PMTCT_ART" & Fiscal_Year == "2019", disag_p == "Life-long ART, New") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE)) %>%
  ungroup()
              



p_stat_pos <- df %>%
  filter(indicator == "PMTCT_STAT" & Fiscal_Year == "2019", StatusHIV == "Positive") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE) %>%
  ungroup() %>%
  rename(p_stat_pos = starts_with("Q")) %>%
  select(-partner, -agency, -SNU1, -PSNU, -site)

p_stat_pos_new <- df %>%
  filter(indicator == "PMTCT_STAT" & Fiscal_Year == "2019", StatusHIV == "Positive", disag_p == "Newly Identified") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_pos_new = starts_with("Q")) %>%
  select(-partner, -agency, -SNU1, -PSNU, -site)

p_stat_pos_alr <- df %>%
  filter(indicator == "PMTCT_STAT" & Fiscal_Year == "2019", StatusHIV == "Positive", disag_p == "Known at Entry") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_pos_alr = starts_with("Q")) %>%
  select(-partner, -agency, -SNU1, -PSNU, -site)











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
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")


wb <- createWorkbook()
addWorksheet(wb, "tx_curr_agency",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb, "tx_curr_agency", tx_curr_agency, tableStyle = "TableStyleLight2")
addStyle(wb, "tx_curr_agency", style = pct, cols=c(7:9), rows = 2:(nrow(tx_curr_agency)+1), gridExpand=TRUE)


addWorksheet(wb, "tx_curr_snu1",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb, "tx_curr_snu1", tx_curr_snu1, tableStyle = "TableStyleLight2")
addStyle(wb, "tx_curr_snu1", style = pct, cols=c(9:11), rows = 2:(nrow(tx_curr_snu1)+1), gridExpand=TRUE)


addWorksheet(wb, "tx_curr_psnu",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb, "tx_curr_psnu", tx_curr_psnu, tableStyle = "TableStyleLight2")
addStyle(wb, "tx_curr_psnu", style = pct, cols=c(10:12), rows = 2:(nrow(tx_curr_psnu)+1), gridExpand=TRUE)


addWorksheet(wb, "tx_curr_site",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb, "tx_curr_site", tx_curr_site, tableStyle = "TableStyleLight2")
addStyle(wb, "tx_curr_site", style = pct, cols=c(12:14), rows = 2:(nrow(tx_curr_site)+1), gridExpand=TRUE)


addWorksheet(wb, "tx_new_agency",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb, "tx_new_agency", tx_new_agency, tableStyle = "TableStyleLight3")
addStyle(wb, "tx_new_agency", style = pct, cols=c(7:8), rows = 2:(nrow(tx_new_agency)+1), gridExpand=TRUE)


addWorksheet(wb, "tx_new_snu1",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb, "tx_new_snu1", tx_new_snu1, tableStyle = "TableStyleLight3")
addStyle(wb, "tx_new_snu1", style = pct, cols=c(9:10), rows = 2:(nrow(tx_new_snu1)+1), gridExpand=TRUE)


addWorksheet(wb, "tx_new_psnu",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb, "tx_new_psnu", tx_new_psnu, tableStyle = "TableStyleLight3")
addStyle(wb, "tx_new_psnu", style = pct, cols=c(10:11), rows = 2:(nrow(tx_new_psnu)+1), gridExpand=TRUE)


addWorksheet(wb, "tx_new_site",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb, "tx_new_site", tx_new_site, tableStyle = "TableStyleLight3")
addStyle(wb, "tx_new_site", style = pct, cols=c(12:13), rows = 2:(nrow(tx_new_site)+1), gridExpand=TRUE)

saveWorkbook(wb, file = "C:/Users/josep/Documents/R/r_projects/mer/output/ct/tx_summary.xlsx", overwrite = TRUE)

