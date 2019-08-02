rm(list = ls())

library(tidyverse)
library(janitor)
library(readxl)
library(openxlsx)


#---------------------------------------------
# GENIE IMPORT AND SET UP OF BASE DATAFRAME. NOTE THAT "FILTER" AND "SELECT" CODE LINES ARE ANALYSIS SPECIFIC 

df <- read.delim("~/R/datasets/Genie_Daily_82e86f5aaba344fa8e890f3ff91fa5ea.txt") %>%
  filter(Fiscal_Year == "2019") %>%
  select(-Qtr1, -Qtr2, -Qtr4) %>%
  rename(target = TARGETS,
         site_id = orgUnitUID,
         disag = standardizedDisaggregate,
         agency = FundingAgency,
         site = SiteName,
         partner = mech_name,
         Q = starts_with("Q")
  )

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
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", target > 0) %>%
  select(partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, target) %>%
  group_by(agency, partner, SNU1, PSNU) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# HTS REFERENCE OBJECT

hts_tst <- df %>%
  filter(indicator == "HTS_TST", disag == "Total Numerator", Q > 0) %>%
  select(indicator, disag, SNU1, PSNU, Fiscal_Year, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst = "Q")

cc_psnu <- tx_curr_target %>%
  left_join(hts_tst, by = "PSNU") %>%
  replace_na(list(hts_tst = 0))

#---------------------------------------------
# HTS_POS REFERENCE OBJECT

hts_tst_pos <- df %>%
  filter(indicator == "HTS_TST_POS", disag == "Total Numerator", Q > 0) %>%
  select(indicator, disag, SNU1, PSNU, Fiscal_Year, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst_pos = "Q")

cc_psnu <- cc_psnu %>%
  left_join(hts_tst_pos, by = "PSNU") %>%
  replace_na(list(hts_tst_pos = 0))

#---------------------------------------------
# TX_NEW REFERENCE OBJECT

tx_new <- df %>%
  filter(indicator == "TX_NEW", disag == "Total Numerator", Q > 0) %>%
  select(indicator, disag, SNU1, PSNU, Fiscal_Year, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_new, by = "PSNU") %>%
  replace_na(list(tx_new = 0))

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", Q > 0) %>%
  select(indicator, disag, SNU1, PSNU, Fiscal_Year, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_curr, by = "PSNU") %>%
  replace_na(list(tx_curr = 0))


#---------------------------------------------
# TX_PLVS_DEN REFERENCE OBJECT

tx_pvls_d <- df %>%
  filter(indicator == "TX_PVLS", disag == "Total Denominator", Q > 0) %>%
  select(indicator, disag, SNU1, PSNU, Fiscal_Year, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_pvls_d, by = "PSNU") %>%
  replace_na(list(tx_pvls_d = 0))


#---------------------------------------------
# TX_PLVS_NUM REFERENCE OBJECT

tx_pvls_n <- df %>%
  filter(indicator == "TX_PVLS", disag == "Total Numerator", Q > 0) %>%
  select(indicator, disag, SNU1, PSNU, Fiscal_Year, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = "Q") 


#---------------------------------------------
## CREATE DISTRICT CLINICAL CASCADE OBJECT

cc_psnu <- cc_psnu %>%
  left_join(tx_pvls_n, by = "PSNU") %>%
  replace_na(list(tx_pvls_n = 0)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_linkage = tx_new / hts_tst_pos,
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(agency, partner, SNU1, PSNU, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


#---------------------------------------------
## CREATE PROVINCIAL CLINICAL CASCADE OBJECT

cc_snu1 <- cc_psnu %>%
  group_by(agency, partner, SNU1) %>%
  summarize(tx_curr_target = sum(tx_curr_target, na.rm = TRUE),
            hts_tst = sum(hts_tst, na.rm = TRUE),
            hts_tst_pos = sum(hts_tst_pos, na.rm = TRUE),
            tx_new = sum(tx_new, na.rm = TRUE),
            tx_curr = sum(tx_curr, na.rm = TRUE),
            tx_pvls_d = sum(tx_pvls_d, na.rm = TRUE),
            tx_pvls_n = sum(tx_pvls_n, na.rm = TRUE)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_linkage = tx_new / hts_tst_pos,
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  ungroup() %>%
  select(agency, partner, SNU1, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


#---------------------------------------------
## CREATE PARTNER CLINICAL CASCADE OBJECT

cc_partner <- cc_psnu %>%
  group_by(agency, partner) %>%
  summarize(tx_curr_target = sum(tx_curr_target, na.rm = TRUE),
            hts_tst = sum(hts_tst, na.rm = TRUE),
            hts_tst_pos = sum(hts_tst_pos, na.rm = TRUE),
            tx_new = sum(tx_new, na.rm = TRUE),
            tx_curr = sum(tx_curr, na.rm = TRUE),
            tx_pvls_d = sum(tx_pvls_d, na.rm = TRUE),
            tx_pvls_n = sum(tx_pvls_n, na.rm = TRUE)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_linkage = tx_new / hts_tst_pos,
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  ungroup() %>%
  select(agency, partner, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


#---------------------------------------------
## CREATE AGENCY CLINICAL CASCADE OBJECT

cc_agency <- cc_psnu %>%
  group_by(agency) %>%
  summarize(tx_curr_target = sum(tx_curr_target, na.rm = TRUE),
            hts_tst = sum(hts_tst, na.rm = TRUE),
            hts_tst_pos = sum(hts_tst_pos, na.rm = TRUE),
            tx_new = sum(tx_new, na.rm = TRUE),
            tx_curr = sum(tx_curr, na.rm = TRUE),
            tx_pvls_d = sum(tx_pvls_d, na.rm = TRUE),
            tx_pvls_n = sum(tx_pvls_n, na.rm = TRUE)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_linkage = tx_new / hts_tst_pos,
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  ungroup() %>%
  select(agency, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


#---------------------------------------------
##CLEAN-UP R ENVIRONMENT BY REMOVING UNNEEDED OBJECTS

rm(tx_curr, tx_curr_target, tx_new, tx_pvls_d, tx_pvls_n, hts_tst, hts_tst_pos)


#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_curr OUTPUT

wb <- createWorkbook()
addWorksheet(wb, "cc_agency",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeData(wb, "cc_agency", cc_agency)
addWorksheet(wb, "cc_partner",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeData(wb, "cc_partner", cc_partner)
addWorksheet(wb, "cc_snu1",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeData(wb, "cc_snu1", cc_snu1)
addWorksheet(wb, "cc_psnu",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeData(wb, "cc_psnu", cc_psnu)
saveWorkbook(wb, file = "C:/Users/josep/Documents/R/r_projects/mer/output/ct/clinical_care_cascade_above_site.xlsx", overwrite = TRUE)
