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
# IMPORT OF COP19 AJUDA SITE LIST

ajuda <- read_excel("R/datasets/ajuda_sites_ia.xlsx") %>%
  replace_na(list(ajuda = 0))


#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = starts_with("Q")) 


#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df %>%
  filter(indicator == "HTS_TST",
         disag == "Total Numerator",
         partner == "CHASS" | partner == "FGH" | partner == "ICAP" | partner == "CCS" | partner == "ARIEL" | partner == "EGPAF",
         Q > 0
  ) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst = starts_with("Q")) %>%
  replace_na(list(hts_tst = 0))


#---------------------------------------------
# HTS_TST_POS REFERENCE OBJECT

hts_tst_pos <- df %>%
  filter(indicator == "HTS_TST_POS",
         disag == "Total Numerator",
         partner == "CHASS" | partner == "FGH" | partner == "ICAP" | partner == "CCS" | partner == "ARIEL" | partner == "EGPAF",
         Q > 0
  ) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst_pos = starts_with("Q")) %>%
  replace_na(list(hts_tst_pos = 0))


#---------------------------------------------
# TX_NEW REFERENCE OBJECT

tx_new <- df %>%
  filter(indicator == "TX_NEW", disag == "Total Numerator", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = starts_with("Q"))


#---------------------------------------------
# TX_PVLS DENOMINATOR REFERENCE OBJECT

tx_pvls_d <- df %>%
  filter(indicator == "TX_PVLS", disag == "Total Denominator", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = starts_with("Q"))


#---------------------------------------------
# TX_PVLS NUMERATOR REFERENCE OBJECT

tx_pvls_n <- df %>%
  filter(indicator == "TX_PVLS", disag == "Total Numerator", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = starts_with("Q"))


#---------------------------------------------
# JOIN ALL CASCADE INDICATOR OBJECTS INTO ONE DATAFRAME

cc <- tx_curr %>%
  left_join(hts_tst, by = "site_id") %>%
  left_join(hts_tst_pos, by = "site_id") %>%
  left_join(tx_new, by = "site_id") %>%
  left_join(tx_pvls_d, by = "site_id") %>%
  left_join(tx_pvls_n, by  = "site_id") %>%
  left_join(ajuda, by = "site_id") %>%
  replace_na(list(hts_tst = 0,
                  hts_tst_pos = 0,
                  tx_new = 0,
                  tx_curr = 0,
                  tx_pvls_d = 0,
                  tx_pvls_n = 0,
                  ajuda = 0)
  ) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_linkage = tx_new / hts_tst_pos,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(site_id, agency, partner, SNU1, PSNU, site, ajuda, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


#---------------------------------------------
##CLEAN-UP R ENVIRONMENT BY REMOVING UNNEEDED OBJECTS

rm(ajuda, tx_curr, tx_new, tx_pvls_d, tx_pvls_n, hts_tst, hts_tst_pos)

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_curr OUTPUT

wb <- createWorkbook()
addWorksheet(wb, "cc_site",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeData(wb, "cc_site", cc)
saveWorkbook(wb, file = "C:/Users/josep/Documents/R/r_projects/mer/output/ct/clinical_care_cascade_site.xlsx", overwrite = TRUE)