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


#---------------------------------------------
## PREPARE AGENCY HTS_TST & HTS_TST_POS BASE DATASETS & JOIN THEN JOIN THE TWO

hts_tst <- df %>%
  filter(indicator == "HTS_TST", 
         disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_hts_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(hts_tst_Q1 = Q1,
         hts_tst_Q2 = Q2,
         hts_tst_Q3 = Q3,
         hts_tst_Q4 = Q4,
         hts_tst_target = target) %>%
  ungroup()

hts_tst_pos <- df %>%
  filter(indicator == "HTS_TST_POS", 
         disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), 
            Q2 = sum(Q2, na.rm = TRUE), 
            Q3 = sum(Q3, na.rm = TRUE), 
            Q4 = sum(Q4, na.rm = TRUE), 
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_hts_pos_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(hts_tst_pos_Q1 = Q1, 
         hts_tst_pos_Q2 = Q2, 
         hts_tst_pos_Q3 = Q3, 
         hts_tst_pos_Q4 = Q4, 
         hts_tst_pos_target = target) %>%
  ungroup()

hts_tst <- hts_tst %>%
  left_join(hts_tst_pos, by = "agency")

hts_tst_agency <- hts_tst %>%
  mutate(yield_Q1 = hts_tst_pos_Q1 / hts_tst_Q1, 
         yield_Q2 = hts_tst_pos_Q2 / hts_tst_Q2, 
         yield_Q3 = hts_tst_pos_Q3 / hts_tst_Q3, 
         yield_Q4 = hts_tst_pos_Q4 / hts_tst_Q4) %>%
  select(agency, hts_tst_Q1, hts_tst_Q2, hts_tst_Q3, hts_tst_Q4, hts_tst_pos_Q1, hts_tst_pos_Q2, hts_tst_pos_Q3, hts_tst_pos_Q4, yield_Q1, yield_Q2, yield_Q3, yield_Q4, per_hts_ytd_vs_target, per_hts_pos_ytd_vs_target)

#---------------------------------------------
## PREPARE AGENCY HTS_TST & HTS_TST_POS BASE DATASETS & JOIN THEN JOIN THE TWO

hts_tst <- df %>%
  filter(indicator == "HTS_TST", 
         disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_hts_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(hts_tst_Q1 = Q1,
         hts_tst_Q2 = Q2,
         hts_tst_Q3 = Q3,
         hts_tst_Q4 = Q4,
         hts_tst_target = target) %>%
  ungroup()

hts_tst_pos <- df %>%
  filter(indicator == "HTS_TST_POS", 
         disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), 
            Q2 = sum(Q2, na.rm = TRUE), 
            Q3 = sum(Q3, na.rm = TRUE), 
            Q4 = sum(Q4, na.rm = TRUE), 
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_hts_pos_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(hts_tst_pos_Q1 = Q1, 
         hts_tst_pos_Q2 = Q2, 
         hts_tst_pos_Q3 = Q3, 
         hts_tst_pos_Q4 = Q4, 
         hts_tst_pos_target = target) %>%
  ungroup()

hts_tst <- hts_tst %>%
  left_join(hts_tst_pos, by = "partner")

hts_tst_partner <- hts_tst %>%
  mutate(yield_Q1 = hts_tst_pos_Q1 / hts_tst_Q1, 
         yield_Q2 = hts_tst_pos_Q2 / hts_tst_Q2, 
         yield_Q3 = hts_tst_pos_Q3 / hts_tst_Q3, 
         yield_Q4 = hts_tst_pos_Q4 / hts_tst_Q4) %>%
  filter(per_hts_ytd_vs_target != "NaN") %>%
  select(partner, hts_tst_Q1, hts_tst_Q2, hts_tst_Q3, hts_tst_Q4, hts_tst_pos_Q1, hts_tst_pos_Q2, hts_tst_pos_Q3, hts_tst_pos_Q4, yield_Q1, yield_Q2, yield_Q3, yield_Q4, per_hts_ytd_vs_target, per_hts_pos_ytd_vs_target)



#---------------------------------------------
## PREPARE PROVINCE HTS_TST & HTS_TST_POS BASE DATASETS & JOIN THEN JOIN THE TWO


#---------------------------------------------
## PREPARE PROVINCE HTS_TST & HTS_TST_POS BASE DATASETS & JOIN THEN JOIN THE TWO

hts_tst <- df %>%
  filter(indicator == "HTS_TST", 
         disag == "Total Numerator") %>%
  select(site_id, SNU1, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(SNU1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_hts_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(hts_tst_Q1 = Q1,
         hts_tst_Q2 = Q2,
         hts_tst_Q3 = Q3,
         hts_tst_Q4 = Q4,
         hts_tst_target = target) %>%
  ungroup()

hts_tst_pos <- df %>%
  filter(indicator == "HTS_TST_POS", 
         disag == "Total Numerator") %>%
  select(site_id, SNU1, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(SNU1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), 
            Q2 = sum(Q2, na.rm = TRUE), 
            Q3 = sum(Q3, na.rm = TRUE), 
            Q4 = sum(Q4, na.rm = TRUE), 
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_hts_pos_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(hts_tst_pos_Q1 = Q1, 
         hts_tst_pos_Q2 = Q2, 
         hts_tst_pos_Q3 = Q3, 
         hts_tst_pos_Q4 = Q4, 
         hts_tst_pos_target = target) %>%
  ungroup()

hts_tst <- hts_tst %>%
  left_join(hts_tst_pos, by = "SNU1")

hts_tst_snu1 <- hts_tst %>%
  mutate(yield_Q1 = hts_tst_pos_Q1 / hts_tst_Q1, 
         yield_Q2 = hts_tst_pos_Q2 / hts_tst_Q2, 
         yield_Q3 = hts_tst_pos_Q3 / hts_tst_Q3, 
         yield_Q4 = hts_tst_pos_Q4 / hts_tst_Q4) %>%
  select(SNU1, hts_tst_Q1, hts_tst_Q2, hts_tst_Q3, hts_tst_Q4, hts_tst_pos_Q1, hts_tst_pos_Q2, hts_tst_pos_Q3, hts_tst_pos_Q4, yield_Q1, yield_Q2, yield_Q3, yield_Q4, per_hts_ytd_vs_target, per_hts_pos_ytd_vs_target)


#---------------------------------------------
## PREPARE DISTRICT HTS_TST & HTS_TST_POS BASE DATASETS & JOIN THEN JOIN THE TWO

hts_tst <- df %>%
  filter(indicator == "HTS_TST", 
         disag == "Total Numerator") %>%
  select(site_id, SNU1, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(SNU1, PSNU) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_hts_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(hts_tst_Q1 = Q1,
         hts_tst_Q2 = Q2,
         hts_tst_Q3 = Q3,
         hts_tst_Q4 = Q4,
         hts_tst_target = target) %>%
  ungroup()

hts_tst_pos <- df %>%
  filter(indicator == "HTS_TST_POS", 
         disag == "Total Numerator") %>%
  select(site_id, SNU1, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(PSNU) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), 
            Q2 = sum(Q2, na.rm = TRUE), 
            Q3 = sum(Q3, na.rm = TRUE), 
            Q4 = sum(Q4, na.rm = TRUE), 
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_hts_pos_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(hts_tst_pos_Q1 = Q1, 
         hts_tst_pos_Q2 = Q2, 
         hts_tst_pos_Q3 = Q3, 
         hts_tst_pos_Q4 = Q4, 
         hts_tst_pos_target = target) %>%
  ungroup()

hts_tst <- hts_tst %>%
  left_join(hts_tst_pos, by = "PSNU")

hts_tst_psnu <- hts_tst %>%
  mutate(yield_Q1 = hts_tst_pos_Q1 / hts_tst_Q1, 
         yield_Q2 = hts_tst_pos_Q2 / hts_tst_Q2, 
         yield_Q3 = hts_tst_pos_Q3 / hts_tst_Q3, 
         yield_Q4 = hts_tst_pos_Q4 / hts_tst_Q4) %>%
  select(SNU1, PSNU, hts_tst_Q1, hts_tst_Q2, hts_tst_Q3, hts_tst_Q4, hts_tst_pos_Q1, hts_tst_pos_Q2, hts_tst_pos_Q3, hts_tst_pos_Q4, yield_Q1, yield_Q2, yield_Q3, yield_Q4, per_hts_ytd_vs_target, per_hts_pos_ytd_vs_target)



wb <- createWorkbook()
addWorksheet(wb, "hts_tst_agency",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeData(wb, "hts_tst_agency", hts_tst_agency)
addWorksheet(wb, "hts_tst_snu1",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeData(wb, "hts_tst_snu1", hts_tst_snu1)
addWorksheet(wb, "hts_tst_psnu",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeData(wb, "hts_tst_psnu", hts_tst_psnu)
saveWorkbook(wb, file = "C:/Users/josep/Documents/R/r_projects/mer/output/hts/hts_summary.xlsx", overwrite = TRUE)

