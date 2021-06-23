rm(list = ls())
library(tidyr)
library(stringr)
library(dplyr)
library(stringr)
library(tibble)
library(janitor)
library(readxl)
library(openxlsx)

#---------------------------------------------
# GENIE IMPORT AND SET UP OF BASE DATAFRAME. NOTE THAT "FILTER" CODE LINE ARE ANALYSIS SPECIFIC 

#df <- read.delim("C:/Users/jlara/Documents/USAID/00. SI/00. HIV/01. M&E/00. Data/01. MER/FY 2019/Q4/99. Data/MER_Structured_Datasets_SITE_IM_FY17-20_20191115_v1_1_Mozambique.txt") %>% ####### NEEDS TO BE MANUALLY IMPORTED THROUGH MENU WITH EACH DATA GENIE UPDATE
df <- read.delim("C:/Users/cnhantumbo/Documents/DataGenie/SOURCE.txt") %>% ####### NEEDS TO BE MANUALLY IMPORTED THROUGH MENU WITH EACH DATA GENIE UPDATE 
 rename(target = targets, 
         site_id = facilityuid, 
         disag = standardizeddisaggregate,
         disag_p = otherdisaggregate,
         agency = fundingagency, 
         site = sitename, 
         partner = mech_name, 
         Q1 = qtr1, 
         Q2 = qtr2, 
         Q3 = qtr3, 
         Q4 = qtr4,
         YtD = cumulative)

df$partner <- as.character(df$partner)


#--------------------------------Updated the Mechanisms---------------------------------------------------------
df <- df %>%
  mutate(partner = 
           if_else(partner == "FADM HIV Treatment Scale-Up Program", "FADM",
                   if_else(partner == "Community HIV Activity in Zambezia", "N'weti Zambezia",
                           if_else(partner == "Friends in Global Health", "FGH", 
                                   if_else(partner == "FADM Prevention and Circumcision Program", "FADM PCP",
                                           if_else(partner == "N'weti - Strengthening Civil Society Engagement to Improve Sexual and Reproductive Health and Service Delivery for Youth", "N'weti", 
                                                   if_else(partner == "VMMC Services in Manica and Tete", "VMMC", 
                                                           if_else(partner =="Community Based HIV Services for the Southern Region","N'weti", 
                                                                   if_else(partner == "Integrated HIV Prevention and Health Services for Key and Priority Populations (HIS-KP)", "Passos",
                                                                           if_else(partner == "FADM Prevention and Circumcision Program", "FADM",
                                                                                   if_else(partner == "Johns Hopkins", "Jhpiego",
                                                                                           if_else(partner == "Efficiencies for Clinical HIV Outcomes (ECHO)", "ECHO", partner)
                                                                                   )
                                                                           )
                                                                   )
                                                           )
                                                   )
                                           )
                                   )
                           )
                   )
           )
  ) 


df_a <- df %>%
  filter(fiscal_year == "2020")

df_q <- df_a  %>%
  select(-Q4, -Q2, -Q3) %>%
  rename(Q = starts_with("Q1")
  )

#ajuda <- read_excel("~/R/datasets/ajuda_ia.xlsx") %>%
ajuda <- read_excel("C:/Users/cnhantumbo/Documents/AJUDA/ajuda_ia.xlsx") %>%
  replace_na(list(ajuda = 0))

##############################################
##############################################
##############################################
## HTS #######################################
##############################################
##############################################
##############################################

#---------------------------------------------
## PREPARE IP HTS_TST & HTS_TST_POS BASE DATASETS & JOIN THEN JOIN THE TWO

hts_tst_partner <- df_a %>%
  filter(indicator == "HTS_TST", 
         disag == "Total Numerator") %>%
  group_by(agency, partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(hts_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target,
         hts_tst_cum = Q1 + Q2 + Q3 + Q4) %>%
  rename(hts_tst_Q1 = Q1,
         hts_tst_Q2 = Q2,
         hts_tst_Q3 = Q3,
         hts_tst_Q4 = Q4,
         hts_tst_target = target) %>%
  ungroup()

hts_tst_pos_partner <- df_a %>%
  filter(indicator == "HTS_TST_POS", 
         disag == "Total Numerator") %>%
  group_by(agency, partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(pos_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target,
         hts_pos_cum = Q1 + Q2 + Q3 + Q4) %>%
  rename(hts_pos_Q1 = Q1,
         hts_pos_Q2 = Q2,
         hts_pos_Q3 = Q3,
         hts_pos_Q4 = Q4,
         hts_pos_target = target) %>%
  ungroup()

ip_snapshot <- hts_tst_partner %>%
  left_join(hts_tst_pos_partner, by = "partner") %>%
  select(-agency.y) %>%
  rename(agency = agency.x) %>%
  select(agency, partner, hts_tst_Q1, hts_pos_Q1, hts_tst_Q2, hts_pos_Q2, hts_tst_Q3, hts_pos_Q3, hts_tst_Q4, hts_pos_Q4, hts_tst_cum, hts_pos_cum, hts_ytd_vs_target, pos_ytd_vs_target)  %>%
  mutate(yield_Q1 = hts_pos_Q1 / hts_tst_Q1,
         yield_Q2 = hts_pos_Q2 / hts_tst_Q2,
         yield_Q3 = hts_pos_Q3 / hts_tst_Q3,
         yield_Q4 = hts_pos_Q4 / hts_tst_Q4,
         yield_cum = hts_pos_cum / hts_tst_cum) %>%
  filter(partner != "Dedup",
         hts_tst_cum > 0)

#---------------------------------------------
## PREPARE MODALITY HTS_TST & HTS_TST_POS BASE DATASETS & JOIN THEN JOIN THE TWO

hts_tst_modality <- df_a %>%
  filter(indicator == "HTS_TST") %>%
  group_by(modality) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(hts_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target,
         hts_tst_cum = Q1 + Q2 + Q3 + Q4) %>%
  rename(hts_tst_Q1 = Q1,
         hts_tst_Q2 = Q2,
         hts_tst_Q3 = Q3,
         hts_tst_Q4 = Q4,
         hts_tst_target = target) %>%
  ungroup()

hts_tst_pos_modality <- df_a %>%
  filter(indicator == "HTS_TST_POS") %>%
  group_by(modality) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(pos_ytd_vs_target = (Q1 + Q2 + Q3 + Q4)/target,
         hts_pos_cum = Q1 + Q2 + Q3 + Q4) %>%
  rename(hts_pos_Q1 = Q1,
         hts_pos_Q2 = Q2,
         hts_pos_Q3 = Q3,
         hts_pos_Q4 = Q4,
         hts_pos_target = target) %>%
  ungroup()

modality_snapshot <- hts_tst_modality %>%
  left_join(hts_tst_pos_modality, by = "modality") %>%
  select(modality, hts_tst_Q1, hts_pos_Q1, hts_tst_Q2, hts_pos_Q2, hts_tst_Q3, hts_pos_Q3, hts_tst_Q4, hts_pos_Q4, hts_tst_cum, hts_pos_cum, hts_ytd_vs_target, pos_ytd_vs_target)  %>%
  mutate(yield_Q1 = hts_pos_Q1 / hts_tst_Q1,
         yield_Q2 = hts_pos_Q2 / hts_tst_Q2,
         yield_Q3 = hts_pos_Q3 / hts_tst_Q3,
         yield_Q4 = hts_pos_Q4 / hts_tst_Q4,
         yield_cum = hts_pos_cum / hts_tst_cum) %>%
  filter(modality != "",
         hts_tst_cum > 0)

#---------
wb_overall <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")
pct2 <- createStyle(numFmt = "0.0%", textDecoration = "italic")

addWorksheet(wb_overall, "ip_snapshot",
             gridLines = FALSE)
writeDataTable(wb_overall, "ip_snapshot", ip_snapshot, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "ip_snapshot", style = pct, cols=c(13:14), rows = 2:(nrow(ip_snapshot)+1), gridExpand=TRUE)
addStyle(wb_overall, "ip_snapshot", style = pct2, cols=c(15:19), rows = 2:(nrow(ip_snapshot)+1), gridExpand=TRUE)
setColWidths(wb_overall, "ip_snapshot", cols = c(1:2), widths = "auto")

addWorksheet(wb_overall, "modality_snapshot",
             gridLines = FALSE)
writeDataTable(wb_overall, "modality_snapshot", modality_snapshot, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "modality_snapshot", style = pct, cols=c(12:13), rows = 2:(nrow(modality_snapshot)+1), gridExpand=TRUE)
addStyle(wb_overall, "modality_snapshot", style = pct2, cols=c(14:18), rows = 2:(nrow(modality_snapshot)+1), gridExpand=TRUE)
setColWidths(wb_overall, "modality_snapshot", cols = c(1:2), widths = "auto")

saveWorkbook(wb_overall, file = "C:/Users/cnhantumbo/Documents/R codes/hts_summary.xlsx", overwrite = TRUE)

##############################################
##############################################
##############################################
## VMMC ######################################
##############################################
##############################################
##############################################

#---SITE LEVEL VMMC_CIRC PERFORMANCE-------

by_site <- df_a %>%
  filter(indicator == "VMMC_CIRC", 
         disag == "Total Numerator",
         (YtD > 0 | target > 0)) %>%
  select(site_id, agency, partner, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, YtD, target) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            YtD = sum(YtD, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_vmmc_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(vmmc_Q1 = Q1,
         vmmc_Q2 = Q2,
         vmmc_Q3 = Q3,
         vmmc_Q4 = Q4,
         vmmc_target = target) %>%
  arrange(agency, partner, snu1, psnu, site)
ungroup()

#---------------------------------------------
##DISTRICT LEVEL VMMC_CIRC PERFORMANCE

by_psnu <- df_a %>%
  filter(indicator == "VMMC_CIRC", 
         disag == "Total Numerator",
         (YtD > 0 | target > 0)) %>%
  select(site_id, agency, partner, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, YtD, target) %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            YtD = sum(YtD, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_vmmc_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(vmmc_Q1 = Q1,
         vmmc_Q2 = Q2,
         vmmc_Q3 = Q3,
         vmmc_Q4 = Q4,
         vmmc_target = target) %>%
  ungroup()

#---------------------------------------------
##PROVINCIAL LEVEL VMMC_CIRC PERFORMANCE

by_snu1 <- df_a %>%
  filter(indicator == "VMMC_CIRC", 
         disag == "Total Numerator",
         (YtD > 0 | target > 0)) %>%
  select(site_id, agency, partner, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, YtD, target) %>%
  group_by(agency, partner, snu1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            YtD = sum(YtD, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_vmmc_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(vmmc_Q1 = Q1,
         vmmc_Q2 = Q2,
         vmmc_Q3 = Q3,
         vmmc_Q4 = Q4,
         vmmc_target = target) %>%
  ungroup()

#---------------------------------------------
##AGENCY LEVEL VMMC_CIRC PERFORMANCE

by_partner <- df_a %>%
  filter(indicator == "VMMC_CIRC", 
         disag == "Total Numerator",
         (YtD > 0 | target > 0)) %>%
  select(site_id, agency, partner, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, YtD, target) %>%
  group_by(agency, partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            YtD = sum(YtD, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_vmmc_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(vmmc_Q1 = Q1,
         vmmc_Q2 = Q2,
         vmmc_Q3 = Q3,
         vmmc_Q4 = Q4,
         vmmc_target = target) %>%
  ungroup()

#---------------------------------------------
##AGENCY LEVEL VMMC_CIRC PERFORMANCE

by_agency <- df_a %>%
  filter(indicator == "VMMC_CIRC", 
         disag == "Total Numerator",
         (YtD > 0 | target > 0)) %>%
  select(site_id, agency, partner, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, YtD, target) %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE),
            Q2 = sum(Q2, na.rm = TRUE),
            Q3 = sum(Q3, na.rm = TRUE),
            Q4 = sum(Q4, na.rm = TRUE),
            YtD = sum(YtD, na.rm = TRUE),
            target = sum(target, na.rm = TRUE)) %>%
  mutate(per_vmmc_vs_target = (Q1 + Q2 + Q3 + Q4)/target) %>%
  rename(vmmc_Q1 = Q1,
         vmmc_Q2 = Q2,
         vmmc_Q3 = Q3,
         vmmc_Q4 = Q4,
         vmmc_target = target) %>%
  ungroup()

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH VMMC_CIRC OUTPUT

wb <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")

wb <- createWorkbook()
addWorksheet(wb, "by_agency",
             gridLines = FALSE)
writeDataTable(wb, "by_agency", by_agency, tableStyle = "TableStyleLight2")
addStyle(wb, "by_agency", style = pct, cols=c(8:9), rows = 2:(nrow(by_agency)+1), gridExpand=TRUE)
setColWidths(wb, "by_agency", cols = c(1:1), widths = "auto")

addWorksheet(wb, "by_partner",
             gridLines = FALSE)
writeDataTable(wb, "by_partner", by_partner, tableStyle = "TableStyleLight2")
addStyle(wb, "by_partner", style = pct, cols=c(9), rows = 2:(nrow(by_partner)+1), gridExpand=TRUE)
setColWidths(wb, "by_partner", cols = c(1:2), widths = "auto")

addWorksheet(wb, "by_snu1",
             gridLines = FALSE)
writeDataTable(wb, "by_snu1", by_snu1, tableStyle = "TableStyleLight2")
addStyle(wb, "by_snu1", style = pct, cols=c(10), rows = 2:(nrow(by_snu1)+1), gridExpand=TRUE)
setColWidths(wb, "by_snu1", cols = c(1:3), widths = "auto")

addWorksheet(wb, "by_psnu",
             gridLines = FALSE)
writeDataTable(wb, "by_psnu", by_psnu, tableStyle = "TableStyleLight2")
addStyle(wb, "by_psnu", style = pct, cols=c(11), rows = 2:(nrow(by_psnu)+1), gridExpand=TRUE)
setColWidths(wb, "by_psnu", cols = c(1:4), widths = "auto")

addWorksheet(wb, "by_site",
             gridLines = FALSE)
writeDataTable(wb, "by_site", by_site, tableStyle = "TableStyleLight2")
addStyle(wb, "by_site", style = pct, cols=c(13), rows = 2:(nrow(by_site)+1), gridExpand=TRUE)
setColWidths(wb, "by_site", cols = c(1:6), widths = "auto")

saveWorkbook(wb, file = "C:/Users/cnhantumbo/Documents/R codes/vmmc_summary.xlsx", overwrite = TRUE)

##############################################
##############################################
##############################################
## PMTCT_STAT ################################
##############################################
##############################################
##############################################

#---------------------------------------------
# PMTCT_STAT_DEN REFERENCE OBJECT

p_stat_d <- df_q %>%
  filter(indicator == "PMTCT_STAT", disag == "Total Denominator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_d = starts_with("Q"))

#---------------------------------------------
# PMTCT_STAT_NUM REFERENCE OBJECT

p_stat_n <- df_q %>%
  filter(indicator == "PMTCT_STAT", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_n = starts_with("Q")) %>%
  select(-partner, -agency, -snu1, -psnu, -site)

#---------------------------------------------
# PMTCT_STAT_POS REFERENCE OBJECT

p_stat_pos <- df_q %>%
  filter(indicator == "PMTCT_STAT", statushiv == "Positive", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_pos = starts_with("Q")) %>%
  select(-partner, -agency, -snu1, -psnu, -site)

#---------------------------------------------
# PMTCT_ART REFERENCE OBJECT

p_art <- df_q %>%
  filter(indicator == "PMTCT_ART", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_art = starts_with("Q")) %>%
  select(-partner, -agency, -snu1, -psnu, -site)

#---------------------------------------------
# NEWLY POSITIVE VS. KNOWN HIV+ AT ENTRY REFERENCE OBJECT

p_stat_pos_new <- df_q %>%
  filter(indicator == "PMTCT_STAT", statushiv == "Positive", disag_p == "Newly Identified", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_pos_new = starts_with("Q")) %>%
  select(-partner, -agency, -snu1, -psnu, -site)

p_stat_pos_alr <- df_q %>%
  filter(indicator == "PMTCT_STAT", statushiv == "Positive", disag_p == "Known at Entry", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_pos_alr = starts_with("Q")) %>%
  select(-partner, -agency, -snu1, -psnu, -site)

#---------------------------------------------
# NEWLY STARTED ART VS. ON ART AT ENTRY REFERENCE OBJECT

p_art_new <- df_q %>%
  filter(indicator == "PMTCT_ART", disag_p == "Life-long ART, New", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_art_new = starts_with("Q")) %>%
  select(-partner, -agency, -snu1, -psnu, -site)

p_art_alr <- df_q %>%
  filter(indicator == "PMTCT_ART", disag_p == "Life-long ART, Already", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_art_alr = starts_with("Q")) %>%
  select(-partner, -agency, -snu1, -psnu, -site)

#---------------------------------------------
# JOIN ALL CASCADE INDICATOR OBJECTS INTO ONE DATAFRAME

pc <- p_stat_d %>%
  full_join(p_stat_n, by = "site_id") %>%
  full_join(p_stat_pos, by = "site_id") %>%
  full_join(p_art, by = "site_id") %>%
  left_join(ajuda, by = "site_id") %>% # USING LEFT JOIN RESULTS IN 590 AJUDA SITES BEING PRESENT IN DATASET
  full_join(p_stat_pos_new, by = "site_id") %>%
  full_join(p_stat_pos_alr, by = "site_id") %>%
  full_join(p_art_new, by = "site_id") %>%
  full_join(p_art_alr, by = "site_id") %>%
  replace_na(list(p_stat_d = 0,
                  p_stat_n = 0,
                  p_stat_pos = 0,
                  p_art = 0,
                  p_stat_pos_new = 0,
                  p_stat_pos_alr = 0,
                  p_art_new = 0,
                  p_art_alr = 0,
                  ajuda = 0)
  ) %>%
  select(1:6, ajuda, everything())

#---------------------------------------------
# PMTCT CASADE BY AGENCY

by_agency <- pc %>%
  group_by(agency) %>%
  summarize(p_stat_d = sum(p_stat_d, na.rm = TRUE),
            p_stat_n = sum(p_stat_n, na.rm = TRUE),
            p_stat_pos = sum(p_stat_pos, na.rm = TRUE),
            p_stat_pos_alr = sum(p_stat_pos_alr, na.rm = TRUE),
            p_stat_pos_new = sum(p_stat_pos_new, na.rm = TRUE),
            p_art = sum(p_art, na.rm = TRUE),
            p_art_alr = sum(p_art_alr, na.rm = TRUE), 
            p_art_new = sum(p_art_new, na.rm = TRUE)) %>%
  ungroup() %>%
  mutate(per_p_stat = p_stat_n / p_stat_d, 
         per_p_stat_pos = p_stat_pos / p_stat_n,
         per_p_art = p_art / p_stat_pos
  ) %>%
  select(1:3, per_p_stat, p_stat_pos, per_p_stat_pos, p_stat_pos_alr, p_stat_pos_new, everything()) %>%
  arrange(agency)

#---------------------------------------------
# PMTCT CASADE BY PARTNER

by_partner <- pc %>%
  group_by(agency, partner) %>%
  summarize(p_stat_d = sum(p_stat_d, na.rm = TRUE),
            p_stat_n = sum(p_stat_n, na.rm = TRUE),
            p_stat_pos = sum(p_stat_pos, na.rm = TRUE),
            p_stat_pos_alr = sum(p_stat_pos_alr, na.rm = TRUE),
            p_stat_pos_new = sum(p_stat_pos_new, na.rm = TRUE),
            p_art = sum(p_art, na.rm = TRUE),
            p_art_alr = sum(p_art_alr, na.rm = TRUE), 
            p_art_new = sum(p_art_new, na.rm = TRUE)) %>%
  ungroup() %>%
  mutate(per_p_stat = p_stat_n / p_stat_d, 
         per_p_stat_pos = p_stat_pos / p_stat_n,
         per_p_art = p_art / p_stat_pos
  ) %>%
  select(1:4, per_p_stat, p_stat_pos, per_p_stat_pos, p_stat_pos_alr, p_stat_pos_new, everything()) %>%
  arrange(agency, partner)

#---------------------------------------------
# PMTCT CASADE BY PROVINCE

by_snu1 <- pc %>%
  group_by(agency, partner, snu1) %>%
  summarize(p_stat_d = sum(p_stat_d, na.rm = TRUE),
            p_stat_n = sum(p_stat_n, na.rm = TRUE),
            p_stat_pos = sum(p_stat_pos, na.rm = TRUE),
            p_stat_pos_alr = sum(p_stat_pos_alr, na.rm = TRUE),
            p_stat_pos_new = sum(p_stat_pos_new, na.rm = TRUE),
            p_art = sum(p_art, na.rm = TRUE),
            p_art_alr = sum(p_art_alr, na.rm = TRUE), 
            p_art_new = sum(p_art_new, na.rm = TRUE)) %>%
  ungroup() %>%
  mutate(per_p_stat = p_stat_n / p_stat_d, 
         per_p_stat_pos = p_stat_pos / p_stat_n,
         per_p_art = p_art / p_stat_pos
  ) %>%
  select(1:5, per_p_stat, p_stat_pos, per_p_stat_pos, p_stat_pos_alr, p_stat_pos_new, everything()) %>%
  arrange(agency, partner, snu1)

#---------------------------------------------
# PMTCT CASADE BY DISTRICT

by_psnu <- pc %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(p_stat_d = sum(p_stat_d, na.rm = TRUE),
            p_stat_n = sum(p_stat_n, na.rm = TRUE),
            p_stat_pos = sum(p_stat_pos, na.rm = TRUE),
            p_stat_pos_alr = sum(p_stat_pos_alr, na.rm = TRUE),
            p_stat_pos_new = sum(p_stat_pos_new, na.rm = TRUE),
            p_art = sum(p_art, na.rm = TRUE),
            p_art_alr = sum(p_art_alr, na.rm = TRUE), 
            p_art_new = sum(p_art_new, na.rm = TRUE)) %>%
  ungroup() %>%
  mutate(per_p_stat = p_stat_n / p_stat_d, 
         per_p_stat_pos = p_stat_pos / p_stat_n,
         per_p_art = p_art / p_stat_pos
  ) %>%
  select(1:6, per_p_stat, p_stat_pos, per_p_stat_pos, p_stat_pos_alr, p_stat_pos_new, everything()) %>%
  arrange(agency, partner, snu1, psnu)

#---------------------------------------------
# PMTCT CASADE BY OU

by_ou <- by_psnu%>%
  mutate(agency = "Moz") %>%
  group_by(agency) %>%
  summarize(p_stat_d = sum(p_stat_d, na.rm = TRUE),
            p_stat_n = sum(p_stat_n, na.rm = TRUE),
            p_stat_pos = sum(p_stat_pos, na.rm = TRUE),
            p_stat_pos_alr = sum(p_stat_pos_alr, na.rm = TRUE),
            p_stat_pos_new = sum(p_stat_pos_new, na.rm = TRUE),
            p_art = sum(p_art, na.rm = TRUE),
            p_art_alr = sum(p_art_alr, na.rm = TRUE), 
            p_art_new = sum(p_art_new, na.rm = TRUE)) %>%
  ungroup() %>%
  mutate(per_p_stat = p_stat_n / p_stat_d, 
         per_p_stat_pos = p_stat_pos / p_stat_n,
         per_p_art = p_art / p_stat_pos
  ) %>%
  select(1:3, per_p_stat, p_stat_pos, per_p_stat_pos, p_stat_pos_alr, p_stat_pos_new, everything())

#---------------------------------------------
# PMTCT CASCADE BY SITE

by_site <- pc %>%
  mutate(per_p_stat = p_stat_n / p_stat_d, 
         per_p_stat_pos = p_stat_pos / p_stat_n,
         per_p_art = p_art / p_stat_pos
  ) %>%
  select(1:9, per_p_stat, p_stat_pos, per_p_stat_pos, p_stat_pos_alr, p_stat_pos_new, everything()) %>%
  arrange(agency, partner, snu1, psnu, site)

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH PMTCT_STAT CASCADE OUTPUT

wb <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")
pct2 <- createStyle(numFmt = "0.0%", textDecoration = "italic")

addWorksheet(wb, "by_ou",
             gridLines = FALSE)
writeDataTable(wb, "by_ou", by_ou, tableStyle = "TableStyleLight2")
addStyle(wb, "by_ou", style = pct, cols=c(4,12), rows = 2:(nrow(by_ou)+1), gridExpand=TRUE)
addStyle(wb, "by_ou", style = pct2, cols=c(6), rows = 2:(nrow(by_ou)+1), gridExpand=TRUE)

addWorksheet(wb, "by_agency",
             gridLines = FALSE)
writeDataTable(wb, "by_agency", by_agency, tableStyle = "TableStyleLight2")
addStyle(wb, "by_agency", style = pct, cols=c(4,6,12), rows = 2:(nrow(by_agency)+1), gridExpand=TRUE)
addStyle(wb, "by_agency", style = pct2, cols=c(6), rows = 2:(nrow(by_agency)+1), gridExpand=TRUE)

addWorksheet(wb, "by_partner",
             gridLines = FALSE)
writeDataTable(wb, "by_partner", by_partner, tableStyle = "TableStyleLight2")
addStyle(wb, "by_partner", style = pct, cols=c(5,14), rows = 2:(nrow(by_partner)+1), gridExpand=TRUE)
addStyle(wb, "by_partner", style = pct2, cols=c(7), rows = 2:(nrow(by_partner)+1), gridExpand=TRUE)

addWorksheet(wb, "by_snu1",
             gridLines = FALSE)
writeDataTable(wb, "by_snu1", by_snu1, tableStyle = "TableStyleLight2")
addStyle(wb, "by_snu1", style = pct, cols=c(6,14), rows = 2:(nrow(by_snu1)+1), gridExpand=TRUE)
addStyle(wb, "by_snu1", style = pct2, cols=c(8), rows = 2:(nrow(by_snu1)+1), gridExpand=TRUE)

addWorksheet(wb, "by_psnu",
             gridLines = FALSE)
writeDataTable(wb, "by_psnu", by_psnu, tableStyle = "TableStyleLight2")
addStyle(wb, "by_psnu", style = pct, cols=c(7,15), rows = 2:(nrow(by_psnu)+1), gridExpand=TRUE)
addStyle(wb, "by_psnu", style = pct2, cols=c(9), rows = 2:(nrow(by_psnu)+1), gridExpand=TRUE)

addWorksheet(wb, "by_site",
             gridLines = FALSE)
writeDataTable(wb, "by_site", by_site, tableStyle = "TableStyleLight2")
addStyle(wb, "by_site", style = pct, cols=c(10,18), rows = 2:(nrow(by_site)+1), gridExpand=TRUE)
addStyle(wb, "by_site", style = pct2, cols=c(12), rows = 2:(nrow(by_site)+1), gridExpand=TRUE)

saveWorkbook(wb, file = "C:/Users/cnhantumbo/Documents/R codes/pmtct_cascade.xlsx", overwrite = TRUE)

##############################################
##############################################
##############################################
## PMTCT_EID #################################
##############################################
##############################################
##############################################


#---PMTCT_EID_D------------------------------------------
# PMTCT_EID_DEN REFERENCE OBJECT

pmtct_eid_d <- df_q %>%
  filter(indicator == "PMTCT_EID", disag == "Total Denominator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(pmtct_eid_d = starts_with("Q"))


#---PMTCT_EID_N------------------------------------------
# PMTCT_EID_NUM REFERENCE OBJECT

pmtct_eid_n <- df_q %>%
  filter(indicator == "PMTCT_EID", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(pmtct_eid_n = starts_with("Q"))


#---PMTCT_EID_N <2M------------------------------------------
# PMTCT_EID_NUM <2M REFERENCE OBJECT

pmtct_eid_2m <- df_q %>% #OK
  filter(indicator == "PMTCT_EID", trendssemifine == "<=02 Months", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(pmtct_eid_2m = starts_with("Q"))

#---PMTCT_EID_N <12M------------------------------------------
# PMTCT_EID_NUM 2-12M REFERENCE OBJECT

pmtct_eid_12m <- df_q %>%
  filter(indicator == "PMTCT_EID", trendssemifine == "02 - 12 Months", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(pmtct_eid_12m = starts_with("Q"))


#---PMTCT_HEI_POS <2M------------------------------------------
# PMTCT_HEI_POS <2M REFERENCE OBJECT

pmtct_hei_pos_2m <- df_q %>%
  filter(indicator == "PMTCT_HEI_POS", categoryoptioncomboname == "<= 2 months, Positive", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(pmtct_hei_pos_2m = starts_with("Q"))


#---PMTCT_HEI_POS <2M ART------------------------------------------
# PMTCT_HEI_POS_ART REFERENCE OBJECT

pmtct_hei_pos_2m_art <- df_q %>%
  filter(indicator == "PMTCT_HEI_POS", categoryoptioncomboname == "<= 2 months, Positive, Receiving ART", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(pmtct_hei_pos_2m_art = starts_with("Q"))


#---PMTCT_HEI_POS------------------------------------------
# PMTCT_HEI_POS REFERENCE OBJECT

pmtct_hei_pos <- df_q %>%
  filter(indicator == "PMTCT_HEI_POS", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(pmtct_hei_pos = starts_with("Q"))


#---PMTCT_HEI_POS_ART------------------------------------------
# PMTCT_HEI_POS_ART REFERENCE OBJECT

pmtct_hei_pos_art <- df_q %>%
  filter(indicator == "PMTCT_HEI_POS", categoryoptioncomboname == "2 - 12 months , Positive, Receiving ART" | categoryoptioncomboname == "<= 2 months, Positive, Receiving ART", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(pmtct_hei_pos_art = starts_with("Q"))


#---JOIN ALL------------------------------------------
# JOIN ALL CASCADE INDICATOR OBJECTS INTO ONE DATAFRAME

pc_eid <- pmtct_eid_d %>%
  full_join(pmtct_eid_n, by = "site_id") %>%
  full_join(pmtct_eid_2m, by = "site_id") %>%
  full_join(pmtct_hei_pos_2m, by = "site_id") %>%
  full_join(pmtct_hei_pos_2m_art, by = "site_id") %>%
  full_join(pmtct_hei_pos, by = "site_id") %>%
  full_join(pmtct_hei_pos_art, by = "site_id") %>%
  left_join(ajuda, by = "site_id") %>%
  replace_na(list(pmtct_eid_d = 0,
                  pmtct_eid_n = 0,
                  pmtct_eid_2m = 0,
                  pmtct_hei_pos_2m = 0,
                  pmtct_hei_pos_2m_art = 0,
                  pmtct_hei_pos = 0,
                  pmtct_hei_pos_art = 0,
                  ajuda = 0)
  ) %>%
  select(1:7, ajuda, everything())


#---ATTRIBUTE INFORMATION------------------------------------------
# ASSURE ALL OBSERVATIONS HAVE SITE ATTRIBUTE INFORMATION

pc_eid$agency.x <- as.character(pc_eid$agency.x)
pc_eid$agency.y <- as.character(pc_eid$agency.y)
pc_eid$agency.x.x <- as.character(pc_eid$agency.x.x)
pc_eid$agency.y.y <- as.character(pc_eid$agency.y.y)
pc_eid$agency.x.x.x <- as.character(pc_eid$agency.x.x.x)
pc_eid$agency.y.y.y <- as.character(pc_eid$agency.y.y.y)


pc_eid_temp <- pc_eid %>%
  mutate(agency.x = 
           if_else(is.na(agency.x), agency.y, 
                   if_else(is.na(agency.x), agency.x.x, 
                           if_else(is.na(agency.x), agency.y.y, agency.x)
                   )
           )
  ) %>%
  mutate(partner.x = 
           if_else(is.na(partner.x), partner.y, 
                   if_else(is.na(partner.x), partner.x.x, 
                           if_else(is.na(partner.x), partner.y.y, partner.x)
                   )
           )
  ) %>%
  mutate(snu1.x = 
           if_else(is.na(snu1.x), snu1.y, 
                   if_else(is.na(snu1.x), snu1.x.x, 
                           if_else(is.na(snu1.x), snu1.y.y, snu1.x)
                   )
           )
  ) %>%
  mutate(psnu.x = 
           if_else(is.na(psnu.x), psnu.y, 
                   if_else(is.na(psnu.x), psnu.x.x, 
                           if_else(is.na(psnu.x), psnu.y.y, psnu.x)
                   )
           )
  ) %>%
  mutate(site.x = 
           if_else(is.na(site.x), site.y, 
                   if_else(is.na(site.x), site.x.x, 
                           if_else(is.na(site.x), site.y.y, site.x)
                   )
           )
  ) %>%
  rename(agency = agency.x,
         partner = partner.x,
         snu1 = snu1.x,
         psnu = psnu.x,
         site = site.x
  ) %>%
  select(-ends_with(".x"), -ends_with(".y"))

#---CASADE BY SITE------------------------------------------
# PMTCT CASADE BY SITE

by_site <- pc_eid_temp %>%
  mutate(per_pmtct_eid = pmtct_eid_n / pmtct_eid_d,
         per_pmtct_eid_2m = pmtct_eid_2m / pmtct_eid_d,
         per_pmtct_hei_pos_2m = pmtct_hei_pos_2m / pmtct_eid_2m,
         per_pmtct_hei_pos_2m_art = pmtct_hei_pos_2m_art / pmtct_hei_pos_2m,
         per_pmtct_hei_pos = pmtct_hei_pos / pmtct_eid_n,
         per_pmtct_hei_pos_art = pmtct_hei_pos_art / pmtct_hei_pos
  ) %>%
  select(site_id, agency, partner, snu1, psnu, site, ajuda, pmtct_eid_d, pmtct_eid_n, per_pmtct_eid, pmtct_eid_2m, per_pmtct_eid_2m, pmtct_hei_pos_2m, per_pmtct_hei_pos_2m, pmtct_hei_pos_2m_art, per_pmtct_hei_pos_2m_art, pmtct_hei_pos, per_pmtct_hei_pos, pmtct_hei_pos_art, per_pmtct_hei_pos_art)

#---CASADE BY DISTRICT------------------------------------------
# PMTCT CASADE BY DISTRICT

by_psnu <- pc_eid_temp %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(pmtct_eid_d = sum(pmtct_eid_d, na.rm = TRUE),
            pmtct_eid_n = sum(pmtct_eid_n, na.rm = TRUE),
            pmtct_eid_2m = sum(pmtct_eid_2m, na.rm = TRUE),
            pmtct_hei_pos_2m = sum(pmtct_hei_pos_2m, na.rm = TRUE),
            pmtct_hei_pos_2m_art = sum(pmtct_hei_pos_2m_art, na.rm = TRUE),
            pmtct_hei_pos = sum(pmtct_hei_pos, na.rm = TRUE),
            pmtct_hei_pos_art = sum(pmtct_hei_pos_art, na.rm = TRUE),
  )%>%
  ungroup() %>%
  mutate(per_pmtct_eid = pmtct_eid_n / pmtct_eid_d,
         per_pmtct_eid_2m = pmtct_eid_2m / pmtct_eid_d,
         per_pmtct_hei_pos_2m = pmtct_hei_pos_2m / pmtct_eid_2m,
         per_pmtct_hei_pos_2m_art = pmtct_hei_pos_2m_art / pmtct_hei_pos_2m,
         per_pmtct_hei_pos = pmtct_hei_pos / pmtct_eid_n,
         per_pmtct_hei_pos_art = pmtct_hei_pos_art / pmtct_hei_pos
  ) %>%
  select(agency, partner, snu1, psnu, pmtct_eid_d, pmtct_eid_n, per_pmtct_eid, pmtct_eid_2m, per_pmtct_eid_2m, pmtct_hei_pos_2m, per_pmtct_hei_pos_2m, pmtct_hei_pos_2m_art, per_pmtct_hei_pos_2m_art, pmtct_hei_pos, per_pmtct_hei_pos, pmtct_hei_pos_art, per_pmtct_hei_pos_art)

#---CASADE BY PROVINCE------------------------------------------
# PMTCT CASADE BY PROVINCE

by_snu1 <- pc_eid_temp %>%
  group_by(agency, partner, snu1) %>%
  summarize(pmtct_eid_d = sum(pmtct_eid_d, na.rm = TRUE),
            pmtct_eid_n = sum(pmtct_eid_n, na.rm = TRUE),
            pmtct_eid_2m = sum(pmtct_eid_2m, na.rm = TRUE),
            pmtct_hei_pos_2m = sum(pmtct_hei_pos_2m, na.rm = TRUE),
            pmtct_hei_pos_2m_art = sum(pmtct_hei_pos_2m_art, na.rm = TRUE),
            pmtct_hei_pos = sum(pmtct_hei_pos, na.rm = TRUE),
            pmtct_hei_pos_art = sum(pmtct_hei_pos_art, na.rm = TRUE),
  )%>%
  ungroup() %>%
  mutate(per_pmtct_eid = pmtct_eid_n / pmtct_eid_d,
         per_pmtct_eid_2m = pmtct_eid_2m / pmtct_eid_d,
         per_pmtct_hei_pos_2m = pmtct_hei_pos_2m / pmtct_eid_2m,
         per_pmtct_hei_pos_2m_art = pmtct_hei_pos_2m_art / pmtct_hei_pos_2m,
         per_pmtct_hei_pos = pmtct_hei_pos / pmtct_eid_n,
         per_pmtct_hei_pos_art = pmtct_hei_pos_art / pmtct_hei_pos
  ) %>%
  select(agency, partner, snu1, pmtct_eid_d, pmtct_eid_n, per_pmtct_eid, pmtct_eid_2m, per_pmtct_eid_2m, pmtct_hei_pos_2m, per_pmtct_hei_pos_2m, pmtct_hei_pos_2m_art, per_pmtct_hei_pos_2m_art, pmtct_hei_pos, per_pmtct_hei_pos, pmtct_hei_pos_art, per_pmtct_hei_pos_art)


#---CASADE BY PARTNER/PROVINCE/AJUDA------------------------------------------
# PMTCT CASADE BY PARTNER/PROVINCE/AJUDA

by_partner_snu1 <- pc_eid_temp %>%
  group_by(agency, partner, snu1, ajuda) %>%
  summarize(pmtct_eid_d = sum(pmtct_eid_d, na.rm = TRUE),
            pmtct_eid_n = sum(pmtct_eid_n, na.rm = TRUE),
            pmtct_eid_2m = sum(pmtct_eid_2m, na.rm = TRUE),
            pmtct_hei_pos_2m = sum(pmtct_hei_pos_2m, na.rm = TRUE),
            pmtct_hei_pos_2m_art = sum(pmtct_hei_pos_2m_art, na.rm = TRUE),
            pmtct_hei_pos = sum(pmtct_hei_pos, na.rm = TRUE),
            pmtct_hei_pos_art = sum(pmtct_hei_pos_art, na.rm = TRUE),
  )%>%
  ungroup() %>%
  mutate(per_pmtct_eid = pmtct_eid_n / pmtct_eid_d,
         per_pmtct_eid_2m = pmtct_eid_2m / pmtct_eid_d,
         per_pmtct_hei_pos_2m = pmtct_hei_pos_2m / pmtct_eid_2m,
         per_pmtct_hei_pos_2m_art = pmtct_hei_pos_2m_art / pmtct_hei_pos_2m,
         per_pmtct_hei_pos = pmtct_hei_pos / pmtct_eid_n,
         per_pmtct_hei_pos_art = pmtct_hei_pos_art / pmtct_hei_pos
  ) %>%
  select(agency, partner, snu1, ajuda, pmtct_eid_d, pmtct_eid_n, per_pmtct_eid, pmtct_eid_2m, per_pmtct_eid_2m, pmtct_hei_pos_2m, per_pmtct_hei_pos_2m, pmtct_hei_pos_2m_art, per_pmtct_hei_pos_2m_art, pmtct_hei_pos, per_pmtct_hei_pos, pmtct_hei_pos_art, per_pmtct_hei_pos_art)


#---CASADE BY PARTNER------------------------------------------
# PMTCT CASADE BY PARTNER

by_partner <- pc_eid_temp %>%
  group_by(agency, partner) %>%
  summarize(pmtct_eid_d = sum(pmtct_eid_d, na.rm = TRUE),
            pmtct_eid_n = sum(pmtct_eid_n, na.rm = TRUE),
            pmtct_eid_2m = sum(pmtct_eid_2m, na.rm = TRUE),
            pmtct_hei_pos_2m = sum(pmtct_hei_pos_2m, na.rm = TRUE),
            pmtct_hei_pos_2m_art = sum(pmtct_hei_pos_2m_art, na.rm = TRUE),
            pmtct_hei_pos = sum(pmtct_hei_pos, na.rm = TRUE),
            pmtct_hei_pos_art = sum(pmtct_hei_pos_art, na.rm = TRUE),
  )%>%
  ungroup() %>%
  mutate(per_pmtct_eid = pmtct_eid_n / pmtct_eid_d,
         per_pmtct_eid_2m = pmtct_eid_2m / pmtct_eid_d,
         per_pmtct_hei_pos_2m = pmtct_hei_pos_2m / pmtct_eid_2m,
         per_pmtct_hei_pos_2m_art = pmtct_hei_pos_2m_art / pmtct_hei_pos_2m,
         per_pmtct_hei_pos = pmtct_hei_pos / pmtct_eid_n,
         per_pmtct_hei_pos_art = pmtct_hei_pos_art / pmtct_hei_pos
  ) %>%
  select(agency, partner, pmtct_eid_d, pmtct_eid_n, per_pmtct_eid, pmtct_eid_2m, per_pmtct_eid_2m, pmtct_hei_pos_2m, per_pmtct_hei_pos_2m, pmtct_hei_pos_2m_art, per_pmtct_hei_pos_2m_art, pmtct_hei_pos, per_pmtct_hei_pos, pmtct_hei_pos_art, per_pmtct_hei_pos_art)

#---CASADE BY AGENCY------------------------------------------
# PMTCT CASADE BY AGENCY

by_agency <- pc_eid_temp %>%
  group_by(agency) %>%
  summarize(pmtct_eid_d = sum(pmtct_eid_d, na.rm = TRUE),
            pmtct_eid_n = sum(pmtct_eid_n, na.rm = TRUE),
            pmtct_eid_2m = sum(pmtct_eid_2m, na.rm = TRUE),
            pmtct_hei_pos_2m = sum(pmtct_hei_pos_2m, na.rm = TRUE),
            pmtct_hei_pos_2m_art = sum(pmtct_hei_pos_2m_art, na.rm = TRUE),
            pmtct_hei_pos = sum(pmtct_hei_pos, na.rm = TRUE),
            pmtct_hei_pos_art = sum(pmtct_hei_pos_art, na.rm = TRUE),
  )%>%
  ungroup() %>%
  mutate(per_pmtct_eid = pmtct_eid_n / pmtct_eid_d,
         per_pmtct_eid_2m = pmtct_eid_2m / pmtct_eid_d,
         per_pmtct_hei_pos_2m = pmtct_hei_pos_2m / pmtct_eid_2m,
         per_pmtct_hei_pos_2m_art = pmtct_hei_pos_2m_art / pmtct_hei_pos_2m,
         per_pmtct_hei_pos = pmtct_hei_pos / pmtct_eid_n,
         per_pmtct_hei_pos_art = pmtct_hei_pos_art / pmtct_hei_pos
  ) %>%
  select(agency, pmtct_eid_d, pmtct_eid_n, per_pmtct_eid, pmtct_eid_2m, per_pmtct_eid_2m, pmtct_hei_pos_2m, per_pmtct_hei_pos_2m, pmtct_hei_pos_2m_art, per_pmtct_hei_pos_2m_art, pmtct_hei_pos, per_pmtct_hei_pos, pmtct_hei_pos_art, per_pmtct_hei_pos_art)

#---CASADE BY AGENCY------------------------------------------ 
# PMTCT CASADE BY AGENCY

by_ou <- by_agency %>%
  mutate(agency = "Moz") %>%
  rename(ou = agency) %>%
  group_by(ou) %>%
  summarize(pmtct_eid_d = sum(pmtct_eid_d, na.rm = TRUE),
            pmtct_eid_n = sum(pmtct_eid_n, na.rm = TRUE),
            pmtct_eid_2m = sum(pmtct_eid_2m, na.rm = TRUE),
            pmtct_hei_pos_2m = sum(pmtct_hei_pos_2m, na.rm = TRUE),
            pmtct_hei_pos_2m_art = sum(pmtct_hei_pos_2m_art, na.rm = TRUE),
            pmtct_hei_pos = sum(pmtct_hei_pos, na.rm = TRUE),
            pmtct_hei_pos_art = sum(pmtct_hei_pos_art, na.rm = TRUE),
  )%>%
  ungroup() %>%
  mutate(per_pmtct_eid = pmtct_eid_n / pmtct_eid_d,
         per_pmtct_eid_2m = pmtct_eid_2m / pmtct_eid_d,
         per_pmtct_hei_pos_2m = pmtct_hei_pos_2m / pmtct_eid_2m,
         per_pmtct_hei_pos_2m_art = pmtct_hei_pos_2m_art / pmtct_hei_pos_2m,
         per_pmtct_hei_pos = pmtct_hei_pos / pmtct_eid_n,
         per_pmtct_hei_pos_art = pmtct_hei_pos_art / pmtct_hei_pos
  ) %>%
  select(ou, pmtct_eid_d, pmtct_eid_n, per_pmtct_eid, pmtct_eid_2m, per_pmtct_eid_2m, pmtct_hei_pos_2m, per_pmtct_hei_pos_2m, pmtct_hei_pos_2m_art, per_pmtct_hei_pos_2m_art, pmtct_hei_pos, per_pmtct_hei_pos, pmtct_hei_pos_art, per_pmtct_hei_pos_art)


#---CASADE BY AJUDA------------------------------------------
# PMTCT CASADE BY AJUDA

by_ajuda <- pc_eid_temp %>%
  group_by(ajuda) %>%
  summarize(pmtct_eid_d = sum(pmtct_eid_d, na.rm = TRUE),
            pmtct_eid_n = sum(pmtct_eid_n, na.rm = TRUE),
            pmtct_eid_2m = sum(pmtct_eid_2m, na.rm = TRUE),
            pmtct_hei_pos_2m = sum(pmtct_hei_pos_2m, na.rm = TRUE),
            pmtct_hei_pos_2m_art = sum(pmtct_hei_pos_2m_art, na.rm = TRUE),
            pmtct_hei_pos = sum(pmtct_hei_pos, na.rm = TRUE),
            pmtct_hei_pos_art = sum(pmtct_hei_pos_art, na.rm = TRUE),
  )%>%
  ungroup() %>%
  mutate(per_pmtct_eid = pmtct_eid_n / pmtct_eid_d,
         per_pmtct_eid_2m = pmtct_eid_2m / pmtct_eid_d,
         per_pmtct_hei_pos_2m = pmtct_hei_pos_2m / pmtct_eid_2m,
         per_pmtct_hei_pos_2m_art = pmtct_hei_pos_2m_art / pmtct_hei_pos_2m,
         per_pmtct_hei_pos = pmtct_hei_pos / pmtct_eid_n,
         per_pmtct_hei_pos_art = pmtct_hei_pos_art / pmtct_hei_pos
  ) %>%
  select(ajuda, pmtct_eid_d, pmtct_eid_n, per_pmtct_eid, pmtct_eid_2m, per_pmtct_eid_2m, pmtct_hei_pos_2m, per_pmtct_hei_pos_2m, pmtct_hei_pos_2m_art, per_pmtct_hei_pos_2m_art, pmtct_hei_pos, per_pmtct_hei_pos, pmtct_hei_pos_art, per_pmtct_hei_pos_art)


#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH PMTCT CASCADE OUTPUT

wb <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")
pct2 <- createStyle(numFmt = "0.0%", textDecoration = "italic")

addWorksheet(wb, "by_ou",
             gridLines = FALSE)
writeDataTable(wb, "by_ou", by_ou, tableStyle = "TableStyleLight2")
addStyle(wb, "by_ou", style = pct, cols=c(4,6,10,14), rows = 2:(nrow(by_ou)+1), gridExpand=TRUE)
addStyle(wb, "by_ou", style = pct2, cols=c(8,12), rows = 2:(nrow(by_ou)+1), gridExpand=TRUE)
setColWidths(wb, "by_ou", cols = c(1), widths = "auto")

addWorksheet(wb, "by_agency",
             gridLines = FALSE)
writeDataTable(wb, "by_agency", by_agency, tableStyle = "TableStyleLight2")
addStyle(wb, "by_agency", style = pct, cols=c(4,6,10,14), rows = 2:(nrow(by_agency)+1), gridExpand=TRUE)
addStyle(wb, "by_agency", style = pct2, cols=c(8,12), rows = 2:(nrow(by_agency)+1), gridExpand=TRUE)
setColWidths(wb, "by_agency", cols = c(1), widths = "auto")

addWorksheet(wb, "by_partner",
             gridLines = FALSE)
writeDataTable(wb, "by_partner", by_partner, tableStyle = "TableStyleLight2")
addStyle(wb, "by_partner", style = pct, cols=c(5,7,11,15), rows = 2:(nrow(by_partner)+1), gridExpand=TRUE)
addStyle(wb, "by_partner", style = pct2, cols=c(9,13), rows = 2:(nrow(by_partner)+1), gridExpand=TRUE)
setColWidths(wb, "by_partner", cols = c(1:2), widths = "auto")

addWorksheet(wb, "by_snu1",
             gridLines = FALSE)
writeDataTable(wb, "by_snu1", by_snu1, tableStyle = "TableStyleLight2")
addStyle(wb, "by_snu1", style = pct, cols=c(6,8,12,16), rows = 2:(nrow(by_snu1)+1), gridExpand=TRUE)
addStyle(wb, "by_snu1", style = pct2, cols=c(10,14), rows = 2:(nrow(by_snu1)+1), gridExpand=TRUE)
setColWidths(wb, "by_snu1", cols = c(1:3), widths = "auto")

addWorksheet(wb, "by_psnu",
             gridLines = FALSE)
writeDataTable(wb, "by_psnu", by_psnu, tableStyle = "TableStyleLight2")
addStyle(wb, "by_psnu", style = pct, cols=c(7,9,13,17), rows = 2:(nrow(by_psnu)+1), gridExpand=TRUE)
addStyle(wb, "by_psnu", style = pct2, cols=c(11,15), rows = 2:(nrow(by_psnu)+1), gridExpand=TRUE)
setColWidths(wb, "by_psnu", cols = c(1:4), widths = "auto")

addWorksheet(wb, "by_site",
             gridLines = FALSE)
writeDataTable(wb, "by_site", by_site, tableStyle = "TableStyleLight2")
addStyle(wb, "by_site", style = pct, cols=c(10,12,16,20), rows = 2:(nrow(by_site)+1), gridExpand=TRUE)
addStyle(wb, "by_site", style = pct2, cols=c(14,18), rows = 2:(nrow(by_site)+1), gridExpand=TRUE)
setColWidths(wb, "by_site", cols = c(1:6), widths = "auto")

saveWorkbook(wb, file = "C:/Users/cnhantumbo/Documents/R codes/pmtct_eid_cascade.xlsx", overwrite = TRUE)

##############################################
##############################################
##############################################
## TX SUMMARY ################################
##############################################
##############################################
##############################################

#---------------------------------------------
## PREPARE TX_CURR DATASET
tx_curr_base <- df_a %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", 
         partner != "Food and Nutrition Technical Assistance III (FANTA-III)",
         partner != "Passos",
         partner != "YouthPower Implementation -  Task Order 1") %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, target) %>%
  left_join(ajuda, by = "site_id") %>%
  replace_na(list(ajuda = 0)
  )

#---------------------------------------------
## PREPARE TX_NEW DATASET
tx_new_base <- df_a %>%
  filter(indicator == "TX_NEW", disag == "Total Numerator",
         partner != "Food and Nutrition Technical Assistance III (FANTA-III)",
         partner != "Passos",
         partner != "YouthPower Implementation -  Task Order 1") %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, target) %>%
  left_join(ajuda, by = "site_id") %>%
  replace_na(list(ajuda = 0)
  )

#---------------------------------------------
##OU LEVEL TX_CURR PERFORMANCE
tx_curr_ou <- tx_curr_base %>%
  mutate(agency = "Moz") %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup()

#---------------------------------------------
##AJUDA LEVEL TX_CURR PERFORMANCE
tx_curr_ajuda <- tx_curr_base %>%
  group_by(ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup() %>%
  arrange(ajuda)

#---------------------------------------------
##AGENCY LEVEL TX_CURR PERFORMANCE
tx_curr_agency <- tx_curr_base %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup() %>%
  arrange(agency)

#---------------------------------------------
##PARTNER LEVEL TX_CURR PERFORMANCE
tx_curr_partner <- tx_curr_base %>%
  group_by(agency, partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup() %>%
  arrange(agency, partner)

#---------------------------------------------
##PROVINCE LEVEL TX_CURR PERFORMANCE
tx_curr_snu1 <- tx_curr_base %>%
  group_by(agency, partner, snu1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1)

#---------------------------------------------
##DISTRICT LEVEL TX_CURR PERFORMANCE
tx_curr_psnu <- tx_curr_base %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu)


#---------------------------------------------
##SITE LEVEL TX_CURR PERFORMANCE
tx_curr_site <- tx_curr_base %>%
  group_by(site_id, agency, partner, snu1, psnu, site, ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu, site)

#---------------------------------------------
##OU LEVEL TX_NEW PERFORMANCE
tx_new_ou <- tx_new_base %>%
  mutate(agency = "Moz") %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()

#---------------------------------------------
##AJUDA LEVEL TX_CURR PERFORMANCE
tx_new_ajuda <- tx_new_base %>%
  group_by(ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(ajuda)

#---------------------------------------------
##AGENCY LEVEL TX_CURR PERFORMANCE
tx_new_agency <- tx_new_base %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(agency)

#---------------------------------------------
##PARTNER LEVEL TX_CURR PERFORMANCE
tx_new_partner <- tx_new_base %>%
  group_by(agency, partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(agency, partner)

#---------------------------------------------
##PROVINCE LEVEL TX_CURR PERFORMANCE
tx_new_snu1 <- tx_new_base %>%
  group_by(agency, partner, snu1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(agency, partner, snu1)

#---------------------------------------------
##DISTRICT LEVEL TX_CURR PERFORMANCE
tx_new_psnu <- tx_new_base %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu)

#---------------------------------------------
##SITE LEVEL TX_CURR PERFORMANCE
tx_new_site <- tx_new_base %>%
  group_by(site_id, agency, partner, snu1, psnu, site, ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu)


#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_overall <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")
pct2 <- createStyle(numFmt = "0.0%", textDecoration = "italic")

wb_overall <- createWorkbook()
addWorksheet(wb_overall, "tx_curr_ou",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_ou", tx_curr_ou, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_ou", style = pct2, cols=c(7:9), rows = 2:(nrow(tx_curr_ou)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_ajuda",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_ajuda", tx_curr_ajuda, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_ajuda", style = pct2, cols=c(7:9), rows = 2:(nrow(tx_curr_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_agency",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_agency", tx_curr_agency, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_agency", style = pct2, cols=c(7:9), rows = 2:(nrow(tx_curr_agency)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_partner",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_partner", tx_curr_partner, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_partner", style = pct2, cols=c(8:10), rows = 2:(nrow(tx_curr_partner)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_snu1",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_snu1", tx_curr_snu1, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_snu1", style = pct2, cols=c(9:11), rows = 2:(nrow(tx_curr_snu1)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_psnu",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_psnu", tx_curr_psnu, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_psnu", style = pct2, cols=c(10:12), rows = 2:(nrow(tx_curr_psnu)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_site",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_site", tx_curr_site, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_site", style = pct2, cols=c(13:15), rows = 2:(nrow(tx_curr_site)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_ou",
             tabColour = "#808080",
             gridLines = FALSE)

writeDataTable(wb_overall, "tx_new_ou", tx_new_ou, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_new_ou", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_ou)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_ajuda",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_ajuda", tx_new_ajuda, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_new_ajuda", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_agency",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_agency", tx_new_agency, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_new_agency", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_agency)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_partner",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_partner", tx_new_partner, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_new_partner", style = pct, cols=c(8:12), rows = 2:(nrow(tx_new_partner)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_snu1",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_snu1", tx_new_snu1, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_new_snu1", style = pct, cols=c(9:13), rows = 2:(nrow(tx_new_snu1)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_psnu",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_psnu", tx_new_psnu, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_new_psnu", style = pct, cols=c(10:14), rows = 2:(nrow(tx_new_psnu)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_site",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_site", tx_new_site, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_new_site", style = pct, cols=c(13:17), rows = 2:(nrow(tx_new_site)+1), gridExpand=TRUE)


#######################
#######################
#--- PEDIATRIC SUMMARY
#######################
#######################

#---------------------------------------------
## PREPARE TX_CURR DATASET
tx_curr_base <- df_a %>%
  filter(indicator == "TX_CURR", trendscoarse == "<15",
         partner != "Food and Nutrition Technical Assistance III (FANTA-III)",
         partner != "Passos",
         partner != "YouthPower Implementation -  Task Order 1") %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, target) %>%
  left_join(ajuda, by = "site_id") %>%
  replace_na(list(ajuda = 0)
  )

#---------------------------------------------
## PREPARE TX_NEW DATASET
tx_new_base <- df_a %>%
  filter(indicator == "TX_NEW", trendscoarse == "<15",
         partner != "Food and Nutrition Technical Assistance III (FANTA-III)",
         partner != "Passos",
         partner != "YouthPower Implementation -  Task Order 1") %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, target) %>%
  left_join(ajuda, by = "site_id") %>%
  replace_na(list(ajuda = 0)
  )

#---------------------------------------------
##OU LEVEL TX_CURR PERFORMANCE
tx_curr_ou <- tx_curr_base %>%
  mutate(agency = "Moz") %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup()

#---------------------------------------------
##AJUDA LEVEL TX_CURR PERFORMANCE
tx_curr_ajuda <- tx_curr_base %>%
  group_by(ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup()

#---------------------------------------------
##AGENCY LEVEL TX_CURR PERFORMANCE
tx_curr_agency <- tx_curr_base %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup %>%
  arrange(agency)

#---------------------------------------------
##PARTNER LEVEL TX_CURR PERFORMANCE
tx_curr_partner <- tx_curr_base %>%
  group_by(agency, partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup() %>%
  arrange(agency, partner)

#---------------------------------------------
##PROVINCE LEVEL TX_CURR PERFORMANCE
tx_curr_snu1 <- tx_curr_base %>%
  group_by(agency, partner, snu1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1)

#---------------------------------------------
##DISTRICT LEVEL TX_CURR PERFORMANCE
tx_curr_psnu <- tx_curr_base %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu)

#---------------------------------------------
##SITE LEVEL TX_CURR PERFORMANCE
tx_curr_site <- tx_curr_base %>%
  group_by(site_id, agency, partner, snu1, psnu, site, ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = (Q2-Q1)/Q1, Q3_netnew = (Q3-Q2)/Q2, Q4_netnew = (Q4-Q3)/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu, site)

#---------------------------------------------
##OU LEVEL TX_NEW PERFORMANCE
tx_new_ou <- tx_new_base %>%
  mutate(agency = "Moz") %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()

#---------------------------------------------
##AJUDA LEVEL TX_CURR PERFORMANCE
tx_new_ajuda <- tx_new_base %>%
  group_by(ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()

#---------------------------------------------
##AGENCY LEVEL TX_CURR PERFORMANCE
tx_new_agency <- tx_new_base %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(agency)

#---------------------------------------------
##PARTNER LEVEL TX_CURR PERFORMANCE
tx_new_partner <- tx_new_base %>%
  group_by(agency, partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(agency, partner)

#---------------------------------------------
##PROVINCE LEVEL TX_CURR PERFORMANCE
tx_new_snu1 <- tx_new_base %>%
  group_by(agency, partner, snu1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(agency, partner, snu1)

#---------------------------------------------
##DISTRICT LEVEL TX_CURR PERFORMANCE
tx_new_psnu <- tx_new_base %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu)

#---------------------------------------------
##SITE LEVEL TX_CURR PERFORMANCE
tx_new_site <- tx_new_base %>%
  group_by(site_id, agency, partner, snu1, psnu, site, ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q1_vs_Target = (Q1)/(target / 4))%>%
  mutate(Q2_vs_Target = (Q2)/(target / 4))%>%
  mutate(Q3_vs_Target = (Q3)/(target / 4))%>%
  mutate(Q4_vs_Target = (Q4)/(target / 4))%>%
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu, site)


#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_pediatrico <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")
pct2 <- createStyle(numFmt = "0.0%", textDecoration = "italic")

wb_pediatrico <- createWorkbook()
addWorksheet(wb_pediatrico, "tx_curr_ou",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_ou", tx_curr_ou, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_ou", style = pct2, cols=c(7:9), rows = 2:(nrow(tx_curr_ou)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_curr_ajuda",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_ajuda", tx_curr_ajuda, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_ajuda", style = pct2, cols=c(7:9), rows = 2:(nrow(tx_curr_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_curr_agency",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_agency", tx_curr_agency, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_agency", style = pct2, cols=c(7:9), rows = 2:(nrow(tx_curr_agency)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_curr_partner",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_partner", tx_curr_partner, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_partner", style = pct2, cols=c(8:10), rows = 2:(nrow(tx_curr_partner)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_curr_snu1",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_snu1", tx_curr_snu1, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_snu1", style = pct2, cols=c(9:11), rows = 2:(nrow(tx_curr_snu1)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_curr_psnu",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_psnu", tx_curr_psnu, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_psnu", style = pct2, cols=c(10:12), rows = 2:(nrow(tx_curr_psnu)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_ou",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_ou", tx_new_ou, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_new_ou", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_ou)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_ajuda",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_ajuda", tx_new_ajuda, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_new_ajuda", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_agency",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_agency", tx_new_agency, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_new_agency", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_agency)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_partner",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_partner", tx_new_partner, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_new_partner", style = pct, cols=c(8:12), rows = 2:(nrow(tx_new_partner)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_snu1",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_snu1", tx_new_snu1, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_new_snu1", style = pct, cols=c(9:13), rows = 2:(nrow(tx_new_snu1)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_psnu",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_psnu", tx_new_psnu, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_new_psnu", style = pct, cols=c(10:14), rows = 2:(nrow(tx_new_psnu)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_site",
             tabColour = "#808080",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_site", tx_new_site, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_new_site", style = pct, cols=c(13:17), rows = 2:(nrow(tx_new_site)+1), gridExpand=TRUE)



#---------------------------------------------
##PRINT WORKBOOKS
saveWorkbook(wb_overall, file = "C:/Users/cnhantumbo/Documents/R codes/tx_summary.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
saveWorkbook(wb_pediatrico, file = "C:/Users/cnhantumbo/Documents/R codes/tx_summary_ped.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!

##############################################
##############################################
##############################################
## TX CASCADE ################################
##############################################
##############################################
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df_q %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", target > 0) %>%
  select(site_id, target) %>%
  group_by(site_id) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df_q %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = starts_with("Q")) 

#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df_q %>%
  filter(indicator == "HTS_TST",
         disag == "Total Numerator",
         partner == "CHASS" | partner == "FGH" | partner == "ICAP" | partner == "CCS" | partner == "ARIEL" | partner == "EGPAF"| partner == "ECHO",
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

hts_tst_pos <- df_q %>%
  filter(indicator == "HTS_TST_POS",
         disag == "Total Numerator",
         partner == "CHASS" | partner == "FGH" | partner == "ICAP" | partner == "CCS" | partner == "ARIEL" | partner == "EGPAF"| partner == "ECHO",
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

tx_new <- df_q %>%
  filter(indicator == "TX_NEW", disag == "Total Numerator", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = starts_with("Q"))

#---------------------------------------------
# TX_PVLS DENOMINATOR REFERENCE OBJECT

tx_pvls_d <- df_q %>%
  filter(indicator == "TX_PVLS", disag == "Total Denominator", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = starts_with("Q"))

#---------------------------------------------
# TX_PVLS NUMERATOR REFERENCE OBJECT

tx_pvls_n <- df_q %>%
  filter(indicator == "TX_PVLS", disag == "Total Numerator", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = starts_with("Q"))

#---------------------------------------------
# JOIN REFERENCE OBJECTS & CREATE SITE CASCADE

cc_site <- tx_curr %>%
  left_join(tx_curr_target, by = "site_id") %>%
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
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(site_id, agency, partner, snu1, psnu, site, ajuda, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls) %>%
  arrange(agency, partner, snu1, psnu, site)


#---------------------------------------------
##CLEAN-UP R ENVIRONMENT BY REMOVING UNNEEDED OBJECTS

rm(tx_curr_target, tx_curr, tx_new, tx_pvls_d, tx_pvls_n, hts_tst, hts_tst_pos)

#---------------------------------------------
## CREATE AJUDA PHASE CASCADE

cc_ajuda <- cc_site %>%
  group_by(ajuda) %>%
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
  select(ajuda, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


##############################################
# CREATE ABOVE-SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df_q %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", target > 0) %>%
  select(partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, target) %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# HTS REFERENCE OBJECT

hts_tst <- df_q %>%
  filter(indicator == "HTS_TST", disag == "Total Numerator", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst = "Q")

cc_psnu <- tx_curr_target %>%
  left_join(hts_tst, by = "psnu") %>%
  replace_na(list(hts_tst = 0))

#---------------------------------------------
# HTS_POS REFERENCE OBJECT

hts_tst_pos <- df_q %>%
  filter(indicator == "HTS_TST_POS", disag == "Total Numerator", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst_pos = "Q")

cc_psnu <- cc_psnu %>%
  left_join(hts_tst_pos, by = "psnu") %>%
  replace_na(list(hts_tst_pos = 0))

#---------------------------------------------
# TX_NEW REFERENCE OBJECT

tx_new <- df_q %>%
  filter(indicator == "TX_NEW", disag == "Total Numerator", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_new, by = "psnu") %>%
  replace_na(list(tx_new = 0))

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df_q %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_curr, by = "psnu") %>%
  replace_na(list(tx_curr = 0))

#---------------------------------------------
# TX_PLVS_DEN REFERENCE OBJECT

tx_pvls_d <- df_q %>%
  filter(indicator == "TX_PVLS", disag == "Total Denominator", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_pvls_d, by = "psnu") %>%
  replace_na(list(tx_pvls_d = 0))

#---------------------------------------------
# TX_PLVS_NUM REFERENCE OBJECT

tx_pvls_n <- df_q %>%
  filter(indicator == "TX_PVLS", disag == "Total Numerator", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = "Q") 


##############################################
# CREATE ABOVE-SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
## CREATE DISTRICT CLINICAL CASCADE OBJECT

cc_psnu <- cc_psnu %>%
  left_join(tx_pvls_n, by = "psnu") %>%
  replace_na(list(tx_pvls_n = 0)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_linkage = tx_new / hts_tst_pos,
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(agency, partner, snu1, psnu, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

#---------------------------------------------
## CREATE PROVINCIAL CLINICAL CASCADE OBJECT

cc_snu1 <- cc_psnu %>%
  group_by(agency, partner, snu1) %>%
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
  select(agency, partner, snu1, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

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
## CREATE OU CLINICAL CASCADE OBJECT

cc_ou <- cc_psnu %>%
  mutate(agency = "Moz") %>%
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

##############################################
# CREATE OBJECT FOR OVERALL RESULTS
##############################################

cc_ou_overall <- cc_ou %>%
  select(-agency) %>%
  mutate(resulttype = "Overall") %>%
  select(resulttype, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


##############################################
# GENERATE EXCEL OUTPUTS
##############################################

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_overall <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")
pct2 <- createStyle(numFmt = "0.0%", textDecoration = "italic")

addWorksheet(wb_overall, "cc_ou",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_ou", cc_ou, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "cc_ou", style = pct, cols=c(7,9,11,13), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)
addStyle(wb_overall, "cc_ou", style = pct2, cols=c(5), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "cc_ajuda",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_ajuda", cc_ajuda, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "cc_ajuda", style = pct, cols=c(7,9,11,13), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)
addStyle(wb_overall, "cc_ajuda", style = pct2, cols=c(5), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "cc_agency",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_agency", cc_agency, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "cc_agency", style = pct, cols=c(7,9,11,13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)
addStyle(wb_overall, "cc_agency", style = pct2, cols=c(5), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "cc_partner",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_partner", cc_partner, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "cc_partner", style = pct, cols=c(8,10,12,14), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)
addStyle(wb_overall, "cc_partner", style = pct2, cols=c(6), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "cc_snu1",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_snu1", cc_snu1, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "cc_snu1", style = pct, cols=c(9,11,13,15), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)
addStyle(wb_overall, "cc_snu1", style = pct2, cols=c(7), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_overall, "cc_psnu",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_psnu", cc_psnu, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "cc_psnu", style = pct, cols=c(10,12,14,16), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)
addStyle(wb_overall, "cc_psnu", style = pct2, cols=c(8), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "cc_site",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_site", cc_site, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "cc_site", style = pct, cols=c(13,15,17,19), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)
addStyle(wb_overall, "cc_site", style = pct2, cols=c(11), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)

##############################################
##############################################
##############################################
## TX CASCADE - FEMALE #######################
##############################################
##############################################
##############################################

df_fem <- df_q %>%
  filter(sex == "Female")

#############################################
# CREATE SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df_fem %>%
  filter(indicator == "TX_CURR", disag == "Age/Sex/HIVStatus", target > 0) %>%
  select(site_id, target) %>%
  group_by(site_id) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df_fem %>%
  filter(indicator == "TX_CURR", disag == "Age/Sex/HIVStatus", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = starts_with("Q")) 

#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df_fem %>%
  filter(indicator == "HTS_TST",
         disag == "Modality/Age/Sex/Result",
         partner == "CHASS" | partner == "FGH" | partner == "ICAP" | partner == "CCS" | partner == "ARIEL" | partner == "EGPAF"| partner == "ECHO",
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

hts_tst_pos <- df_fem %>%
  filter(indicator == "HTS_TST_POS",
         disag == "Modality/Age/Sex/Result",
         partner == "CHASS" | partner == "FGH" | partner == "ICAP" | partner == "CCS" | partner == "ARIEL" | partner == "EGPAF"| partner == "ECHO",
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

tx_new <- df_fem %>%
  filter(indicator == "TX_NEW", disag == "Age/Sex/HIVStatus", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = starts_with("Q"))

#---------------------------------------------
# TX_PVLS DENOMINATOR REFERENCE OBJECT

tx_pvls_d <- df_fem %>%
  filter(indicator == "TX_PVLS", disag == "Age/Sex/Indication/HIVStatus", numeratordenom == "D", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = starts_with("Q"))

#---------------------------------------------
# TX_PVLS NUMERATOR REFERENCE OBJECT

tx_pvls_n <- df_fem %>%
  filter(indicator == "TX_PVLS", disag == "Age/Sex/Indication/HIVStatus", numeratordenom == "N", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = starts_with("Q"))

#---------------------------------------------
# JOIN REFERENCE OBJECTS

cc_site <- tx_curr %>%
  left_join(tx_curr_target, by = "site_id") %>%
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
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(site_id, agency, partner, snu1, psnu, site, ajuda, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls) %>%
  arrange(agency, partner, snu1, psnu, site)


#---------------------------------------------
##CLEAN-UP R ENVIRONMENT BY REMOVING UNNEEDED OBJECTS

rm(tx_curr_target, tx_curr, tx_new, tx_pvls_d, tx_pvls_n, hts_tst, hts_tst_pos)

#---------------------------------------------
## CREATE AJUDA PHASE CASCADE

cc_ajuda <- cc_site %>%
  group_by(ajuda) %>%
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
  select(ajuda, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


##############################################
# CREATE ABOVE-SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df_fem %>%
  filter(indicator == "TX_CURR", disag == "Age/Sex/HIVStatus", target > 0) %>%
  select(partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, target) %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# HTS REFERENCE OBJECT

hts_tst <- df_fem %>%
  filter(indicator == "HTS_TST", disag == "Modality/Age/Sex/Result", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst = "Q")

cc_psnu <- tx_curr_target %>%
  left_join(hts_tst, by = "psnu") %>%
  replace_na(list(hts_tst = 0))

#---------------------------------------------
# HTS_POS REFERENCE OBJECT

hts_tst_pos <- df_fem %>%
  filter(indicator == "HTS_TST_POS", disag == "Modality/Age/Sex/Result", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst_pos = "Q")

cc_psnu <- cc_psnu %>%
  left_join(hts_tst_pos, by = "psnu") %>%
  replace_na(list(hts_tst_pos = 0))

#---------------------------------------------
# TX_NEW REFERENCE OBJECT

tx_new <- df_fem %>%
  filter(indicator == "TX_NEW", disag == "Age/Sex/HIVStatus", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_new, by = "psnu") %>%
  replace_na(list(tx_new = 0))

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df_fem %>%
  filter(indicator == "TX_CURR", disag == "Age/Sex/HIVStatus", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_curr, by = "psnu") %>%
  replace_na(list(tx_curr = 0))

#---------------------------------------------
# TX_PLVS_DEN REFERENCE OBJECT

tx_pvls_d <- df_fem %>%
  filter(indicator == "TX_PVLS", disag == "Age/Sex/Indication/HIVStatus", numeratordenom == "D", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_pvls_d, by = "psnu") %>%
  replace_na(list(tx_pvls_d = 0))

#---------------------------------------------
# TX_PLVS_NUM REFERENCE OBJECT

tx_pvls_n <- df_fem %>%
  filter(indicator == "TX_PVLS", disag == "Age/Sex/Indication/HIVStatus", numeratordenom == "N", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = "Q") 


##############################################
# CREATE ABOVE-SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
## CREATE DISTRICT CLINICAL CASCADE OBJECT

cc_psnu <- cc_psnu %>%
  left_join(tx_pvls_n, by = "psnu") %>%
  replace_na(list(tx_pvls_n = 0)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_linkage = tx_new / hts_tst_pos,
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(agency, partner, snu1, psnu, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

#---------------------------------------------
## CREATE PROVINCIAL CLINICAL CASCADE OBJECT

cc_snu1 <- cc_psnu %>%
  group_by(agency, partner, snu1) %>%
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
  select(agency, partner, snu1, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

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
## CREATE OU CLINICAL CASCADE OBJECT

cc_ou <- cc_psnu %>%
  mutate(agency = "Moz") %>%
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

##############################################
# CREATE OBJECT FOR OVERALL RESULTS
##############################################

cc_ou_female <- cc_ou %>%
  select(-agency) %>%
  mutate(resulttype = "Female") %>%
  select(resulttype, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

##############################################
# GENERATE EXCEL OUTPUTS
##############################################

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_female <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")
pct2 <- createStyle(numFmt = "0.0%", textDecoration = "italic")

addWorksheet(wb_female, "cc_ou",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_ou", cc_ou, tableStyle = "TableStyleLight2")
addStyle(wb_female, "cc_ou", style = pct, cols=c(7,9,11,13), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)
addStyle(wb_female, "cc_ou", style = pct2, cols=c(5), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)

addWorksheet(wb_female, "cc_ajuda",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_ajuda", cc_ajuda, tableStyle = "TableStyleLight2")
addStyle(wb_female, "cc_ajuda", style = pct, cols=c(7,9,11,13), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)
addStyle(wb_female, "cc_ajuda", style = pct2, cols=c(5), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_female, "cc_agency",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_agency", cc_agency, tableStyle = "TableStyleLight2")
addStyle(wb_female, "cc_agency", style = pct, cols=c(7,9,11,13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)
addStyle(wb_female, "cc_agency", style = pct2, cols=c(5), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)

addWorksheet(wb_female, "cc_partner",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_partner", cc_partner, tableStyle = "TableStyleLight2")
addStyle(wb_female, "cc_partner", style = pct, cols=c(8,10,12,14), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)
addStyle(wb_female, "cc_partner", style = pct2, cols=c(6), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)

addWorksheet(wb_female, "cc_snu1",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_snu1", cc_snu1, tableStyle = "TableStyleLight2")
addStyle(wb_female, "cc_snu1", style = pct, cols=c(9,11,13,15), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)
addStyle(wb_female, "cc_snu1", style = pct2, cols=c(7), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)

addWorksheet(wb_female, "cc_psnu",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_psnu", cc_psnu, tableStyle = "TableStyleLight2")
addStyle(wb_female, "cc_psnu", style = pct, cols=c(10,12,14,16), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)
addStyle(wb_female, "cc_psnu", style = pct2, cols=c(8), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)


addWorksheet(wb_female, "cc_site",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_site", cc_site, tableStyle = "TableStyleLight2")
addStyle(wb_female, "cc_site", style = pct, cols=c(13,15,17,19), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)
addStyle(wb_female, "cc_site", style = pct2, cols=c(11), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)

##############################################
##############################################
##############################################
## TX CASCADE - MALE #########################
##############################################
##############################################
##############################################

df_mal <- df_q %>%
  filter(sex == "Male")

#############################################
# CREATE SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df_mal %>%
  filter(indicator == "TX_CURR", disag == "Age/Sex/HIVStatus", target > 0) %>%
  select(site_id, target) %>%
  group_by(site_id) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df_mal %>%
  filter(indicator == "TX_CURR", disag == "Age/Sex/HIVStatus", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = starts_with("Q")) 


# TX_CURR REFERENCE OBJECT

tx_curr <- df_mal %>%
  filter(indicator == "TX_CURR", disag == "Age/Sex/HIVStatus", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = starts_with("Q")) 


#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df_mal %>%
  filter(indicator == "HTS_TST",
         disag == "Modality/Age/Sex/Result",
         partner == "CHASS" | partner == "FGH" | partner == "ICAP" | partner == "CCS" | partner == "ARIEL" | partner == "EGPAF"| partner == "ECHO",
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

hts_tst_pos <- df_mal %>%
  filter(indicator == "HTS_TST_POS",
         disag == "Modality/Age/Sex/Result",
         partner == "CHASS" | partner == "FGH" | partner == "ICAP" | partner == "CCS" | partner == "ARIEL" | partner == "EGPAF"| partner == "ECHO",
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

tx_new <- df_mal %>%
  filter(indicator == "TX_NEW", disag == "Age/Sex/HIVStatus", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = starts_with("Q"))

#---------------------------------------------
# TX_PVLS DENOMINATOR REFERENCE OBJECT

tx_pvls_d <- df_mal %>%
  filter(indicator == "TX_PVLS", disag == "Age/Sex/Indication/HIVStatus", numeratordenom == "D", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = starts_with("Q"))

#---------------------------------------------
# TX_PVLS NUMERATOR REFERENCE OBJECT

tx_pvls_n <- df_mal %>%
  filter(indicator == "TX_PVLS", disag == "Age/Sex/Indication/HIVStatus", numeratordenom == "N", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = starts_with("Q"))

#---------------------------------------------
# JOIN REFERENCE OBJECTS

cc_site <- tx_curr %>%
  left_join(tx_curr_target, by = "site_id") %>%
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
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(site_id, agency, partner, snu1, psnu, site, ajuda, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls) %>%
  arrange(agency, partner, snu1, psnu, site)


#---------------------------------------------
##CLEAN-UP R ENVIRONMENT BY REMOVING UNNEEDED OBJECTS

rm(tx_curr_target, tx_curr, tx_new, tx_pvls_d, tx_pvls_n, hts_tst, hts_tst_pos)

#---------------------------------------------
## CREATE AJUDA PHASE CASCADE

cc_ajuda <- cc_site %>%
  group_by(ajuda) %>%
  summarize(tx_curr_target = sum(tx_curr_target, na. = TRUE),
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
  select(ajuda, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


##############################################
# CREATE ABOVE-SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df_mal %>%
  filter(indicator == "TX_CURR", disag == "Age/Sex/HIVStatus", target > 0) %>%
  select(partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, target) %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# HTS REFERENCE OBJECT

hts_tst <- df_mal %>%
  filter(indicator == "HTS_TST", disag == "Modality/Age/Sex/Result", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst = "Q")

cc_psnu <- tx_curr_target %>%
  left_join(hts_tst, by = "psnu") %>%
  replace_na(list(hts_tst = 0))

#---------------------------------------------
# HTS_POS REFERENCE OBJECT

hts_tst_pos <- df_mal %>%
  filter(indicator == "HTS_TST_POS", disag == "Modality/Age/Sex/Result", Q > 0) %>%
  select(indicator, disag, snu1,psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst_pos = "Q")

cc_psnu <- cc_psnu %>%
  left_join(hts_tst_pos, by = "psnu") %>%
  replace_na(list(hts_tst_pos = 0))

#---------------------------------------------
# TX_NEW REFERENCE OBJECT

tx_new <- df_mal %>%
  filter(indicator == "TX_NEW", disag == "Age/Sex/HIVStatus", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_new, by = "psnu") %>%
  replace_na(list(tx_new = 0))

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df_mal %>%
  filter(indicator == "TX_CURR", disag == "Age/Sex/HIVStatus", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_curr, by = "psnu") %>%
  replace_na(list(tx_curr = 0))

#---------------------------------------------
# TX_PLVS_DEN REFERENCE OBJECT

tx_pvls_d <- df_mal %>%
  filter(indicator == "TX_PVLS", disag == "Age/Sex/Indication/HIVStatus", numeratordenom == "D", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = "Q") 

cc_psnu <- cc_psnu %>%
  left_join(tx_pvls_d, by = "psnu") %>%
  replace_na(list(tx_pvls_d = 0))

#---------------------------------------------
# TX_PLVS_NUM REFERENCE OBJECT

tx_pvls_n <- df_mal %>%
  filter(indicator == "TX_PVLS", disag == "Age/Sex/Indication/HIVStatus", numeratordenom == "N", Q > 0) %>%
  select(indicator, disag, snu1, psnu, fiscal_year, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = "Q") 


##############################################
# CREATE ABOVE-SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
## CREATE DISTRICT CLINICAL CASCADE OBJECT

cc_psnu <- cc_psnu %>%
  left_join(tx_pvls_n, by = "psnu") %>%
  replace_na(list(tx_pvls_n = 0)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_linkage = tx_new / hts_tst_pos,
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(agency, partner, snu1, psnu, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

#---------------------------------------------
## CREATE PROVINCIAL CLINICAL CASCADE OBJECT

cc_snu1 <- cc_psnu %>%
  group_by(agency, partner, snu1) %>%
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
  select(agency, partner, snu1, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

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
## CREATE OU CLINICAL CASCADE OBJECT

cc_ou <- cc_psnu %>%
  mutate(agency = "Moz") %>%
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

##############################################
# CREATE OBJECT FOR OVERALL RESULTS
##############################################

cc_ou_male <- cc_ou %>%
  select(-agency) %>%
  mutate(resulttype = "Male") %>%
  select(resulttype, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_male <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")
pct2 <- createStyle(numFmt = "0.0%", textDecoration = "italic")

addWorksheet(wb_male, "cc_ou",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_ou", cc_ou, tableStyle = "TableStyleLight2")
addStyle(wb_male, "cc_ou", style = pct, cols=c(7,9,11,13), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)
addStyle(wb_male, "cc_ou", style = pct2, cols=c(5), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)

addWorksheet(wb_male, "cc_ajuda",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_ajuda", cc_ajuda, tableStyle = "TableStyleLight2")
addStyle(wb_male, "cc_ajuda", style = pct, cols=c(7,9,11,13), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)
addStyle(wb_male, "cc_ajuda", style = pct2, cols=c(5), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_male, "cc_agency",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_agency", cc_agency, tableStyle = "TableStyleLight2")
addStyle(wb_male, "cc_agency", style = pct, cols=c(7,9,11,13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)
addStyle(wb_male, "cc_agency", style = pct2, cols=c(5), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)

addWorksheet(wb_male, "cc_partner",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_partner", cc_partner, tableStyle = "TableStyleLight2")
addStyle(wb_male, "cc_partner", style = pct, cols=c(8,10,12,14), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)
addStyle(wb_male, "cc_partner", style = pct2, cols=c(6), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)

addWorksheet(wb_male, "cc_snu1",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_snu1", cc_snu1, tableStyle = "TableStyleLight2")
addStyle(wb_male, "cc_snu1", style = pct, cols=c(9,11,13,15), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)
addStyle(wb_male, "cc_snu1", style = pct2, cols=c(7), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)

addWorksheet(wb_male, "cc_psnu",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_psnu", cc_psnu, tableStyle = "TableStyleLight2")
addStyle(wb_male, "cc_psnu", style = pct, cols=c(10,12,14,16), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)
addStyle(wb_male, "cc_psnu", style = pct2, cols=c(8), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)

addWorksheet(wb_male, "cc_site",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_site", cc_site, tableStyle = "TableStyleLight2")
addStyle(wb_male, "cc_site", style = pct, cols=c(13,15,17,19), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)
addStyle(wb_male, "cc_site", style = pct2, cols=c(11), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)

##############################################
##############################################
##############################################
## TX CASCADE - PEDS #########################
##############################################
##############################################
##############################################

df_ped <- df_q %>%
  filter(trendscoarse == "<15")

##############################################
# CREATE SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df_ped %>%
  filter(indicator == "TX_CURR", target > 0) %>%
  select(site_id, target) %>%
  group_by(site_id) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df_ped %>%
  filter(indicator == "TX_CURR", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = starts_with("Q")) 

#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df_ped %>%
  filter(indicator == "HTS_TST",
         partner == "CHASS" | partner == "FGH" | partner == "ICAP" | partner == "CCS" | partner == "ARIEL" | partner == "EGPAF"| partner == "ECHO",
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

hts_tst_pos <- df_ped %>%
  filter(indicator == "HTS_TST_POS",
         partner == "CHASS" | partner == "FGH" | partner == "ICAP" | partner == "CCS" | partner == "ARIEL" | partner == "EGPAF"| partner == "ECHO",
         Q > 0
  ) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst_pos = starts_with("Q")) %>%
  replace_na(list(hts_tst_pos = 0))

#---------------------------------------------
# PMTCT_EID_NUM REFERENCE OBJECT

p_eid_n <- df_ped %>%
  filter(indicator == "PMTCT_EID", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_n = starts_with("Q"))

#---------------------------------------------
# PMTCT_HEI_POS REFERENCE OBJECT

p_eid_hei_pos <- df_ped %>%
  filter(indicator == "PMTCT_HEI_POS", disag == "Age/HIVStatus", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_hei_pos = starts_with("Q"))

#---------------------------------------------
# TX_NEW REFERENCE OBJECT

tx_new <- df_ped %>%
  filter(indicator == "TX_NEW", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = starts_with("Q"))

#---------------------------------------------
# TX_PVLS DENOMINATOR REFERENCE OBJECT

tx_pvls_d <- df_ped %>%
  filter(indicator == "TX_PVLS", numeratordenom == "D", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = starts_with("Q"))

#---------------------------------------------
# TX_PVLS NUMERATOR REFERENCE OBJECT

tx_pvls_n <- df_ped %>%
  filter(indicator == "TX_PVLS", numeratordenom == "N", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = starts_with("Q"))

#---------------------------------------------
# JOIN REFERENCE OBJECTS

cc_ped_site <- tx_curr %>%
  left_join(tx_curr_target, by = "site_id") %>%
  left_join(hts_tst, by = "site_id") %>%
  left_join(p_eid_n, by = "site_id") %>%
  left_join(hts_tst_pos, by = "site_id") %>%
  left_join(p_eid_hei_pos, by = "site_id") %>%
  left_join(tx_new, by = "site_id") %>%
  left_join(tx_pvls_d, by = "site_id") %>%
  left_join(tx_pvls_n, by  = "site_id") %>%
  left_join(ajuda, by = "site_id") %>%
  replace_na(list(hts_tst = 0,
                  p_eid_n = 0,
                  hts_tst_pos = 0,
                  p_eid_hei_pos = 0,
                  tx_new = 0,
                  tx_curr = 0,
                  tx_pvls_d = 0,
                  tx_pvls_n = 0,
                  ajuda = 0)
  ) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_eid_pos = p_eid_hei_pos / p_eid_n,
         per_linkage = tx_new / (hts_tst_pos + p_eid_hei_pos),
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(site_id, agency, partner, snu1, psnu, site, ajuda, tx_curr_target, p_eid_n, p_eid_hei_pos, per_eid_pos, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls) %>%
  arrange(agency, partner, snu1, psnu, site)


#---------------------------------------------
# CLEAN-UP R ENVIRONMENT BY REMOVING UNNEEDED OBJECTS

rm(tx_curr_target, tx_curr, tx_new, tx_pvls_d, tx_pvls_n, hts_tst, hts_tst_pos, p_eid_n, p_eid_hei_pos)

#---------------------------------------------
# CREATE PROVINCE CASCADE

cc_ped_ajuda <- cc_ped_site %>%
  group_by(ajuda) %>%
  summarize(tx_curr_target = sum(tx_curr_target, na.rm = TRUE),
            hts_tst = sum(hts_tst, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),            
            hts_tst_pos = sum(hts_tst_pos, na.rm = TRUE),
            p_eid_hei_pos = sum(p_eid_hei_pos, na.rm = TRUE),   
            tx_new = sum(tx_new, na.rm = TRUE),
            tx_curr = sum(tx_curr, na.rm = TRUE),
            tx_pvls_d = sum(tx_pvls_d, na.rm = TRUE),
            tx_pvls_n = sum(tx_pvls_n, na.rm = TRUE)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_eid_pos = p_eid_hei_pos / p_eid_n,
         per_linkage = tx_new / (hts_tst_pos + p_eid_hei_pos),
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  ungroup() %>%
  select(ajuda, tx_curr_target, p_eid_n, p_eid_hei_pos, per_eid_pos, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


##############################################
# CREATE ABOVE-SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df_ped %>%
  filter(indicator == "TX_CURR", target > 0) %>%
  select(psnu, target) %>%
  group_by(psnu) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df_ped %>%
  filter(indicator == "TX_CURR", Q > 0) %>%
  select(agency, partner, snu1, psnu, Q) %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = starts_with("Q")) 

#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df_ped %>%
  filter(indicator == "HTS_TST", Q > 0) %>%
  select(psnu, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst = starts_with("Q")) %>%
  replace_na(list(hts_tst = 0))

#---------------------------------------------
# HTS_TST_POS REFERENCE OBJECT

hts_tst_pos <- df_ped %>%
  filter(indicator == "HTS_TST_POS", Q > 0) %>%
  select(psnu, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst_pos = starts_with("Q")) %>%
  replace_na(list(hts_tst = 0))

#---------------------------------------------
# PMTCT_EID_NUM REFERENCE OBJECT

p_eid_n <- df_ped %>%
  filter(indicator == "PMTCT_EID", Q > 0) %>%
  select(psnu, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_n = starts_with("Q"))

#---------------------------------------------
# PMTCT_HEI_POS REFERENCE OBJECT

p_eid_hei_pos <- df_ped %>%
  filter(indicator == "PMTCT_HEI_POS", disag == "Age/HIVStatus", Q > 0) %>%
  select(psnu, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_hei_pos = starts_with("Q"))

#---------------------------------------------
# TX_NEW REFERENCE OBJECT

tx_new <- df_ped %>%
  filter(indicator == "TX_NEW", Q > 0) %>%
  select(psnu, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = starts_with("Q"))

#---------------------------------------------
# TX_PVLS DENOMINATOR REFERENCE OBJECT

tx_pvls_d <- df_ped %>%
  filter(indicator == "TX_PVLS", numeratordenom == "D", Q > 0) %>%
  select(psnu, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = starts_with("Q"))

#---------------------------------------------
# TX_PVLS NUMERATOR REFERENCE OBJECT

tx_pvls_n <- df_ped %>%
  filter(indicator == "TX_PVLS", numeratordenom == "N", Q > 0) %>%
  select(psnu, Q) %>%
  group_by(psnu) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = starts_with("Q"))

#---------------------------------------------
# JOIN REFERENCE OBJECTS & CREATE DISTRICT CASCADE

cc_ped_psnu <- tx_curr %>%
  left_join(tx_curr_target, by = "psnu") %>%
  left_join(hts_tst, by = "psnu") %>%
  left_join(p_eid_n, by = "psnu") %>%
  left_join(hts_tst_pos, by = "psnu") %>%
  left_join(p_eid_hei_pos, by = "psnu") %>%
  left_join(tx_new, by = "psnu") %>%
  left_join(tx_pvls_d, by = "psnu") %>%
  left_join(tx_pvls_n, by  = "psnu") %>%
  replace_na(list(hts_tst = 0,
                  p_eid_n = 0,
                  hts_tst_pos = 0,
                  p_eid_hei_pos = 0,
                  tx_new = 0,
                  tx_curr = 0,
                  tx_pvls_d = 0,
                  tx_pvls_n = 0)
  ) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_eid_pos = p_eid_hei_pos / p_eid_n,
         per_linkage = tx_new / (hts_tst_pos + p_eid_hei_pos),
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(agency, partner, snu1, psnu, tx_curr_target, p_eid_n, p_eid_hei_pos, per_eid_pos, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

#---------------------------------------------
# CREATE PROVINCE CASCADE

cc_ped_snu1 <- cc_ped_psnu %>%
  group_by(agency, partner, snu1) %>%
  summarize(tx_curr_target = sum(tx_curr_target, na.rm = TRUE),
            hts_tst = sum(hts_tst, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),            
            hts_tst_pos = sum(hts_tst_pos, na.rm = TRUE),
            p_eid_hei_pos = sum(p_eid_hei_pos, na.rm = TRUE),   
            tx_new = sum(tx_new, na.rm = TRUE),
            tx_curr = sum(tx_curr, na.rm = TRUE),
            tx_pvls_d = sum(tx_pvls_d, na.rm = TRUE),
            tx_pvls_n = sum(tx_pvls_n, na.rm = TRUE)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_eid_pos = p_eid_hei_pos / p_eid_n,
         per_linkage = tx_new / (hts_tst_pos + p_eid_hei_pos),
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  ungroup() %>%
  select(agency, partner, snu1, tx_curr_target, p_eid_n, p_eid_hei_pos, per_eid_pos, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


#---------------------------------------------
# CREATE PARTNER CASCADE

cc_ped_partner <- cc_ped_psnu %>%
  group_by(agency, partner) %>%
  summarize(tx_curr_target = sum(tx_curr_target, na.rm = TRUE),
            hts_tst = sum(hts_tst, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),            
            hts_tst_pos = sum(hts_tst_pos, na.rm = TRUE),
            p_eid_hei_pos = sum(p_eid_hei_pos, na.rm = TRUE),   
            tx_new = sum(tx_new, na.rm = TRUE),
            tx_curr = sum(tx_curr, na.rm = TRUE),
            tx_pvls_d = sum(tx_pvls_d, na.rm = TRUE),
            tx_pvls_n = sum(tx_pvls_n, na.rm = TRUE)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_eid_pos = p_eid_hei_pos / p_eid_n,
         per_linkage = tx_new / (hts_tst_pos + p_eid_hei_pos),
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  ungroup() %>%
  select(agency, partner, tx_curr_target, p_eid_n, p_eid_hei_pos, per_eid_pos, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

#---------------------------------------------
# CREATE AGENCY CASCADE

cc_ped_agency <- cc_ped_psnu %>%
  group_by(agency) %>%
  summarize(tx_curr_target = sum(tx_curr_target, na.rm = TRUE),
            hts_tst = sum(hts_tst, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),            
            hts_tst_pos = sum(hts_tst_pos, na.rm = TRUE),
            p_eid_hei_pos = sum(p_eid_hei_pos, na.rm = TRUE),   
            tx_new = sum(tx_new, na.rm = TRUE),
            tx_curr = sum(tx_curr, na.rm = TRUE),
            tx_pvls_d = sum(tx_pvls_d, na.rm = TRUE),
            tx_pvls_n = sum(tx_pvls_n, na.rm = TRUE)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_eid_pos = p_eid_hei_pos / p_eid_n,
         per_linkage = tx_new / (hts_tst_pos + p_eid_hei_pos),
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  ungroup() %>%
  select(agency, tx_curr_target, p_eid_n, p_eid_hei_pos, per_eid_pos, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

#---------------------------------------------
# CREATE AJUDA CASCADE

cc_ped_ou <- cc_ped_psnu %>%
  mutate(agency = "Moz") %>%
  group_by(agency) %>%
  summarize(tx_curr_target = sum(tx_curr_target, na.rm = TRUE),
            hts_tst = sum(hts_tst, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),            
            hts_tst_pos = sum(hts_tst_pos, na.rm = TRUE),
            p_eid_hei_pos = sum(p_eid_hei_pos, na.rm = TRUE),   
            tx_new = sum(tx_new, na.rm = TRUE),
            tx_curr = sum(tx_curr, na.rm = TRUE),
            tx_pvls_d = sum(tx_pvls_d, na.rm = TRUE),
            tx_pvls_n = sum(tx_pvls_n, na.rm = TRUE)) %>%
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_eid_pos = p_eid_hei_pos / p_eid_n,
         per_linkage = tx_new / (hts_tst_pos + p_eid_hei_pos),
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  ungroup() %>%
  select(agency, tx_curr_target, p_eid_n, p_eid_hei_pos, per_eid_pos, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


#---------------------------------------------
##CLEAN-UP R ENVIRONMENT BY REMOVING UNNEEDED OBJECTS

rm(tx_curr, tx_curr_target, tx_new, tx_pvls_d, tx_pvls_n, hts_tst, hts_tst_pos, p_eid_hei_pos, p_eid_n)

##############################################
# CREATE OBJECT FOR OVERALL RESULTS
##############################################

cc_ou_ped <- cc_ou %>%
  select(-agency) %>%
  mutate(resulttype = "Pediatric") %>%
  select(resulttype, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

##############################################
# GENERATE EXCEL OUTPUTS
##############################################

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH CLINICAL CASCADE (PEDS)

wb_pediatrico <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")
pct2 <- createStyle(numFmt = "0.0%", textDecoration = "italic")

addWorksheet(wb_pediatrico, "cc_ped_ou",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_ou", cc_ped_ou, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "cc_ped_ou", style = pct, cols=c(10,12,14,16), rows = 2:(nrow(cc_ped_ou)+1), gridExpand=TRUE)
addStyle(wb_pediatrico, "cc_ped_ou", style = pct2, cols=c(5,8), rows = 2:(nrow(cc_ped_ou)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "cc_ped_ajuda",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_ajuda", cc_ped_ajuda, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "cc_ped_ajuda", style = pct, cols=c(10,12,14,16), rows = 2:(nrow(cc_ped_ajuda)+1), gridExpand=TRUE)
addStyle(wb_pediatrico, "cc_ped_ajuda", style = pct2, cols=c(5,8), rows = 2:(nrow(cc_ped_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "cc_ped_agency",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_agency", cc_ped_agency, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "cc_ped_agency", style = pct, cols=c(10,12,14,16), rows = 2:(nrow(cc_ped_agency)+1), gridExpand=TRUE)
addStyle(wb_pediatrico, "cc_ped_agency", style = pct2, cols=c(5,8), rows = 2:(nrow(cc_ped_agency)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "cc_ped_partner",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_partner", cc_ped_partner, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "cc_ped_partner", style = pct, cols=c(11,13,15,17), rows = 2:(nrow(cc_ped_partner)+1), gridExpand=TRUE)
addStyle(wb_pediatrico, "cc_ped_partner", style = pct2, cols=c(6,9), rows = 2:(nrow(cc_ped_partner)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "cc_ped_snu1",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_snu1", cc_ped_snu1, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "cc_ped_snu1", style = pct, cols=c(12,14,16,18), rows = 2:(nrow(cc_ped_snu1)+1), gridExpand=TRUE)
addStyle(wb_pediatrico, "cc_ped_snu1", style = pct2, cols=c(7,10), rows = 2:(nrow(cc_ped_snu1)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "cc_ped_psnu",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_psnu", cc_ped_psnu, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "cc_ped_psnu", style = pct, cols=c(13,15,17,19), rows = 2:(nrow(cc_ped_psnu)+1), gridExpand=TRUE)
addStyle(wb_pediatrico, "cc_ped_psnu", style = pct2, cols=c(8,11), rows = 2:(nrow(cc_ped_psnu)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "cc_ped_site",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_site", cc_ped_site, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "cc_ped_site", style = pct, cols=c(16,18,20,22), rows = 2:(nrow(cc_ped_site)+1), gridExpand=TRUE)
addStyle(wb_pediatrico, "cc_ped_site", style = pct2, cols=c(11,14), rows = 2:(nrow(cc_ped_site)+1), gridExpand=TRUE)


saveWorkbook(wb_overall, file = "C:/Users/cnhantumbo/Documents/R codes/cc_cascade_overall.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
saveWorkbook(wb_female, file = "C:/Users/cnhantumbo/Documents/R codes/cc_cascade_female.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
saveWorkbook(wb_male, file = "C:/Users/cnhantumbo/Documents/R codes/cc_cascade_male.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
saveWorkbook(wb_pediatrico, file = "C:/Users/cnhantumbo/Documents/R codes/cc_cascade_ped.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!

