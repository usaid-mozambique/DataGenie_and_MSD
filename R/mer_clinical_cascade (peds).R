rm(list = ls())

library(tidyverse)
library(janitor)
library(readxl)
library(openxlsx)


##############################################
# IMPORT SOURCE FILES
##############################################

#---------------------------------------------
# GENIE IMPORT AND SET UP OF BASE DATAFRAME. NOTE THAT "FILTER" AND "SELECT" CODE LINES ARE ANALYSIS SPECIFIC 

df <- read.delim("~/R/datasets/Genie_Daily_80bc795352084a32adb1299ebaa7a132.txt") %>%
  filter(Fiscal_Year == "2019", TrendsCoarse == "<15") %>%
  select(-Qtr1, -Qtr2, -Qtr4) %>%
  rename(target = TARGETS,
         site_id = FacilityUID, 
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

ajuda <- read_excel("~/R/datasets/ajuda_ia.xlsx") %>%
  replace_na(list(ajuda = 0))


##############################################
# CREATE SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df %>%
  filter(indicator == "TX_CURR", target > 0) %>%
  select(site_id, target) %>%
  group_by(site_id) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df %>%
  filter(indicator == "TX_CURR", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = starts_with("Q")) 

#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df %>%
  filter(indicator == "HTS_TST",
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
# PMTCT_EID_NUM REFERENCE OBJECT

p_eid_n <- df %>%
  filter(indicator == "PMTCT_EID", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_n = starts_with("Q"))

#---------------------------------------------
# PMTCT_HEI_POS REFERENCE OBJECT

p_eid_hei_pos <- df %>%
  filter(indicator == "PMTCT_HEI_POS", disag == "Age/HIVStatus", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_hei_pos = starts_with("Q"))

#---------------------------------------------
# TX_NEW REFERENCE OBJECT

tx_new <- df %>%
  filter(indicator == "TX_NEW", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = starts_with("Q"))

#---------------------------------------------
# TX_PVLS DENOMINATOR REFERENCE OBJECT

tx_pvls_d <- df %>%
  filter(indicator == "TX_PVLS", numeratorDenom == "D", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = starts_with("Q"))

#---------------------------------------------
# TX_PVLS NUMERATOR REFERENCE OBJECT

tx_pvls_n <- df %>%
  filter(indicator == "TX_PVLS", numeratorDenom == "N", Q > 0) %>%
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
  select(site_id, agency, partner, SNU1, PSNU, site, ajuda, tx_curr_target, p_eid_n, p_eid_hei_pos, per_eid_pos, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

#---------------------------------------------
# CLEAN-UP R ENVIRONMENT BY REMOVING UNNEEDED OBJECTS

rm(ajuda, tx_curr_target, tx_curr, tx_new, tx_pvls_d, tx_pvls_n, hts_tst, hts_tst_pos, p_eid_n, p_eid_hei_pos)


##############################################
# CREATE ABOVE-SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df %>%
  filter(indicator == "TX_CURR", target > 0) %>%
  select(PSNU, target) %>%
  group_by(PSNU) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df %>%
  filter(indicator == "TX_CURR", Q > 0) %>%
  select(agency, partner, SNU1, PSNU, Q) %>%
  group_by(agency, partner, SNU1, PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = starts_with("Q")) 

#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df %>%
  filter(indicator == "HTS_TST", Q > 0) %>%
  select(PSNU, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst = starts_with("Q")) %>%
  replace_na(list(hts_tst = 0))

#---------------------------------------------
# HTS_TST_POS REFERENCE OBJECT

hts_tst_pos <- df %>%
  filter(indicator == "HTS_TST_POS", Q > 0) %>%
  select(PSNU, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(hts_tst_pos = starts_with("Q")) %>%
  replace_na(list(hts_tst = 0))

#---------------------------------------------
# PMTCT_EID_NUM REFERENCE OBJECT

p_eid_n <- df %>%
  filter(indicator == "PMTCT_EID", Q > 0) %>%
  select(PSNU, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_n = starts_with("Q"))

#---------------------------------------------
# PMTCT_HEI_POS REFERENCE OBJECT

p_eid_hei_pos <- df %>%
  filter(indicator == "PMTCT_HEI_POS", disag == "Age/HIVStatus", Q > 0) %>%
  select(PSNU, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_hei_pos = starts_with("Q"))

#---------------------------------------------
# TX_NEW REFERENCE OBJECT

tx_new <- df %>%
  filter(indicator == "TX_NEW", Q > 0) %>%
  select(PSNU, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = starts_with("Q"))

#---------------------------------------------
# TX_PVLS DENOMINATOR REFERENCE OBJECT

tx_pvls_d <- df %>%
  filter(indicator == "TX_PVLS", numeratorDenom == "D", Q > 0) %>%
  select(PSNU, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = starts_with("Q"))

#---------------------------------------------
# TX_PVLS NUMERATOR REFERENCE OBJECT

tx_pvls_n <- df %>%
  filter(indicator == "TX_PVLS", numeratorDenom == "N", Q > 0) %>%
  select(PSNU, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_n = starts_with("Q"))

#---------------------------------------------
# JOIN REFERENCE OBJECTS & CREATE DISTRICT CASCADE

cc_ped_psnu <- tx_curr %>%
  left_join(tx_curr_target, by = "PSNU") %>%
  left_join(hts_tst, by = "PSNU") %>%
  left_join(p_eid_n, by = "PSNU") %>%
  left_join(hts_tst_pos, by = "PSNU") %>%
  left_join(p_eid_hei_pos, by = "PSNU") %>%
  left_join(tx_new, by = "PSNU") %>%
  left_join(tx_pvls_d, by = "PSNU") %>%
  left_join(tx_pvls_n, by  = "PSNU") %>%
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
  select(agency, partner, SNU1, PSNU, tx_curr_target, p_eid_n, p_eid_hei_pos, per_eid_pos, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)

#---------------------------------------------
# CREATE PROVINCE CASCADE

cc_ped_snu1 <- cc_ped_psnu %>%
  group_by(agency, partner, SNU1) %>%
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
  select(agency, partner, SNU1, tx_curr_target, p_eid_n, p_eid_hei_pos, per_eid_pos, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls)


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
##CLEAN-UP R ENVIRONMENT BY REMOVING UNNEEDED OBJECTS

rm(tx_curr, tx_curr_target, tx_new, tx_pvls_d, tx_pvls_n, hts_tst, hts_tst_pos, p_eid_hei_pos, p_eid_n)


##############################################
# GENERATE EXCEL OUTPUTS
##############################################

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH CLINICAL CASCADE (PEDS)

wb <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")

addWorksheet(wb, "cc_ped_agency",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb, "cc_ped_agency", cc_ped_agency, tableStyle = "TableStyleLight4")
addStyle(wb, "cc_ped_agency", style = pct, cols=c(5,8,10,12,14,16), rows = 2:(nrow(cc_ped_agency)+1), gridExpand=TRUE)

addWorksheet(wb, "cc_ped_partner",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb, "cc_ped_partner", cc_ped_partner, tableStyle = "TableStyleLight4")
addStyle(wb, "cc_ped_partner", style = pct, cols=c(6,9,11,13,15,17), rows = 2:(nrow(cc_ped_partner)+1), gridExpand=TRUE)

addWorksheet(wb, "cc_ped_snu1",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb, "cc_ped_snu1", cc_ped_snu1, tableStyle = "TableStyleLight4")
addStyle(wb, "cc_ped_snu1", style = pct, cols=c(7,10,12,14,16,18), rows = 2:(nrow(cc_ped_snu1)+1), gridExpand=TRUE)


addWorksheet(wb, "cc_ped_psnu",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb, "cc_ped_psnu", cc_ped_psnu, tableStyle = "TableStyleLight4")
addStyle(wb, "cc_ped_psnu", style = pct, cols=c(8,11,13,15,17,19), rows = 2:(nrow(cc_ped_psnu)+1), gridExpand=TRUE)


addWorksheet(wb, "cc_ped_site",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb, "cc_ped_site", cc_ped_site, tableStyle = "TableStyleLight4")
addStyle(wb, "cc_ped_site", style = pct, cols=c(11,14,14,16,18,20,22), rows = 2:(nrow(cc_ped_site)+1), gridExpand=TRUE)

saveWorkbook(wb, file = "C:/Users/josep/Documents/R/r_projects/mer/output/ct/clinical_care_cascade (peds).xlsx", overwrite = TRUE)
