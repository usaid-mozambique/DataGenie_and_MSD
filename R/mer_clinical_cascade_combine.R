rm(list = ls())
library(tidyr)
library(stringr)
library(dplyr)
library(stringr)
library(tibble)
library(janitor)
library(readxl)
library(openxlsx)

#########################added by Cicero: but is not necessary
#library(rlang)
#library(devtools)
#devtools::install_version('textfeatures', version='0.2.0', repos='http://cran.us.r-project.org')


##############################################
# IMPORT SOURCE FILES
##############################################

#---------------------------------------------
# GENIE IMPORT AND SET UP OF BASE DATAFRAME. NOTE THAT "FILTER" AND "SELECT" CODE LINES ARE ANALYSIS SPECIFIC 
df_complete <- read.delim("C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q4 Pre-Cleaning/MER_Structured_Datasets_SITE_IM_FY17-20_20191115_v1_1_Mozambique.txt")

df <- read.delim("C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q4 Pre-Cleaning/MER_Structured_Datasets_SITE_IM_FY17-20_20191115_v1_1_Mozambique.txt") %>% ####### NEEDS TO BE MANUALLY IMPORTED THROUGH MENU WITH EACH DATA GENIE UPDATE
  filter(fiscal_year == "2019") %>%
  rename(target = targets,
         site_id = facilityuid, 
         disag = standardizeddisaggregate,
         agency = fundingagency,
         site = sitename,
         partner = mech_name,
         Q = starts_with("Qtr4") ####### DEFINE RESULTS PERIOD DESIRED
  )

#This line was a request to from Tracy to extract HRH_CURR, not needed here
#write.csv(df, file = "df_HRH.xlsx")


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

ajuda <- read_excel("C:/Users/cnhantumbo/Documents/AJUDA/ajuda_ia_phase_disag.xlsx") %>%
  replace_na(list(ajuda = 0))


##############################################
# CREATE SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", target > 0) %>%
  select(site_id, target) %>%
  group_by(site_id) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
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
  
# Error caused because is missing site_id in ajuda dataframe  
  
#  left_join(tx_curr_target, by = "site_id") %>%
#  left_join(hts_tst, by = "site_id") %>%
#  left_join(hts_tst_pos, by = "site_id") %>%
#  left_join(tx_new, by = "site_id") %>%
#  left_join(tx_pvls_d, by = "site_id") %>%
#  left_join(tx_pvls_n, by  = "site_id") %>%
# left_join(ajuda, by = "site_id") %>%
#  replace_na(list(hts_tst = 0,
#                  hts_tst_pos = 0,
#                  tx_new = 0,
#                  tx_curr = 0,
#                  tx_pvls_d = 0,
#                  tx_pvls_n = 0)
#                 ajuda = 0)
#  ) %>%

  
  mutate(per_hts_tst_pos = hts_tst_pos / hts_tst,
         per_linkage = tx_new / hts_tst_pos,
         per_tx_curr_target = tx_curr / tx_curr_target,
         per_tx_pvl = tx_pvls_d / tx_curr,
         per_tx_pvls = tx_pvls_n / tx_pvls_d
  ) %>%
  select(site_id, agency, partner, snu1, psnu, site, ajuda, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls) %>%
  arrange(agency, partner, snu1, psnu, site)

# Warning message: Column `site_id` joining factor and character vector, coercing into character vector

# select(site_id, agency, partner, snu1, psnu, site, tx_curr_target, hts_tst, hts_tst_pos, per_hts_tst_pos, tx_new, per_linkage, tx_curr, per_tx_curr_target, tx_pvls_d, per_tx_pvl, tx_pvls_n, per_tx_pvls) %>%
#  arrange(agency, partner, snu1, psnu, site)


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

tx_curr_target <- df %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", target > 0) %>%
  select(partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, target) %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# HTS REFERENCE OBJECT

hts_tst <- df %>%
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

hts_tst_pos <- df %>%
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

tx_new <- df %>%
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

tx_curr <- df %>%
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

tx_pvls_d <- df %>%
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

tx_pvls_n <- df %>%
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

#addStyle(wb, "cc_agency", style = ctr, cols=c(2:13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_overall <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")

addWorksheet(wb_overall, "cc_ou",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_ou", cc_ou, tableStyle = "TableStyleLight6")
addStyle(wb_overall, "cc_ou", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "cc_ajuda",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_ajuda", cc_ajuda, tableStyle = "TableStyleLight6")
addStyle(wb_overall, "cc_ajuda", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "cc_agency",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_agency", cc_agency, tableStyle = "TableStyleLight6")
addStyle(wb_overall, "cc_agency", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)


addWorksheet(wb_overall, "cc_partner",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_partner", cc_partner, tableStyle = "TableStyleLight6")
addStyle(wb_overall, "cc_partner", style = pct, cols=c(6,8,10,12,14), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)


addWorksheet(wb_overall, "cc_snu1",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_snu1", cc_snu1, tableStyle = "TableStyleLight6")
addStyle(wb_overall, "cc_snu1", style = pct, cols=c(7,9,11,13,15), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_overall, "cc_psnu",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_psnu", cc_psnu, tableStyle = "TableStyleLight6")
addStyle(wb_overall, "cc_psnu", style = pct, cols=c(8,10,12,14,16), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)


addWorksheet(wb_overall, "cc_site",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall, "cc_site", cc_site, tableStyle = "TableStyleLight6")
addStyle(wb_overall, "cc_site", style = pct, cols=c(11,13,15,17,19), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)



#############################################
# BEGIN FEMALE CASCADE
##############################################

df_fem <- df %>%
  filter(fiscal_year == "2019", sex == "Female")

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

hts_tst_pos <- df_fem %>%
  filter(indicator == "HTS_TST_POS",
         disag == "Modality/Age/Sex/Result",
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

addStyle(wb, "cc_agency", style = ctr, cols=c(2:13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_female <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")

addWorksheet(wb_female, "cc_ou",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_ou", cc_ou, tableStyle = "TableStyleLight5")
addStyle(wb_female, "cc_ou", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)

addWorksheet(wb_female, "cc_ajuda",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_ajuda", cc_ajuda, tableStyle = "TableStyleLight5")
addStyle(wb_female, "cc_ajuda", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_female, "cc_agency",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_agency", cc_agency, tableStyle = "TableStyleLight5")
addStyle(wb_female, "cc_agency", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)


addWorksheet(wb_female, "cc_partner",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_partner", cc_partner, tableStyle = "TableStyleLight5")
addStyle(wb_female, "cc_partner", style = pct, cols=c(6,8,10,12,14), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)


addWorksheet(wb_female, "cc_snu1",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_snu1", cc_snu1, tableStyle = "TableStyleLight5")
addStyle(wb_female, "cc_snu1", style = pct, cols=c(7,9,11,13,15), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_female, "cc_psnu",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_psnu", cc_psnu, tableStyle = "TableStyleLight5")
addStyle(wb_female, "cc_psnu", style = pct, cols=c(8,10,12,14,16), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)


addWorksheet(wb_female, "cc_site",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female, "cc_site", cc_site, tableStyle = "TableStyleLight5")
addStyle(wb_female, "cc_site", style = pct, cols=c(11,13,15,17,19), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)



#############################################
# BEGIN MALE CASCADE
##############################################

df_mal <- df %>%
  filter(fiscal_year == "2019", sex == "Male")

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

#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df_mal %>%
  filter(indicator == "HTS_TST",
         disag == "Modality/Age/Sex/Result",
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

hts_tst_pos <- df_mal %>%
  filter(indicator == "HTS_TST_POS",
         disag == "Modality/Age/Sex/Result",
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

##############################################
# GENERATE EXCEL OUTPUTS
##############################################

addStyle(wb, "cc_agency", style = ctr, cols=c(2:13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_male <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")

addWorksheet(wb_male, "cc_ou",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_ou", cc_ou, tableStyle = "TableStyleLight3")
addStyle(wb_male, "cc_ou", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)

addWorksheet(wb_male, "cc_ajuda",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_ajuda", cc_ajuda, tableStyle = "TableStyleLight3")
addStyle(wb_male, "cc_ajuda", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_male, "cc_agency",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_agency", cc_agency, tableStyle = "TableStyleLight3")
addStyle(wb_male, "cc_agency", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)


addWorksheet(wb_male, "cc_partner",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_partner", cc_partner, tableStyle = "TableStyleLight3")
addStyle(wb_male, "cc_partner", style = pct, cols=c(6,8,10,12,14), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)


addWorksheet(wb_male, "cc_snu1",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_snu1", cc_snu1, tableStyle = "TableStyleLight3")
addStyle(wb_male, "cc_snu1", style = pct, cols=c(7,9,11,13,15), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_male, "cc_psnu",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_psnu", cc_psnu, tableStyle = "TableStyleLight3")
addStyle(wb_male, "cc_psnu", style = pct, cols=c(8,10,12,14,16), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)


addWorksheet(wb_male, "cc_site",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male, "cc_site", cc_site, tableStyle = "TableStyleLight3")
addStyle(wb_male, "cc_site", style = pct, cols=c(11,13,15,17,19), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)



#############################################
# BEGIN PEDIATRIC CASCADE
##############################################

df_ped <- df %>%
  filter(fiscal_year == "2019", trendscoarse == "<15")

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

hts_tst_pos <- df_ped %>%
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

addWorksheet(wb_pediatrico, "cc_ped_ou",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_ou", cc_ped_ou, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico, "cc_ped_ou", style = pct, cols=c(5,8,10,12,14,16), rows = 2:(nrow(cc_ped_ou)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "cc_ped_ajuda",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_ajuda", cc_ped_ajuda, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico, "cc_ped_ajuda", style = pct, cols=c(5,8,10,12,14,16), rows = 2:(nrow(cc_ped_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "cc_ped_agency",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_agency", cc_ped_agency, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico, "cc_ped_agency", style = pct, cols=c(5,8,10,12,14,16), rows = 2:(nrow(cc_ped_agency)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "cc_ped_partner",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_partner", cc_ped_partner, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico, "cc_ped_partner", style = pct, cols=c(6,9,11,13,15,17), rows = 2:(nrow(cc_ped_partner)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "cc_ped_snu1",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_snu1", cc_ped_snu1, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico, "cc_ped_snu1", style = pct, cols=c(7,10,12,14,16,18), rows = 2:(nrow(cc_ped_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_pediatrico, "cc_ped_psnu",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_psnu", cc_ped_psnu, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico, "cc_ped_psnu", style = pct, cols=c(8,11,13,15,17,19), rows = 2:(nrow(cc_ped_psnu)+1), gridExpand=TRUE)


addWorksheet(wb_pediatrico, "cc_ped_site",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "cc_ped_site", cc_ped_site, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico, "cc_ped_site", style = pct, cols=c(11,14,14,16,18,20,22), rows = 2:(nrow(cc_ped_site)+1), gridExpand=TRUE)

##############################################
# GENERATE CUMULATIVE CASCADE ANALYSES
##############################################

#---------------------------------------------
# GENIE IMPORT AND SET UP OF BASE DATAFRAME. NOTE THAT "FILTER" AND "SELECT" CODE LINES ARE ANALYSIS SPECIFIC 

df_cum <- df %>%
  select(-Q) %>%
  rename(
         Q = starts_with("Cumulative") ####### DEFINE RESULTS PERIOD DESIRED 
  )


##############################################
# CREATE SITE REFERENCE OBJECTS
##############################################

#---------------------------------------------
# TX_CURR TARGET REFERENCE OBJECT

tx_curr_target <- df_cum %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", target > 0) %>%
  select(site_id, target) %>%
  group_by(site_id) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# TX_CURR REFERENCE OBJECT

tx_curr <- df_cum %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q) %>%
  group_by(site_id, agency, partner, snu1, psnu, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr = starts_with("Q")) 

#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df_cum %>%
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

hts_tst_pos <- df_cum %>%
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

tx_new <- df_cum %>%
  filter(indicator == "TX_NEW", disag == "Total Numerator", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_new = starts_with("Q"))

#---------------------------------------------
# TX_PVLS DENOMINATOR REFERENCE OBJECT

tx_pvls_d <- df_cum %>%
  filter(indicator == "TX_PVLS", disag == "Total Denominator", Q > 0) %>%
  select(site_id, Q) %>%
  group_by(site_id) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_pvls_d = starts_with("Q"))

#---------------------------------------------
# TX_PVLS NUMERATOR REFERENCE OBJECT

tx_pvls_n <- df_cum %>%
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

tx_curr_target <- df_cum %>%
  filter(indicator == "TX_CURR", disag == "Total Numerator", target > 0) %>%
  select(partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, target) %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(target = sum(target, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(tx_curr_target = "target") 

#---------------------------------------------
# HTS REFERENCE OBJECT

hts_tst <- df_cum %>%
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

hts_tst_pos <- df_cum %>%
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

tx_new <- df_cum %>%
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

tx_curr <- df_cum %>%
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

tx_pvls_d <- df_cum %>%
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

tx_pvls_n <- df_cum %>%
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

#addStyle(wb, "cc_agency", style = ctr, cols=c(2:13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_overall_cum <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")

addWorksheet(wb_overall_cum, "cc_ou",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall_cum, "cc_ou", cc_ou, tableStyle = "TableStyleLight6")
addStyle(wb_overall_cum, "cc_ou", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)

addWorksheet(wb_overall_cum, "cc_ajuda",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall_cum, "cc_ajuda", cc_ajuda, tableStyle = "TableStyleLight6")
addStyle(wb_overall_cum, "cc_ajuda", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_overall_cum, "cc_agency",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall_cum, "cc_agency", cc_agency, tableStyle = "TableStyleLight6")
addStyle(wb_overall_cum, "cc_agency", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)


addWorksheet(wb_overall_cum, "cc_partner",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall_cum, "cc_partner", cc_partner, tableStyle = "TableStyleLight6")
addStyle(wb_overall_cum, "cc_partner", style = pct, cols=c(6,8,10,12,14), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)


addWorksheet(wb_overall_cum, "cc_snu1",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall_cum, "cc_snu1", cc_snu1, tableStyle = "TableStyleLight6")
addStyle(wb_overall_cum, "cc_snu1", style = pct, cols=c(7,9,11,13,15), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_overall_cum, "cc_psnu",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall_cum, "cc_psnu", cc_psnu, tableStyle = "TableStyleLight6")
addStyle(wb_overall_cum, "cc_psnu", style = pct, cols=c(8,10,12,14,16), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)


addWorksheet(wb_overall_cum, "cc_site",
             tabColour = "#33CCCC",
             gridLines = FALSE)
writeDataTable(wb_overall_cum, "cc_site", cc_site, tableStyle = "TableStyleLight6")
addStyle(wb_overall_cum, "cc_site", style = pct, cols=c(11,13,15,17,19), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)



#############################################
# BEGIN FEMALE CASCADE
##############################################

df_fem <- df_cum %>%
  filter(fiscal_year == "2019", sex == "Female")

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

hts_tst_pos <- df_fem %>%
  filter(indicator == "HTS_TST_POS",
         disag == "Modality/Age/Sex/Result",
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

addStyle(wb, "cc_agency", style = ctr, cols=c(2:13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_female_cum <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")

addWorksheet(wb_female_cum, "cc_ou",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female_cum, "cc_ou", cc_ou, tableStyle = "TableStyleLight5")
addStyle(wb_female_cum, "cc_ou", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)

addWorksheet(wb_female_cum, "cc_ajuda",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female_cum, "cc_ajuda", cc_ajuda, tableStyle = "TableStyleLight5")
addStyle(wb_female_cum, "cc_ajuda", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_female_cum, "cc_agency",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female_cum, "cc_agency", cc_agency, tableStyle = "TableStyleLight5")
addStyle(wb_female_cum, "cc_agency", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)


addWorksheet(wb_female_cum, "cc_partner",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female_cum, "cc_partner", cc_partner, tableStyle = "TableStyleLight5")
addStyle(wb_female_cum, "cc_partner", style = pct, cols=c(6,8,10,12,14), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)


addWorksheet(wb_female_cum, "cc_snu1",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female_cum, "cc_snu1", cc_snu1, tableStyle = "TableStyleLight5")
addStyle(wb_female_cum, "cc_snu1", style = pct, cols=c(7,9,11,13,15), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_female_cum, "cc_psnu",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female_cum, "cc_psnu", cc_psnu, tableStyle = "TableStyleLight5")
addStyle(wb_female_cum, "cc_psnu", style = pct, cols=c(8,10,12,14,16), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)


addWorksheet(wb_female_cum, "cc_site",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeDataTable(wb_female_cum, "cc_site", cc_site, tableStyle = "TableStyleLight5")
addStyle(wb_female_cum, "cc_site", style = pct, cols=c(11,13,15,17,19), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)



#############################################
# BEGIN MALE CASCADE
##############################################

df_mal <- df_cum %>%
  filter(fiscal_year == "2019", sex == "Male")

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

#---------------------------------------------
# HTS_TST REFERENCE OBJECT

hts_tst <- df_mal %>%
  filter(indicator == "HTS_TST",
         disag == "Modality/Age/Sex/Result",
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

hts_tst_pos <- df_mal %>%
  filter(indicator == "HTS_TST_POS",
         disag == "Modality/Age/Sex/Result",
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

##############################################
# GENERATE EXCEL OUTPUTS
##############################################

addStyle(wb, "cc_agency", style = ctr, cols=c(2:13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)

#---------------------------------------------
##CREATE EXCEL WORKBOOK WITH TX_CURR OUTPUT

wb_male_cum <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")

addWorksheet(wb_male_cum, "cc_ou",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male_cum, "cc_ou", cc_ou, tableStyle = "TableStyleLight3")
addStyle(wb_male_cum, "cc_ou", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ou)+1), gridExpand=TRUE)

addWorksheet(wb_male_cum, "cc_ajuda",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male_cum, "cc_ajuda", cc_ajuda, tableStyle = "TableStyleLight3")
addStyle(wb_male_cum, "cc_ajuda", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_male_cum, "cc_agency",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male_cum, "cc_agency", cc_agency, tableStyle = "TableStyleLight3")
addStyle(wb_male_cum, "cc_agency", style = pct, cols=c(5,7,9,11,13), rows = 2:(nrow(cc_agency)+1), gridExpand=TRUE)


addWorksheet(wb_male_cum, "cc_partner",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male_cum, "cc_partner", cc_partner, tableStyle = "TableStyleLight3")
addStyle(wb_male_cum, "cc_partner", style = pct, cols=c(6,8,10,12,14), rows = 2:(nrow(cc_partner)+1), gridExpand=TRUE)


addWorksheet(wb_male_cum, "cc_snu1",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male_cum, "cc_snu1", cc_snu1, tableStyle = "TableStyleLight3")
addStyle(wb_male_cum, "cc_snu1", style = pct, cols=c(7,9,11,13,15), rows = 2:(nrow(cc_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_male_cum, "cc_psnu",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male_cum, "cc_psnu", cc_psnu, tableStyle = "TableStyleLight3")
addStyle(wb_male_cum, "cc_psnu", style = pct, cols=c(8,10,12,14,16), rows = 2:(nrow(cc_psnu)+1), gridExpand=TRUE)


addWorksheet(wb_male_cum, "cc_site",
             tabColour = "#993366",
             gridLines = FALSE)
writeDataTable(wb_male_cum, "cc_site", cc_site, tableStyle = "TableStyleLight3")
addStyle(wb_male_cum, "cc_site", style = pct, cols=c(11,13,15,17,19), rows = 2:(nrow(cc_site)+1), gridExpand=TRUE)



#############################################
# BEGIN PEDIATRIC CASCADE
##############################################

df_ped <- df_cum %>%
  filter(fiscal_year == "2019", trendscoarse == "<15")

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

hts_tst_pos <- df_ped %>%
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

wb_pediatrico_cum <- createWorkbook()
pct <- createStyle(numFmt = "0%", textDecoration = "italic")

addWorksheet(wb_pediatrico_cum, "cc_ped_ou",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico_cum, "cc_ped_ou", cc_ped_ou, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico_cum, "cc_ped_ou", style = pct, cols=c(5,8,10,12,14,16), rows = 2:(nrow(cc_ped_ou)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico_cum, "cc_ped_ajuda",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico_cum, "cc_ped_ajuda", cc_ped_ajuda, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico_cum, "cc_ped_ajuda", style = pct, cols=c(5,8,10,12,14,16), rows = 2:(nrow(cc_ped_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico_cum, "cc_ped_agency",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico_cum, "cc_ped_agency", cc_ped_agency, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico_cum, "cc_ped_agency", style = pct, cols=c(5,8,10,12,14,16), rows = 2:(nrow(cc_ped_agency)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico_cum, "cc_ped_partner",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico_cum, "cc_ped_partner", cc_ped_partner, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico_cum, "cc_ped_partner", style = pct, cols=c(6,9,11,13,15,17), rows = 2:(nrow(cc_ped_partner)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico_cum, "cc_ped_snu1",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico_cum, "cc_ped_snu1", cc_ped_snu1, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico_cum, "cc_ped_snu1", style = pct, cols=c(7,10,12,14,16,18), rows = 2:(nrow(cc_ped_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_pediatrico_cum, "cc_ped_psnu",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico_cum, "cc_ped_psnu", cc_ped_psnu, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico_cum, "cc_ped_psnu", style = pct, cols=c(8,11,13,15,17,19), rows = 2:(nrow(cc_ped_psnu)+1), gridExpand=TRUE)


addWorksheet(wb_pediatrico_cum, "cc_ped_site",
             tabColour = "#99CC00",
             gridLines = FALSE)
writeDataTable(wb_pediatrico_cum, "cc_ped_site", cc_ped_site, tableStyle = "TableStyleLight4")
addStyle(wb_pediatrico_cum, "cc_ped_site", style = pct, cols=c(11,14,14,16,18,20,22), rows = 2:(nrow(cc_ped_site)+1), gridExpand=TRUE)



##############################################
# SAVE ALL CLINICAL CASCADE WORKBOOKS
##############################################

#saveWorkbook(wb_overall, file = "C:/Users/jlara/Documents/R/r_projects/mer/output/ct/fy20q1/cc_cascade_overall.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
#saveWorkbook(wb_female, file = "C:/Users/jlara/Documents/R/r_projects/mer/output/ct/fy20q1/cc_cascade_female.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
#saveWorkbook(wb_male, file = "C:/Users/jlara/Documents/R/r_projects/mer/output/ct/fy20q1/cc_cascade_male.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
#saveWorkbook(wb_pediatrico, file = "C:/Users/jlara/Documents/R/r_projects/mer/output/ct/fy20q1/cc_cascade_ped.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
#saveWorkbook(wb_overall_cum, file = "C:/Users/jlara/Documents/R/r_projects/mer/output/ct/fy20q1/cc_cascade_overall_cum.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
#saveWorkbook(wb_female_cum, file = "C:/Users/jlara/Documents/R/r_projects/mer/output/ct/fy20q1/cc_cascade_female_cum.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
#saveWorkbook(wb_male_cum, file = "C:/Users/jlara/Documents/R/r_projects/mer/output/ct/fy20q1/cc_cascade_male_cum.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
#saveWorkbook(wb_pediatrico_cum, file = "C:/Users/jlara/Documents/R/r_projects/mer/output/ct/fy20q1/cc_cascade_ped_cum.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!






