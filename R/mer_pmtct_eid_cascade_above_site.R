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
         disag_p = otherDisaggregate,
         agency = FundingAgency,
         site = SiteName,
         partner = mech_name,
         Q = starts_with("Q")
         )

df$partner <- as.character(df$partner)

df <- df %>%
  mutate(partner = 
           if_else(partner == "Clinical Services System Strenghening (CHASS)", "CHASS",
                   if_else(partner == "Friends in Global Health", "FGH", partner)))

#---------------------------------------------
# IMPORT OF COP19 AJUDA SITE LIST

ajuda <- read_excel("R/datasets/ajuda_ia.xlsx") %>%
  replace_na(list(ajuda = 0))


#---------------------------------------------
# PMTCT_EID_DEN REFERENCE OBJECT

p_eid_d <- df %>%
  filter(indicator == "PMTCT_EID", disag == "Total Denominator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, Fiscal_Year, Q) %>%
  group_by(agency, partner, SNU1, PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_d = starts_with("Q"))


#---------------------------------------------
# PMTCT_EID_NUM REFERENCE OBJECT

p_eid_n <- df %>%
  filter(indicator == "PMTCT_EID", disag == "Total Numerator", Q > 0) %>%
  select(PSNU, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_n = starts_with("Q"))

#---------------------------------------------
# PMTCT_EID_NUM 2-12M REFERENCE OBJECT

p_eid_n_12 <- df %>%
  filter(indicator == "PMTCT_EID", TrendsSemiFine == "02 - 12 Months", Q > 0) %>%
  select(PSNU, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_n_12 = starts_with("Q"))


#---------------------------------------------
# PMTCT_EID_NUM <2M REFERENCE OBJECT

p_eid_n_2 <- df %>%
  filter(indicator == "PMTCT_EID", TrendsSemiFine == "<=02 Months", Q > 0) %>%
  select(PSNU, Q) %>%
  group_by(PSNU) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_n_2 = starts_with("Q"))


#---------------------------------------------
# JOIN ALL CASCADE INDICATOR OBJECTS INTO ONE DATAFRAME

pc_eid <- p_eid_d %>%
  left_join(p_eid_n, by = "PSNU") %>%
  left_join(p_eid_n_2, by = "PSNU") %>%
  left_join(p_eid_n_12, by = "PSNU") %>%
  replace_na(list(p_eid_d = 0,
                  p_eid_n = 0,
                  p_eid_n_2 = 0,
                  p_eid_n_12 = 0,
                  ajuda = 0)
  )



#---------------------------------------------
# PMTCT CASADE BY AGENCY

pc_eid_agency <- pc_eid %>%
  group_by(agency) %>%
  summarize(p_eid_d = sum(p_eid_d, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),
            p_eid_n_2 = sum(p_eid_n_2, na.rm = TRUE),
            p_eid_n_12 = sum(p_eid_n_12, na.rm = TRUE)
            ) %>%
  ungroup() %>%
  mutate(per_p_eid = p_eid_n / p_eid_d, 
         per_p_eid_n_2 = p_eid_n_2 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_n_12 = p_eid_n_12 / (p_eid_n_2 + p_eid_n_12)
  ) %>%
  select(1:3, p_eid_d, p_eid_n, per_p_eid, everything())


#---------------------------------------------
# PMTCT CASADE BY PARTNER

pc_eid_partner <- pc_eid %>%
  group_by(agency, partner) %>%
  summarize(p_eid_d = sum(p_eid_d, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),
            p_eid_n_2 = sum(p_eid_n_2, na.rm = TRUE),
            p_eid_n_12 = sum(p_eid_n_12, na.rm = TRUE)
            ) %>%
  ungroup() %>%
  mutate(per_p_eid = p_eid_n / p_eid_d, 
         per_p_eid_n_2 = p_eid_n_2 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_n_12 = p_eid_n_12 / (p_eid_n_2 + p_eid_n_12)
  ) %>%
  select(1:4, p_eid_d, p_eid_n, per_p_eid, everything())


#---------------------------------------------
# PMTCT CASADE BY PROVINCE

pc_eid_snu1 <- pc_eid %>%
  group_by(agency, partner, SNU1) %>%
  summarize(p_eid_d = sum(p_eid_d, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),
            p_eid_n_2 = sum(p_eid_n_2, na.rm = TRUE),
            p_eid_n_12 = sum(p_eid_n_12, na.rm = TRUE)
  ) %>%
  ungroup() %>%
  mutate(per_p_eid = p_eid_n / p_eid_d, 
         per_p_eid_n_2 = p_eid_n_2 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_n_12 = p_eid_n_12 / (p_eid_n_2 + p_eid_n_12)
  ) %>%
  select(1:5, p_eid_d, p_eid_n, per_p_eid, everything())


#---------------------------------------------
# PMTCT CASADE BY DISTRICT

pc_eid_psnu <- pc_eid %>%
  mutate(per_p_eid = p_eid_n / p_eid_d, 
         per_p_eid_n_2 = p_eid_n_2 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_n_12 = p_eid_n_12 / (p_eid_n_2 + p_eid_n_12)
  ) %>%
  select(1:6, p_eid_d, p_eid_n, per_p_eid, everything())



#---------------------------------------------
# REMOVE UNNEEDED OBJECTS AND PRINT PMTCT CASCADE TO .CSV

rm(p_eid_d, p_eid_n, p_eid_n_12, p_eid_n_2, pc_eid)


wb <- createWorkbook()
addWorksheet(wb, "pc_eid_agency",
             tabColour = "#CCCCFF",
             gridLines = FALSE)
writeData(wb, "pc_eid_agency", pc_eid_agency)
addWorksheet(wb, "pc_eid_partner",
             tabColour = "#CCCCFF",
             gridLines = FALSE)
writeData(wb, "pc_eid_partner", pc_eid_partner)
addWorksheet(wb, "pc_eid_snu1",
             tabColour = "#CCCCFF",
             gridLines = FALSE)
writeData(wb, "pc_eid_snu1", pc_eid_snu1)
addWorksheet(wb, "pc_eid_psnu",
             tabColour = "#CCCCFF",
             gridLines = FALSE)
writeData(wb, "pc_eid_psnu", pc_eid_psnu)
saveWorkbook(wb, file = "C:/Users/josep/Documents/R/r_projects/mer/output/pmtct/pmtct_eid_cascade.xlsx", overwrite = TRUE)




