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
# PMTCT_STAT_DEN REFERENCE OBJECT

p_stat_d <- df %>%
  filter(indicator == "PMTCT_STAT", disag == "Total Denominator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_d = starts_with("Q"))


#---------------------------------------------
# PMTCT_STAT_NUM REFERENCE OBJECT

p_stat_n <- df %>%
  filter(indicator == "PMTCT_STAT", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_n = starts_with("Q")) %>%
  select(-partner, -agency, -SNU1, -PSNU, -site)


#---------------------------------------------
# PMTCT_STAT_POS REFERENCE OBJECT

p_stat_pos <- df %>%
  filter(indicator == "PMTCT_STAT", StatusHIV == "Positive", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_pos = starts_with("Q")) %>%
  select(-partner, -agency, -SNU1, -PSNU, -site)


#---------------------------------------------
# PMTCT_ART REFERENCE OBJECT

p_art <- df %>%
  filter(indicator == "PMTCT_ART", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_art = starts_with("Q")) %>%
  select(-partner, -agency, -SNU1, -PSNU, -site)


#---------------------------------------------
# NEWLY POSITIVE VS. KNOWN HIV+ AT ENTRY REFERENCE OBJECT

p_stat_pos_new <- df %>%
  filter(indicator == "PMTCT_STAT", StatusHIV == "Positive", disag_p == "Newly Identified", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_pos_new = starts_with("Q")) %>%
  select(-partner, -agency, -SNU1, -PSNU, -site)

p_stat_pos_alr <- df %>%
  filter(indicator == "PMTCT_STAT", StatusHIV == "Positive", disag_p == "Known at Entry", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_stat_pos_alr = starts_with("Q")) %>%
  select(-partner, -agency, -SNU1, -PSNU, -site)


#---------------------------------------------
# NEWLY STARTED ART VS. ON ART AT ENTRY REFERENCE OBJECT

p_art_new <- df %>%
  filter(indicator == "PMTCT_ART", disag_p == "Life-long ART, New", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_art_new = starts_with("Q")) %>%
  select(-partner, -agency, -SNU1, -PSNU, -site)

p_art_alr <- df %>%
  filter(indicator == "PMTCT_ART", disag_p == "Life-long ART, Already", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_art_alr = starts_with("Q")) %>%
  select(-partner, -agency, -SNU1, -PSNU, -site)


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

pc_agency <- pc %>%
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
  )


#---------------------------------------------
# PMTCT CASADE BY PARTNER

pc_partner <- pc %>%
  group_by(partner) %>%
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
  )



#---------------------------------------------
# PMTCT CASADE BY PROVINCE

pc_snu1 <- pc %>%
  group_by(partner, SNU1) %>%
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
  )

#---------------------------------------------
# PMTCT CASADE BY DISTRICT

pc_psnu <- pc %>%
  group_by(partner, SNU1, PSNU) %>%
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
  )


#---------------------------------------------
# PMTCT CASCADE BY SITE

pc_site <- pc %>%
  mutate(per_p_stat = p_stat_n / p_stat_d, 
         per_p_stat_pos = p_stat_pos / p_stat_n,
         per_p_art = p_art / p_stat_pos
  )


#---------------------------------------------
# REMOVE UNNEEDED OBJECTS AND PRINT PMTCT CASCADE TO .CSV

rm(p_stat_d, p_stat_n, p_stat_pos, p_art, p_stat_pos_new, p_stat_pos_alr, p_art_new, p_art_alr, ajuda, pc)


wb <- createWorkbook()
addWorksheet(wb, "pc_agency",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeData(wb, "pc_agency", pc_agency)
addWorksheet(wb, "pc_partner",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeData(wb, "pc_partner", pc_partner)
addWorksheet(wb, "pc_snu1",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeData(wb, "pc_snu1", pc_snu1)
addWorksheet(wb, "pc_psnu",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeData(wb, "pc_psnu", pc_psnu)
addWorksheet(wb, "pc_site",
             tabColour = "#CC99FF",
             gridLines = FALSE)
writeData(wb, "pc_site", pc_site)
saveWorkbook(wb, file = "C:/Users/josep/Documents/R/r_projects/mer/output/pmtct/pmtct_cascade.xlsx", overwrite = TRUE)









#---------------------------------------------
# IDENTIFICATION OF ILLOGICAL RELATIONSHIPS IN PMTCT CASCADE 

pc_flag1 <- pc %>% # FLAG 1 IDENTIFIES SITES WHERE PMTCT_STAT NUMERATOR IS > THAN DENOMINATOR
  mutate(flag1 = 
           if_else(p_stat_n > p_stat_d, 1, 0)) %>%
  filter(flag1 == 1)

pc_flag2 <- pc %>% # FLAG 2 IDENTIFIES SITES WHERE PMTCT_STAT_POS IS > THAN PMTCT_STAT NUMERATOR
  mutate(flag2 = 
           if_else(p_stat_pos > p_stat_n, 1, 0)) %>%
  filter(flag2 == 1)

pc_flag3 <- pc %>% # FLAG 3 IDENTIFIES SITES WHERE PMTCT_ART NUMERATOR IS > THAN PMTCT_STAT_POS
  mutate(flag3 = 
           if_else(p_art > p_stat_pos, 1, 0)) %>%
  filter(flag3 == 1)


