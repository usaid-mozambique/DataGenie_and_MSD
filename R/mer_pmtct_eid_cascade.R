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
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_d = starts_with("Q"))


#---------------------------------------------
# PMTCT_EID_NUM REFERENCE OBJECT

p_eid_n <- df %>%
  filter(indicator == "PMTCT_EID", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_n = starts_with("Q"))

#---------------------------------------------
# PMTCT_EID_NUM 2-12M REFERENCE OBJECT

p_eid_n_12 <- df %>%
  filter(indicator == "PMTCT_EID", TrendsSemiFine == "02 - 12 Months", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_n_12 = starts_with("Q"))


#---------------------------------------------
# PMTCT_EID_NUM <2M REFERENCE OBJECT

p_eid_n_2 <- df %>%
  filter(indicator == "PMTCT_EID", TrendsSemiFine == "<=02 Months", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_n_2 = starts_with("Q"))


#---------------------------------------------
# PMTCT_HEI_POS REFERENCE OBJECT

p_eid_hei_pos <- df %>%
  filter(indicator == "PMTCT_HEI_POS", disag == "Total Numerator", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_hei_pos = starts_with("Q"))


#---------------------------------------------
# PMTCT_HEI_POS_ART REFERENCE OBJECT

p_eid_hei_pos_art <- df %>%
  filter(indicator == "PMTCT_HEI_POS", categoryOptionComboName == "2 - 12 months , Positive, Receiving ART" | categoryOptionComboName == "<= 2 months, Positive, Receiving ART", Q > 0) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q = sum(Q, na.rm = TRUE)) %>%
  ungroup() %>%
  rename(p_eid_hei_pos_art = starts_with("Q"))


#---------------------------------------------
# JOIN ALL CASCADE INDICATOR OBJECTS INTO ONE DATAFRAME

pc_eid <- p_eid_d %>%
  full_join(p_eid_n, by = "site_id") %>%
  full_join(p_eid_n_2, by = "site_id") %>%
  full_join(p_eid_n_12, by = "site_id") %>%
  full_join(p_eid_hei_pos, by = "site_id") %>%
  full_join(p_eid_hei_pos_art, by = "site_id") %>%
  left_join(ajuda, by = "site_id") %>%
  replace_na(list(p_eid_d = 0,
                  p_eid_n = 0,
                  p_eid_n_2 = 0,
                  p_eid_n_12 = 0,
                  p_eid_hei_pos = 0,
                  p_eid_hei_pos_art = 0,
                  ajuda = 0)
  ) %>%
  select(1:6, ajuda, everything())


#---------------------------------------------
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
  mutate(SNU1.x = 
           if_else(is.na(SNU1.x), SNU1.y, 
                   if_else(is.na(SNU1.x), SNU1.x.x, 
                           if_else(is.na(SNU1.x), SNU1.y.y, SNU1.x)
                   )
           )
  ) %>%
  mutate(PSNU.x = 
           if_else(is.na(PSNU.x), PSNU.y, 
                   if_else(is.na(PSNU.x), PSNU.x.x, 
                           if_else(is.na(PSNU.x), PSNU.y.y, PSNU.x)
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
         SNU1 = SNU1.x,
         PSNU = PSNU.x,
         site = site.x
         ) %>%
  select(-ends_with(".x"), -ends_with(".y"))

#---------------------------------------------
# PMTCT CASADE BY SITE

pc_eid_site <- pc_eid_temp %>%
  mutate(per_p_eid = p_eid_n / p_eid_d, 
         per_p_eid_n_2 = p_eid_n_2 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_n_12 = p_eid_n_12 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_hei_pos_art = p_eid_hei_pos_art / p_eid_hei_pos
  ) %>%
  select(1:7, p_eid_d, p_eid_n, per_p_eid, p_eid_n_2, p_eid_n_12, per_p_eid_n_2, per_p_eid_n_12, everything())

#---------------------------------------------
# PMTCT CASADE BY DISTRICT

pc_eid_psnu <- pc_eid_temp %>%
  group_by(agency, partner, SNU1, PSNU) %>%
  summarize(p_eid_d = sum(p_eid_d, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),
            p_eid_n_2 = sum(p_eid_n_2, na.rm = TRUE),
            p_eid_n_12 = sum(p_eid_n_12, na.rm = TRUE),
            p_eid_hei_pos = sum(p_eid_hei_pos, na.rm = TRUE),
            p_eid_hei_pos_art = sum(p_eid_hei_pos_art, na.rm = TRUE)
            )%>%
  ungroup() %>%
  mutate(per_p_eid = p_eid_n / p_eid_d, 
         per_p_eid_n_2 = p_eid_n_2 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_n_12 = p_eid_n_12 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_hei_pos_art = p_eid_hei_pos_art / p_eid_hei_pos
  ) %>%
  select(2:5, p_eid_d, p_eid_n, per_p_eid, p_eid_n_2, p_eid_n_12, per_p_eid_n_2, per_p_eid_n_12, everything())

#---------------------------------------------
# PMTCT CASADE BY PROVINCE

pc_eid_snu1 <- pc_eid_temp %>%
  group_by(agency, partner, SNU1) %>%
  summarize(p_eid_d = sum(p_eid_d, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),
            p_eid_n_2 = sum(p_eid_n_2, na.rm = TRUE),
            p_eid_n_12 = sum(p_eid_n_12, na.rm = TRUE),
            p_eid_hei_pos = sum(p_eid_hei_pos, na.rm = TRUE),
            p_eid_hei_pos_art = sum(p_eid_hei_pos_art, na.rm = TRUE)
  )%>%
  ungroup() %>%
  mutate(per_p_eid = p_eid_n / p_eid_d, 
         per_p_eid_n_2 = p_eid_n_2 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_n_12 = p_eid_n_12 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_hei_pos_art = p_eid_hei_pos_art / p_eid_hei_pos
  ) %>%
  select(2:4, p_eid_d, p_eid_n, per_p_eid, p_eid_n_2, p_eid_n_12, per_p_eid_n_2, per_p_eid_n_12, everything())

#---------------------------------------------
# PMTCT CASADE BY PARTNER

pc_eid_partner <- pc_eid_temp %>%
  group_by(agency, partner) %>%
  summarize(p_eid_d = sum(p_eid_d, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),
            p_eid_n_2 = sum(p_eid_n_2, na.rm = TRUE),
            p_eid_n_12 = sum(p_eid_n_12, na.rm = TRUE),
            p_eid_hei_pos = sum(p_eid_hei_pos, na.rm = TRUE),
            p_eid_hei_pos_art = sum(p_eid_hei_pos_art, na.rm = TRUE)
  )%>%
  ungroup() %>%
  mutate(per_p_eid = p_eid_n / p_eid_d, 
         per_p_eid_n_2 = p_eid_n_2 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_n_12 = p_eid_n_12 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_hei_pos_art = p_eid_hei_pos_art / p_eid_hei_pos
  ) %>%
  select(2:3, p_eid_d, p_eid_n, per_p_eid, p_eid_n_2, p_eid_n_12, per_p_eid_n_2, per_p_eid_n_12, everything())

#---------------------------------------------
# PMTCT CASADE BY AGENCY

pc_eid_agency <- pc_eid_temp %>%
  group_by(agency) %>%
  summarize(p_eid_d = sum(p_eid_d, na.rm = TRUE),
            p_eid_n = sum(p_eid_n, na.rm = TRUE),
            p_eid_n_2 = sum(p_eid_n_2, na.rm = TRUE),
            p_eid_n_12 = sum(p_eid_n_12, na.rm = TRUE),
            p_eid_hei_pos = sum(p_eid_hei_pos, na.rm = TRUE),
            p_eid_hei_pos_art = sum(p_eid_hei_pos_art, na.rm = TRUE)
  )%>%
  ungroup() %>%
  mutate(per_p_eid = p_eid_n / p_eid_d, 
         per_p_eid_n_2 = p_eid_n_2 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_n_12 = p_eid_n_12 / (p_eid_n_2 + p_eid_n_12),
         per_p_eid_hei_pos_art = p_eid_hei_pos_art / p_eid_hei_pos
  ) %>%
  select(agency, p_eid_d, p_eid_n, per_p_eid, p_eid_n_2, p_eid_n_12, per_p_eid_n_2, per_p_eid_n_12, everything())


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

rm(pc_eid_temp, p_eid_d, p_eid_n, p_eid_n_2, p_eid_n_12, p_eid_hei_pos, p_eid_hei_pos_art)


wb <- createWorkbook()
addWorksheet(wb, "pc_eid_agency",
             tabColour = "#0000FF",
             gridLines = FALSE)
writeData(wb, "pc_eid_agency", pc_eid_agency)
addWorksheet(wb, "pc_eid_partner",
             tabColour = "#0000FF",
             gridLines = FALSE)
writeData(wb, "pc_eid_partner", pc_eid_partner)
addWorksheet(wb, "pc_snu1",
             tabColour = "#0000FF",
             gridLines = FALSE)
writeData(wb, "pc_eid_snu1", pc_eid_snu1)
addWorksheet(wb, "pc_eid_psnu",
             tabColour = "#0000FF",
             gridLines = FALSE)
writeData(wb, "pc_eid_psnu", pc_eid_psnu)
addWorksheet(wb, "pc_eid_site",
             tabColour = "#0000FF",
             gridLines = FALSE)
writeData(wb, "pc_eid_site", pc_eid_site)
saveWorkbook(wb, file = "C:/Users/josep/Documents/R/r_projects/mer/output/pmtct/pmtct__eid_cascade.xlsx", overwrite = TRUE)

