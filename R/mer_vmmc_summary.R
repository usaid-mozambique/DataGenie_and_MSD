rm(list = ls())

library(tidyverse)
library(janitor)
library(readxl)
library(openxlsx)


#---------------------------------------------
# GENIE IMPORT AND SET UP OF BASE DATAFRAME. NOTE THAT "FILTER" AND "SELECT" CODE LINES ARE ANALYSIS SPECIFIC 

df <- read.delim("~/R/datasets/Genie_Daily_82e86f5aaba344fa8e890f3ff91fa5ea.txt") %>% ####### NEEDS TO BE MANUALLY IMPORTED THROUGH MENU WITH EACH DATA GENIE UPDATE
  rename(target = TARGETS, 
         site_id = orgUnitUID, 
         disag = standardizedDisaggregate, 
         agency = FundingAgency, 
         site = SiteName, 
         partner = mech_name, 
         Q1 = Qtr1, 
         Q2 = Qtr2, 
         Q3 = Qtr3, 
         Q4 = Qtr4,
         YtD = Cumulative)

df$partner <- as.character(df$partner)

df <- df %>%
  mutate(partner = 
           if_else(partner == "FADM Prevention and Circumcision Program", "FADM",
                   if_else(partner == "Johns Hopkins", "Jhpiego",
                           if_else(partner == "Strengthening High Impact Interventions for an AIDS-Free Generation (AIDSFree) Project", "AIDS-Free", partner)
                   )
           )
  )

#---------------------------------------------
##SITE LEVEL VMMC_CIRC PERFORMANCE

vmmc_circ_site <- df %>%
  filter(indicator == "VMMC_CIRC", 
         disag == "Total Numerator",
         Fiscal_Year == "2019",
         (YtD > 0 | target > 0)) %>%
  select(site_id, agency, partner, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, YtD, target) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
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
##DISTRICT LEVEL VMMC_CIRC PERFORMANCE

vmmc_circ_psnu <- df %>%
  filter(indicator == "VMMC_CIRC", 
         disag == "Total Numerator",
         Fiscal_Year == "2019",
         (YtD > 0 | target > 0)) %>%
  select(site_id, agency, partner, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, YtD, target) %>%
  group_by(agency, partner, SNU1, PSNU) %>%
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

vmmc_circ_snu1 <- df %>%
  filter(indicator == "VMMC_CIRC", 
         disag == "Total Numerator",
         Fiscal_Year == "2019",
         (YtD > 0 | target > 0)) %>%
  select(site_id, agency, partner, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, YtD, target) %>%
  group_by(agency, partner, SNU1) %>%
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

vmmc_circ_partner <- df %>%
  filter(indicator == "VMMC_CIRC", 
         disag == "Total Numerator",
         Fiscal_Year == "2019",
         (YtD > 0 | target > 0)) %>%
  select(site_id, agency, partner, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, YtD, target) %>%
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

vmmc_circ_agency <- df %>%
  filter(indicator == "VMMC_CIRC", 
         disag == "Total Numerator",
         Fiscal_Year == "2019",
         (YtD > 0 | target > 0)) %>%
  select(site_id, agency, partner, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, YtD, target) %>%
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
addWorksheet(wb, "vmmc_circ_agency",
             tabColour = "#FF00FF",
             gridLines = FALSE)
writeData(wb, "vmmc_circ_agency", vmmc_circ_agency)
addWorksheet(wb, "vmmc_circ_partner",
             tabColour = "#FF00FF",
             gridLines = FALSE)
writeData(wb, "vmmc_circ_partner", vmmc_circ_partner)
addWorksheet(wb, "vmmc_circ_snu1",
             tabColour = "#FF00FF",
             gridLines = FALSE)
writeData(wb, "vmmc_circ_snu1", vmmc_circ_snu1)
addWorksheet(wb, "vmmc_circ_psnu",
             tabColour = "#FF00FF",
             gridLines = FALSE)
writeData(wb, "vmmc_circ_psnu", vmmc_circ_psnu)
addWorksheet(wb, "vmmc_circ_site",
             tabColour = "#FF00FF",
             gridLines = FALSE)
writeData(wb, "vmmc_circ_site", vmmc_circ_site)
saveWorkbook(wb, file = "C:/Users/josep/Documents/R/r_projects/mer/output/vmmc/vmmc_summary.xlsx", overwrite = TRUE)



