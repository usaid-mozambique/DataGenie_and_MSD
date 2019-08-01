rm(list = ls())

library(tidyverse)
library(janitor)
library(devtools)
#devtools::install_github("tidyverse/tidyr")
library(tidyr)
#install.packages("magrittr") 
library(magrittr)
library(dplyr)
library(ellipsis)
library(purrr)
library(rlang)
library(stringi)
library(tibble)
library(vctrs)
library(assertthat)
library(BH)
library(backports)
library(digest)


#mer <- read.delim("~/R/datasets/Genie_Daily_da918404311247cba21f6b73107db8fd.txt") %>% ####### NEEDS TO BE MANUALLY IMPORTED THROUGH MENU WITH EACH DATA GENIE UPDATE
#  rename(target = TARGETS, site_id = orgUnitUID, disag = standardizedDisaggregate, agency = FundingAgency, site = SiteName, partner = PrimePartner, Q1 = Qtr1, Q2 = Qtr2, Q3 = Qtr3, Q4 = Qtr4)

mer <- read.delim("C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q2 (Post-Cleaning)/MER_Structured_Dataset_Site_IM_FY17-19_20190621_v2_1_Mozambique.txt") %>% ####### NEEDS TO BE MANUALLY IMPORTED THROUGH MENU WITH EACH DATA GENIE UPDATE
  rename(target = TARGETS, site_id = orgUnitUID, disag = standardizedDisaggregate, agency = FundingAgency, site = SiteName, partner = PrimePartner, Q1 = Qtr1, Q2 = Qtr2, Q3 = Qtr3, Q4 = Qtr4)




##################################
##################################
##  TX_CURR PERFORMANCE TABLES  ##
##################################
##################################

## PREPARE TX_CURR DATASET
tx_curr_base <- mer %>%
  filter(indicator == "TX_CURR" & Fiscal_Year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target)
  
##AGENCY LEVEL TX_CURR PERFORMANCE
tx_curr <- tx_curr_base %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()

write_csv(tx_curr, "C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q2 (Post-Cleaning)/tx_curr_agency.csv", append = FALSE)

##PROVINCE LEVEL TX_CURR PERFORMANCE
tx_curr <- tx_curr_base %>%
  group_by(agency, partner, SNU1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()

write_csv(tx_curr, "C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q2 (Post-Cleaning)/tx_curr_snu1.csv", append = FALSE)

##DISTRICT LEVEL TX_CURR PERFORMANCE
tx_curr <- tx_curr_base %>%
  group_by(agency, partner, SNU1, PSNU) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()

write_csv(tx_curr, "C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q2 (Post-Cleaning)/tx_curr_psnu.csv", append = FALSE)

##SITE LEVEL TX_CURR PERFORMANCE
tx_curr <- tx_curr_base %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()

write_csv(tx_curr, "C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q2 (Post-Cleaning)/tx_curr_site.csv", append = FALSE)

rm(tx_curr_base)

##################################
#################################
##  TX_NEW PERFORMANCE TABLES  ##
#################################
##################################

##AGENCY LEVEL TX_NEW PERFORMANCE
tx_new <- mer %>%
  filter(indicator == "TX_NEW" & Fiscal_Year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q_vs_Target = (Q3)/(target / 4)) %>% ####### QUARTER SPECIFIC CODE
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()

write_csv(tx_new, "C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q2 (Post-Cleaning)/tx_new_agency.csv", append = FALSE)

##PROVINCE LEVEL TX_NEW PERFORMANCE
tx_new <- mer %>%
  filter(indicator == "TX_NEW" & Fiscal_Year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(agency, partner, SNU1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q_vs_Target = (Q3)/(target / 4)) %>% ####### QUARTER SPECIFIC CODE
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()

write_csv(tx_new, "C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q2 (Post-Cleaning)/tx_new_snu1.csv", append = FALSE)

##DISTRICT LEVEL TX_NEW PERFORMANCE
tx_new <- mer %>%
  filter(indicator == "TX_NEW" & Fiscal_Year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(agency, partner, SNU1, PSNU) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q_vs_Target = (Q3)/(target / 4)) %>% ####### QUARTER SPECIFIC CODE
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()

write_csv(tx_new, "C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q2 (Post-Cleaning)/tx_new_psnu.csv", append = FALSE)

##SITE LEVEL TX_NEW PERFORMANCE
tx_new <- mer %>%
  filter(indicator == "TX_NEW" & Fiscal_Year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4, target) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q_vs_Target = (Q3)/(target / 4)) %>% ####### QUARTER SPECIFIC CODE
  mutate(Cum_vs_Target = (Q1 + Q2 + Q3 + Q4)/(target)) %>%
  ungroup()

write_csv(tx_new, "C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q2 (Post-Cleaning)/tx_new_site.csv", append = FALSE)

#####################################
#####################################
##  PTMCT_STAT PERFORMANCE TABLES  ##
#####################################
#####################################

pmtct_stat_base <- mer %>%
  filter(indicator == "PMTCT_STAT" & Fiscal_Year == "2019" & (disag == "Total Numerator" | disag == "Total Denominator")) %>%
  select(site_id, partner, agency, indicator, disag, SNU1, PSNU, site, Fiscal_Year, Q1, Q2, Q3, Q4) %>%
  group_by(site_id, agency, partner, SNU1, PSNU, site, indicator, disag) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE)) %>%
  ungroup()


#This portion is not working: pivot_wider or pivot_wide
pmtct_stat1 <- #pmtct_stat_base %>%
  pivot_wider(pmtct_stat_base ,names_from = "disag", values_from = c(Q1, Q2, Q3, Q4))  %>%
  #pivot_wider(pmtct_stat_base ,names_from = disag, values_from = c(Q1, Q2, Q3, Q4))  
  #pivot_wider(names_from = disag, values_from = c(Q1, Q2, Q3, Q4)) 
  #pivot_wide(names_from = disag, values_from = c(Q1, Q2, Q3, Q4))  %>%

  rename(PMTCT_STAT_Q1_D = starts_with("Q1_Total D"), PMTCT_STAT_Q1_N = starts_with("Q1_Total N"), PMTCT_STAT_Q2_D = starts_with("Q2_Total D"), PMTCT_STAT_Q2_N = starts_with("Q2_Total N"), PMTCT_STAT_Q3_D = starts_with("Q3_Total D"), PMTCT_STAT_Q3_N = starts_with("Q3_Total N"), PMTCT_STAT_Q4_D = starts_with("Q4_Total D"), PMTCT_STAT_Q4_N = starts_with("Q4_Total N")) %>%
  mutate(PMTCT_STAT_Cum_D = (PMTCT_STAT_Q1_D + PMTCT_STAT_Q2_D + PMTCT_STAT_Q3_D + PMTCT_STAT_Q4_D), PMTCT_STAT_Cum_N = (PMTCT_STAT_Q1_N + PMTCT_STAT_Q2_N + PMTCT_STAT_Q3_N + PMTCT_STAT_Q4_N)) %>%
  mutate(PMTCT_STAT_Q3_Per = (PMTCT_STAT_Q3_N / PMTCT_STAT_Q3_D), PMTCT_STAT_Q3_Per = (PMTCT_STAT_Q3_N / PMTCT_STAT_Q3_D)) %>% ####### QUARTER SPECIFIC CODE
  mutate(PMTCT_STAT_Cum_Per = (PMTCT_STAT_Cum_N / PMTCT_STAT_Cum_D), PMTCT_STAT_Cum_Per = (PMTCT_STAT_Cum_N / PMTCT_STAT_Cum_D))

write_csv(pmtct_stat, "C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q2 (Post-Cleaning)/pmtct_stat_site_Q.csv", append = FALSE)

rm(pmtct_stat_base)
