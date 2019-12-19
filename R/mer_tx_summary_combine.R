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
# GENIE IMPORT AND SET UP OF BASE DATAFRAME. NOTE THAT "FILTER" AND "SELECT" CODE LINES ARE ANALYSIS SPECIFIC 

df <- read.delim("C:/Users/cnhantumbo/Documents/ICPI/MER FY2019 Q4 Pre-Cleaning/MER_Structured_Datasets_SITE_IM_FY17-20_20191115_v1_1_Mozambique.txt") %>% ####### NEEDS TO BE MANUALLY IMPORTED THROUGH MENU WITH EACH DATA GENIE UPDATE
  rename(target = targets, 
         site_id = facilityuid, 
         disag = standardizeddisaggregate, 
         agency = fundingagency, 
         site = sitename, 
         partner = mech_name, 
         Q1 = qtr1, 
         Q2 = qtr2, 
         Q3 = qtr3, 
         Q4 = qtr4)

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

ajuda <- read_excel("C:/Users/cnhantumbo/Documents/AJUDA/ajuda_ia_phase_disag.xlsx") %>% # UPDATED PATH!!!!!
  replace_na(list(ajuda = 0))


#######################
#######################
#--- OVERALL SUMMARY
#######################
#######################


#---------------------------------------------
## PREPARE TX_CURR DATASET
tx_curr_base <- df %>%
  filter(indicator == "TX_CURR" & fiscal_year == "2019", disag == "Total Numerator") %>%
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
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()

#---------------------------------------------
##AJUDA LEVEL TX_CURR PERFORMANCE
tx_curr_ajuda <- tx_curr_base %>%
  group_by(ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup() %>%
  arrange(ajuda)

#---------------------------------------------
##AGENCY LEVEL TX_CURR PERFORMANCE
tx_curr_agency <- tx_curr_base %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup() %>%
  arrange(agency)

#---------------------------------------------
##PARTNER LEVEL TX_CURR PERFORMANCE
tx_curr_partner <- tx_curr_base %>%
  group_by(agency, partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup() %>%
  arrange(agency, partner)

#---------------------------------------------
##PROVINCE LEVEL TX_CURR PERFORMANCE
tx_curr_snu1 <- tx_curr_base %>%
  group_by(agency, partner, snu1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1)


#---------------------------------------------
##DISTRICT LEVEL TX_CURR PERFORMANCE
tx_curr_psnu <- tx_curr_base %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu)


#---------------------------------------------
##SITE LEVEL TX_CURR PERFORMANCE
tx_curr_site <- tx_curr_base %>%
  group_by(site_id, agency, partner, snu1, psnu, site, ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu, site)

rm(tx_curr_base)


#---------------------------------------------
## PREPARE TX_NEW DATASET
tx_new_base <- df %>%
  filter(indicator == "TX_NEW" & fiscal_year == "2019", disag == "Total Numerator") %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, target) %>%
  left_join(ajuda, by = "site_id") %>%
  replace_na(list(ajuda = 0)
  )


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


wb_overall <- createWorkbook()
addWorksheet(wb_overall, "tx_curr_ou",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_ou", tx_curr_ou, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_ou", style = pct, cols=c(7:9), rows = 2:(nrow(tx_curr_ou)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_ajuda",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_ajuda", tx_curr_ajuda, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_ajuda", style = pct, cols=c(7:9), rows = 2:(nrow(tx_curr_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_agency",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_agency", tx_curr_agency, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_agency", style = pct, cols=c(7:9), rows = 2:(nrow(tx_curr_agency)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_partner",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_partner", tx_curr_partner, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_partner", style = pct, cols=c(8:10), rows = 2:(nrow(tx_curr_partner)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_snu1",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_snu1", tx_curr_snu1, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_snu1", style = pct, cols=c(9:11), rows = 2:(nrow(tx_curr_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_overall, "tx_curr_psnu",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_psnu", tx_curr_psnu, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_psnu", style = pct, cols=c(10:12), rows = 2:(nrow(tx_curr_psnu)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_curr_site",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_curr_site", tx_curr_site, tableStyle = "TableStyleLight2")
addStyle(wb_overall, "tx_curr_site", style = pct, cols=c(13:15), rows = 2:(nrow(tx_curr_site)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_ou",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_ou", tx_new_ou, tableStyle = "TableStyleLight3")
addStyle(wb_overall, "tx_new_ou", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_ou)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_ajuda",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_ajuda", tx_new_ajuda, tableStyle = "TableStyleLight3")
addStyle(wb_overall, "tx_new_ajuda", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_agency",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_agency", tx_new_agency, tableStyle = "TableStyleLight3")
addStyle(wb_overall, "tx_new_agency", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_agency)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_partner",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_partner", tx_new_partner, tableStyle = "TableStyleLight3")
addStyle(wb_overall, "tx_new_partner", style = pct, cols=c(8:12), rows = 2:(nrow(tx_new_partner)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_snu1",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_snu1", tx_new_snu1, tableStyle = "TableStyleLight3")
addStyle(wb_overall, "tx_new_snu1", style = pct, cols=c(9:13), rows = 2:(nrow(tx_new_snu1)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_psnu",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_psnu", tx_new_psnu, tableStyle = "TableStyleLight3")
addStyle(wb_overall, "tx_new_psnu", style = pct, cols=c(10:14), rows = 2:(nrow(tx_new_psnu)+1), gridExpand=TRUE)

addWorksheet(wb_overall, "tx_new_site",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_overall, "tx_new_site", tx_new_site, tableStyle = "TableStyleLight3")
addStyle(wb_overall, "tx_new_site", style = pct, cols=c(13:17), rows = 2:(nrow(tx_new_site)+1), gridExpand=TRUE)




#######################
#######################
#--- PEDIATRIC SUMMARY
#######################
#######################


#---------------------------------------------
## PREPARE TX_CURR DATASET
tx_curr_base <- df %>%
  filter(indicator == "TX_CURR" & fiscal_year == "2019", trendscoarse == "<15") %>%
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
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()

#---------------------------------------------
##AJUDA LEVEL TX_CURR PERFORMANCE
tx_curr_ajuda <- tx_curr_base %>%
  group_by(ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup()

#---------------------------------------------
##AGENCY LEVEL TX_CURR PERFORMANCE
tx_curr_agency <- tx_curr_base %>%
  group_by(agency) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup %>%
  arrange(agency)

#---------------------------------------------
##PARTNER LEVEL TX_CURR PERFORMANCE
tx_curr_partner <- tx_curr_base %>%
  group_by(agency, partner) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup() %>%
  arrange(agency, partner)


#---------------------------------------------
##PROVINCE LEVEL TX_CURR PERFORMANCE
tx_curr_snu1 <- tx_curr_base %>%
  group_by(agency, partner, snu1) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1)


#---------------------------------------------
##DISTRICT LEVEL TX_CURR PERFORMANCE
tx_curr_psnu <- tx_curr_base %>%
  group_by(agency, partner, snu1, psnu) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu)


#---------------------------------------------
##SITE LEVEL TX_CURR PERFORMANCE
tx_curr_site <- tx_curr_base %>%
  group_by(site_id, agency, partner, snu1, psnu, site, ajuda) %>%
  summarize(Q1 = sum(Q1, na.rm = TRUE), Q2 = sum(Q2, na.rm = TRUE), Q3 = sum(Q3, na.rm = TRUE), Q4 = sum(Q4, na.rm = TRUE), target = sum(target, na.rm = TRUE)) %>%
  mutate(Q2_netnew = Q2/Q1, Q3_netnew = Q3/Q2, Q4_netnew = Q4/Q3) %>%
  ungroup() %>%
  arrange(agency, partner, snu1, psnu, site)

rm(tx_curr_base)


#---------------------------------------------
## PREPARE TX_NEW DATASET
tx_new_base <- df %>%
  filter(indicator == "TX_NEW" & fiscal_year == "2019", trendscoarse == "<15") %>%
  select(site_id, partner, agency, indicator, disag, snu1, psnu, site, fiscal_year, Q1, Q2, Q3, Q4, target) %>%
  left_join(ajuda, by = "site_id") %>%
  replace_na(list(ajuda = 0)
  )


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


wb_pediatrico <- createWorkbook()
addWorksheet(wb_pediatrico, "tx_curr_ou",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_ou", tx_curr_ou, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_ou", style = pct, cols=c(7:9), rows = 2:(nrow(tx_curr_ou)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_curr_ajuda",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_ajuda", tx_curr_ajuda, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_ajuda", style = pct, cols=c(7:9), rows = 2:(nrow(tx_curr_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_curr_agency",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_agency", tx_curr_agency, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_agency", style = pct, cols=c(7:9), rows = 2:(nrow(tx_curr_agency)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_curr_partner",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_partner", tx_curr_partner, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_partner", style = pct, cols=c(8:10), rows = 2:(nrow(tx_curr_partner)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_curr_snu1",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_snu1", tx_curr_snu1, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_snu1", style = pct, cols=c(9:11), rows = 2:(nrow(tx_curr_snu1)+1), gridExpand=TRUE)


addWorksheet(wb_pediatrico, "tx_curr_psnu",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_psnu", tx_curr_psnu, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_psnu", style = pct, cols=c(10:12), rows = 2:(nrow(tx_curr_psnu)+1), gridExpand=TRUE)


addWorksheet(wb_pediatrico, "tx_curr_site",
             tabColour = "#99CCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_curr_site", tx_curr_site, tableStyle = "TableStyleLight2")
addStyle(wb_pediatrico, "tx_curr_site", style = pct, cols=c(13:15), rows = 2:(nrow(tx_curr_site)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_ou",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_ou", tx_new_ou, tableStyle = "TableStyleLight3")
addStyle(wb_pediatrico, "tx_new_ou", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_ou)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_ajuda",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_ajuda", tx_new_ajuda, tableStyle = "TableStyleLight3")
addStyle(wb_pediatrico, "tx_new_ajuda", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_ajuda)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_agency",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_agency", tx_new_agency, tableStyle = "TableStyleLight3")
addStyle(wb_pediatrico, "tx_new_agency", style = pct, cols=c(7:11), rows = 2:(nrow(tx_new_agency)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_partner",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_partner", tx_new_partner, tableStyle = "TableStyleLight3")
addStyle(wb_pediatrico, "tx_new_partner", style = pct, cols=c(8:12), rows = 2:(nrow(tx_new_partner)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_snu1",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_snu1", tx_new_snu1, tableStyle = "TableStyleLight3")
addStyle(wb_pediatrico, "tx_new_snu1", style = pct, cols=c(9:13), rows = 2:(nrow(tx_new_snu1)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_psnu",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_psnu", tx_new_psnu, tableStyle = "TableStyleLight3")
addStyle(wb_pediatrico, "tx_new_psnu", style = pct, cols=c(10:14), rows = 2:(nrow(tx_new_psnu)+1), gridExpand=TRUE)

addWorksheet(wb_pediatrico, "tx_new_site",
             tabColour = "#FFCCFF",
             gridLines = FALSE)
writeDataTable(wb_pediatrico, "tx_new_site", tx_new_site, tableStyle = "TableStyleLight3")
addStyle(wb_pediatrico, "tx_new_site", style = pct, cols=c(13:17), rows = 2:(nrow(tx_new_site)+1), gridExpand=TRUE)


#---------------------------------------------
##PRINT WORKBOOKS
#saveWorkbook(wb_overall, file = "C:/Users/jlara/Documents/R/r_projects/mer/output/ct/test/tx_summary.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!
#saveWorkbook(wb_pediatrico, file = "C:/Users/jlara/Documents/R/r_projects/mer/output/ct/test/tx_summary_ped.xlsx", overwrite = TRUE) # UPDATED PATH!!!!!





