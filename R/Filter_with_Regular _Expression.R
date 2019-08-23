#####################################Coded by: Green-EyE########################################################
#Use just a few functions from the 'stringr' package and use a little bit of the "regular expression"########### 




rm(list = ls())

#install.packages("stringr")

# Every time you open R, you will need to "open" or load your packages, using library()
library(tidyverse)
library(ICPIutilities)
library(stringr)


myfile7 <- read.delim (file = "C:/Users/cnhantumbo/Documents/DataGenie/PEPFAR-Data-Genie-SiteByIMs-2019-08-09.txt")

#Let's say you are interested in finding indicators whose names contains HRH_CURR
HRH_CURR1 <- filter(myfile7, str_detect(indicator, "HRH_CURR")) %>%
# calling 'view()' function to open the table 
view(HRH_CURR1) %>%
# calling 'count()' function at the end to make it easier to see if the result is reflecting the intention 
#of the 'filter()' command
count(indicator)


# we can use '*' (asterisk) symbol right after the '.' (dot). '*' (asterisk) is used to
#match any preceeding characters zero or more times
HRH_CURR2 <- filter(myfile7, str_detect(indicator, "HRH_CURR.*Cadre"))
view (HRH_CURR2)

# Note that there is another symbol you could use, that is '+' (plus) symbol. You can use this when you want 
#to match at least once while you can use '*' (asterisk) to match zero or multiple times
HRH_CURR3 <- filter(myfile7, str_detect(indicator, "HRH_CURR.+Cadre"))



# write and don't forget to add your working directory
write.csv(HRH_CURR1, "HRH_CURR1.csv")