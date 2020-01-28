rm(list = ls())
library(readr)
library(dplyr)


#*************These functions override the set functions provided in base to make them generic so that efficient versions******** 
#*******************for data frames and other tables can be provided. The default methods call the base versions.************** 
#***************************Beware that intersect(), union() and setdiff() remove duplicates.*************************************

ds1 <- read.delim("C:/Users/cnhantumbo/Documents/DataGenie/PEPFAR-Data-Genie-SiteByIMs-2020-01-27.txt")
ds2 <- read.delim("C:/Users/cnhantumbo/Documents/DataGenie/PEPFAR-Data-Genie-SiteByIMs-2020-01-28.txt")

#set difference, duplicates removed 
ds3 <-setdiff(ds1, ds2)

#intersection, duplicates removed
ds4 <- intersect(ds1, ds2)

# union, duplicates removed 
ds5 <- union(ds1, ds2)

#union all does not remove duplicates
ds6 <- union_all(ds1, ds2)