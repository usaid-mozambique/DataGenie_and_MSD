#-----------------------------------------------------------------------------------
##  LOAD PACKAGES
library(readr)
rm(list = ls())

# This script Extract USAID targets only from MSD dataset


#-----------------------------------------------------------------------------------
# IMPORT ECHO SUBMISSION MER_Structured_Datasets_Site_IM

df <- read.delim("C:/Users/cnhantumbo/Documents/COP21/MER_Structured_Datasets_Site_IM_FY19-21_20210514_v1_1_Mozambique.txt")


#This approach will set the data frame’s internal pointer to that 
#single column to NULL, releasing the space and will remove the required column from the R data frame.
df$qtr1 <- NULL 
df$qtr2 <- NULL 
df$qtr3 <- NULL 
df$qtr4 <- NULL 
df$cumulative <- NULL 
df$source_name <- NULL 

#Delete other fundingagencies from df data - No need to manage a new data.frame.
df <- subset (df,df$fundingagency == "USAID" )
print(df)


write.table (df, "C:/Users/cnhantumbo/Documents/COP21/MER_Structured_Datasets_Site_IM_FY19-21_20210514_v1_1_Mozambique_df.txt"
            , sep = "\t",  row.names = FALSE )
#df_out <- read.delim("C:/Users/cnhantumbo/Documents/COP21/MER_Structured_Datasets_Site_IM_FY19-21_20210514_v1_1_Mozambique_df.txt")

                                  
