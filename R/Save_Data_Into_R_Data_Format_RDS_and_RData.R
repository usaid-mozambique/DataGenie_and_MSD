####################################################################################
#Saving Data into R Data Format: RDS and RDATA
####################################################################################
#Set the working directory
setwd("~/ICPI/ICPI_MER_Structured_Dataset_SITE_IM_FY17-18_Mozambique_20180515_v1_1")
#Save an object to a file
saveRDS(myfile, file = "myfile.RDS")
#Restore the object
myfile2 <- readRDS("myfile.RDS")
