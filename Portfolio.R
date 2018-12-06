#install.packages("PortfolioAnalytics")
#install.packages("xts")
#install.packages("zoo")

library(xlsx)
library(PortfolioAnalytics)
library(xts)
library(zoo)

raw_data <- read.xlsx("S:\\Treasury\\Riskmanagement\\Portfolio\\Portfolio.xlsx", sheetName="Sheet2", rowNames=TRUE)

#Calculate return series using logs with 1 day lag
log_return <- as.xts(c(NA, diff(log(raw_data$M2AP.Index))),
                 order.by = as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y"))

colnames(log_return) <- c("")
log_return$NDX <- as.xts(c(NA, diff(log(raw_data$MSDEE15G.Index))),
                     order.by = as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y"))

log_return$NKY <- as.xts(c(NA, diff(log(raw_data$SPTRTE.Index))),
                     order.by = as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y"))

log_return$FTSE <- as.xts(c(NA, diff(log(raw_data$NCEDE15G.Index))),
                      order.by = as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y"))

log_return$DAX <- as.xts(c(NA, diff(log(raw_data$H06732EU.Index))),
                     order.by = as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y"))

log_return$DAX <- as.xts(c(NA, diff(log(raw_data$TRNGLE.Index))),
                         order.by = as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y"))

log_return$DAX <- as.xts(c(NA, diff(log(raw_data$XAUEUR.Curncy))),
                         order.by = as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y"))

log_return$DAX <- as.xts(c(NA, diff(log(raw_data$M2EF.Index))),
                         order.by = as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y"))

log_return$DAX <- as.xts(c(NA, diff(log(raw_data$LET1TREU.Index))),
                         order.by = as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y"))


#omit first date that became NA due to the return conversion
log_return <- na.omit(log_return)  

#export return table as .csv
write.zoo(log_return, file="S:\\Treasury\\Riskmanagement\\Portfolio\\Portfolio.xlsx", sheetName="Sheet2", rowNames=TRUE)