#install.packages("PortfolioAnalytics")
#install.packages("xts")
#install.packages("zoo")
#install.packages("ROI")
#install.packages("DEoptim")
#install.packages("ROI.plugin.glpk")
#install.packages("ROI.plugin.quadprog")

library(xlsx)
library(PortfolioAnalytics)
library(xts)
library(zoo)
library(ROI)

require(ROI.plugin.glpk)
require(ROI.plugin.quadprog)

raw_data <- read.xlsx("S:\\Treasury\\Riskmanagement\\Portfolio\\Portfolio.xlsx", sheetName="Sheet2")

#Calculate return series using logs with 1 day lag
log_return <- as.xts(c(NA, diff(log(raw_data$M2AP.Index))),
                     order.by = as.Date(as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y")))

colnames(log_return) <- c("EQ_MSCI_APAC")

log_return$EQ_MSCI_EU <- as.xts(c(NA, diff(log(raw_data$MSDEE15G.Index))),
                                order.by = as.Date(as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y")))

log_return$EQ_MSCI_EM <- as.xts(c(NA, diff(log(raw_data$M2EF.Index))),
                                order.by = as.Date(as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y")))

log_return$EQ_SPX_US <- as.xts(c(NA, diff(log(raw_data$SPTRTE.Index))),
                               order.by = as.Date(as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y")))

log_return$EQ_SME_EU <- as.xts(c(NA, diff(log(raw_data$H06732EU.Index))),
                               order.by = as.Date(as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y")))

log_return$FI_EUROPE <- as.xts(c(NA, diff(log(raw_data$NCEDE15G.Index))),
                               order.by = as.Date(as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y")))

log_return$FI_TREASURY <- as.xts(c(NA, diff(log(raw_data$LET1TREU.Index))),
                                 order.by = as.Date(as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y")))

log_return$RE_FTSE<- as.xts(c(NA, diff(log(raw_data$TRNGLE.Index))),
                            order.by = as.Date(as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y")))

log_return$CM_GOLD <- as.xts(c(NA, diff(log(raw_data$XAUEUR.Curncy))),
                             order.by = as.Date(as.POSIXct(raw_data$Dates, tryFormat = "%d.%m.%Y")))

#omit first date that became NA due to the return conversion
log_return <- na.omit(log_return)

#export return table as .csv
write.xlsx(log_return, file="S:\\Treasury\\Riskmanagement\\Portfolio\\Portfolio.xlsx", sheetName="Sheet3")

#Get a character vector of the fund names
fund.names <- colnames(log_return)

#specify a portfolio object by passing a character vector for the assets argument
pspec <- portfolio.spec(assets=fund.names)
print.default(pspec)

#Target return constraints
#pspec <- add.constraint(portfolio=pspec, type="return", return_target=0.01)

#Max mean return with ROI
maxret <- add.objective(portfolio=pspec, type="return", name="mean")
minvar <- add.objective(portfolio=pspec, type="risk", name="var")

pspec <- add.constraint(portfolio=pspec, type="group",
                              groups=list(groupEQ=c(1:5),
                                          groupFI=c(6:7),
                                          groupRE=8,
                                          groupCM=9),
                                          group_min=c(-0.05,0.0,0.0,0.0),
                                          group_max=c(0.2,1.0,0.025,0.08))

#pspec <- add.constraint(portfolio=pspec, type="return", return_target=0.005)

opt_maxret <- optimize.portfe3olio(R=log_return, portfolio=maxret, optimize_method="ROI", trace=TRUE)
opt_minvar <- optimize.portfolio(R=log_return, portfolio=minvar, optimize_method="ROI", trace=TRUE)

plot(opt_minvar, risk_col="StdDev", return.col="mean")
print(opt_minvar)

plot(log_return)
plot(log_return$EQ_SPX_US)

print(opt_maxret)

