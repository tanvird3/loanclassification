library(RODBC)
library(jrvFinance)
library(lubridate)
library(ycinterextra)
#library(RQuantLib)
library(plyr)
library(dplyr)
library(openxlsx)
options(stringsAsFactors = F)

options(scipen = 999)

# Advances
d1 <- read.xlsx("december.xlsx", detectDates = T)

d11 <- filter(d1, LOAN_STATUS_DESC == "UC" &
                TOTAL_BALANCE >= 50000000)

d11 <- mutate(d11, AC = paste0(BRANCH_ID, ACCOUNT_NO))

#d2<-read.xlsx("june.xlsx", detectDates = T)

d3 <- read.xlsx("Oct_ict.xlsx", detectDates = T)

d31 <- mutate(d3, AC = paste0(BRANCH_ID, ACCOUNT_NO))

d3x <-
  filter(d31, AC %in% d11$AC & !LON_STATUS_SH_NM %in% c("UC", "SMA"))

d3x$CL <-
  mapvalues(d3x$LON_STATUS_SH_NM,
            from = c("SS", "DF", "BL"),
            to = c(1, 2, 3))

d3y <- select(d3x, 35, 4, 7, 8, 11, 19, 20, 25, 26, 27, 31, 33)

d3y <- mutate(d3y, BALANCE = BALANCE / 10000000)

d3y <-
  mutate(d3y,
         SANCTION_LIMIT = SANCTION_LIMIT / 10000000,
         OVERDUE_AMT = OVERDUE_AMT / 10000000)


d3y <- arrange(d3y, desc(BALANCE))

d3y <- select(d3y, 1:3, 8, 4:7, 9, 12, 10, 11)


g <-
  as.data.frame(
    d3x %>% group_by(LOAN_TYPE, LON_STATUS_SH_NM) %>% summarise(
      NO_OF_AC = length(BALANCE),
      Outstanding = sum(BALANCE) / 10000000
    )
  ) 

g <- select(g,-CL)

if (as.vector(Sys.info()['sysname']) == "Windows") {
  Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")
}

write.xlsx(g, 'class_summary.xlsx', row.names = F)

write.xlsx(d3y,
           'class_details.xlsx',
           row.names = F,
           asTable = T)
