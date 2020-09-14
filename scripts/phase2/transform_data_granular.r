library("readxl")
library("writexl")
library("dplyr")

# Phase 2 Excel
input <- "C:/Users/its-student/Desktop/Phase2In.xlsx"
output <- "C:/Users/its-student/Desktop/Phase2OutGranular.xlsx"

# read in sheets
data        <- read_excel(input, sheet = "Data")
demographic <- read_excel(input, sheet = "Demographic") # join by group id
sic         <- read_excel(input, sheet = "SIC")         # join by sic code


# drop reserves not related to STD
data <- data %>% select(-c(ICOS, WAIVER_IBNR, GAAP_RESV, WAIVER_RESERVE))

# create date from incurred year and month
data <- data %>% mutate(INC_DATE = as.Date(paste(INC_YEAR, INC_MONTH, "15", sep = "-")))

# drop redundant incurred date columns
data <- data %>% select(-c(INC_MONTH, INC_YEAR, INC_MONTHYEAR))

# format policy effective date
data <- data %>% mutate(POLICY_EFFECTIVE_DATE = as.Date(data$POLICY_EFFECTIVE_DATE, "%Y-%m-%d"))

# positive lives, duration, premiums, claims and reserves
data <- data %>% filter(MAX_LIVES > 0, POLICY_DURATION >= 0,
                        EST_ANNUALIZED_NET_PREM > 0, PREM > 0,
                        PAID_CLAIMS >= 0, IBNR > 0)

# remove home and national regional office
data <- data %>% filter(!grepl("^HOME", REG_OFFICE))
data <- data %>% filter(REG_OFFICE != "NATIONAL")

# city to lat-lon map
latlon <- list("ATLANTA"   = c(33.74, -84.38),  "BOSTON"   = c(42.36, -71.05),  "CHARLOTT"   = c(35.22, -80.84), 
               "CHICAGO"   = c(41.87, -87.62),  "CINCINNA" = c(39.10, -84.51),  "CLEVELAN"   = c(41.49, -81.69),
               "DALLAS"    = c(32.77, -96.79),  "DENVER"   = c(39.73, -104.99), "DETROIT"    = c(42.33, -83.04),
               "FT_LAUD"   = c(26.12, -80.13),  "HOUSTON"  = c(29.76, -95.36),  "INDIANAP"   = c(39.76, -86.15),
               "KAN_CITY"  = c(39.09, -94.57),  "LOS_ANGL" = c(34.05, -118.24), "MINNEAPO"   = c(44.97, -93.26),
               "NASHVILL"  = c(36.16, -86.78),  "NEWYORK"  = c(40.71, -74.00),  "OMAHA"      = c(41.25, -95.93),
               "ORLANDO"   = c(28.53, -81.37),  "PHILADEL" = c(39.95, -75.16),  "PHOENIX"    = c(33.44, -112.07),
               "PITTSBUR"  = c(40.44, -79.99),  "PORTLAND" = c(45.51, -122.65), "SAN_FRAN"   = c(37.77, -122.41),
               "SEATTLE"   = c(47.60, -122.33), "ST_LOUIS" = c(38.63, -90.20),  "WASHDC"     = c(38.90, -77.03))

# create latitude and longitude columns
data <- data %>% rowwise() %>% mutate(LAT = latlon[[REG_OFFICE]][1], LON = latlon[[REG_OFFICE]][2])

# left outer-join on industry code
data <- merge(x = data, y = sic, by.x = "SIC", by.y = "SIC_CODE", all.x = TRUE)

# drop sic & description
data <- data %>% select(-c(SIC, SIC_DESC))

# remove rows where the sic has no industry (e.g. SIC = 1790)
data <- data %>% filter(!is.na(INDUSTRY))

# left outer-join on demographics (age, gender & salary)
data <- merge(x = data, y = demographic, by.x = "GROUP_ID", by.y = "GROUP_ID", all.x = TRUE)

# remove groups with missing demographics
data <- data %>% filter(!is.na(AVG_AGE), !is.na(AVG_SALARY), !is.na(PCT_FEMALE))

# situs state to region map
regions <- list("east"     = c("AL","CT","DC","DE","GA","MA","MD","ME","MS","NC","NH","NJ","NY","PA","RI","SC","TN","VA","VT"),
                "central"  = c("FL","IA","IL","IN","KY","MI","MN","MO","ND","NE","OH","SD","WI","WV"),
                "west"     = c("AK","AR","AZ","CA","CO","HI","ID","KS","LA","MT","NM","NV","OK","OR","TX","UT","WA","WY"))

# create region column
data <- data %>% mutate(REGION = case_when(STATE %in% regions$east ~ "east",
                                           STATE %in% regions$central ~ "central",
                                           STATE %in% regions$west ~ "west"))

# re-order columns
data <- data %>% select("GROUP_ID", "DIST_ID", "REP_ID",
                        "REG_OFFICE", "LAT", "LON", "STATE", "REGION",
                        "INDUSTRY", "SUB_INDUSTRY",
                        "AVG_SALARY", "AVG_AGE", "PCT_FEMALE",
                        "POLICY_EFFECTIVE_DATE", "POLICY_DURATION",
                        "COVG_CODE", "TRUE_GROUP_VOL", "ACTIVE_TERMED", "LTD_INDICATOR",
                        "MAX_LIVES", "PREM", "EST_ANNUALIZED_NET_PREM",
                        "INC_DATE", "PAID_COMMISSION", "PAID_CLAIMS", "IBNR")

# write data to Excel
write_xlsx(data, path = output, col_names = TRUE)

