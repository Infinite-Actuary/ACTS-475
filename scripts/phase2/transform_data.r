library("readxl")
library("writexl")
library("dplyr")

# Phase 2 Excel
input <- "C:/Users/its-student/Desktop/Phase2In.xlsx"
output <- "C:/Users/its-student/Desktop/Phase2Out.xlsx"

# read in sheets
data        <- read_excel(input, sheet = "Data")
commission  <- read_excel(input, sheet = "Commission")  # join by group id & policy duration
demographic <- read_excel(input, sheet = "Demographic") # join by group id
expense     <- read_excel(input, sheet = "Expense")     # join by annualized net premium
rtn         <- read_excel(input, sheet = "RTN")         # join by max lives, voluntary & policy duration
sic         <- read_excel(input, sheet = "SIC")         # join by sic code
tax         <- read_excel(input, sheet = "Tax")         # join by state

# drop reserves not related to STD
data <- data %>% select(-c(ICOS, WAIVER_IBNR, GAAP_RESV, WAIVER_RESERVE))

# create date from incurred year and month
#data <- data %>% mutate(INC_DATE = as.Date(paste(INC_YEAR, INC_MONTH, "01", sep = "-")))

# drop redundant incurred date columns, keep INC_YEAR for grouping
data <- data %>% select(-c(INC_MONTH, INC_MONTHYEAR))

# drop broker and LFG sales rep hash key
#data <- data %>% select(-c(DIST_ID, REP_ID))

# format policy effective date
data <- data %>% mutate(POLICY_EFFECTIVE_DATE = as.Date(data$POLICY_EFFECTIVE_DATE, "%Y-%m-%d"))

# format incurred year as numeric
data <- data %>% mutate(INC_YEAR = as.numeric(INC_YEAR))

# group claims on annual basis
data <- data %>% group_by(GROUP_ID, DIST_ID, REP_ID,
                  COVG_CODE, TRUE_GROUP_VOL, POLICY_EFFECTIVE_DATE, 
                  REG_OFFICE, SIC, STATE, 
                  MAX_LIVES, ACTIVE_TERMED, LTD_INDICATOR, INC_YEAR, POLICY_DURATION) %>% 
  summarise(PREM = sum(PREM), EST_ANNUALIZED_NET_PREM = mean(EST_ANNUALIZED_NET_PREM),
            PAID_COMMISSION = sum(PAID_COMMISSION), PAID_CLAIMS = sum(PAID_CLAIMS), IBNR = sum(IBNR))

# positive lives, duration, premiums, claims and reserves
data <- data %>% filter(MAX_LIVES > 0, POLICY_DURATION >= 0,
                        EST_ANNUALIZED_NET_PREM > 0, PREM > 0,
                        PAID_CLAIMS >= 0, IBNR > 0)

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

# append percent commission
data <- merge(x = data, 
              y = commission[, c("GROUP_ID", "POLICY_DURATION", "PERCENT_COMMISSION")], 
              by = c("GROUP_ID" = "GROUP_ID", "POLICY_DURATION" = "POLICY_DURATION"), 
              all.x = TRUE)

# remove rows where percent comission is NA or negative due to chargebacks
data <- data %>% filter(0 <= PERCENT_COMMISSION & PERCENT_COMMISSION <= 1)

# append state premium tax
data <- merge(x = data, y = tax, by.x = "STATE", by.y = "STATE", all.x = TRUE)

# remove rows with unmapped state tax (e.g. 91, FO)
data <- data %>% filter(!is.na(PREMIUM_TAX))

# state to lat-lon map
latlon <- list("AL" = c(33,-87),   "AK" = c(61,-152),  "AZ" = c(34,-111),  "AR" = c(35,-92),
               "CA" = c(36,-120),  "CO" = c(39,-105),  "CT" = c(42,-73),   "DE" = c(39,-76),
               "DC" = c(39,-77),   "FL" = c(28,-82),   "GA" = c(33,-84),   "HI" = c(21,-157),
               "ID" = c(44,-114),  "IL" = c(40,-89),   "IN" = c(40,-86),   "IA" = c(42,-93),
               "KS" = c(39,-97),   "KY" = c(38,-85),   "LA" = c(31,-92),   "ME" = c(45,-69),
               "MD" = c(39,-77),   "MA" = c(42,-72),   "MI" = c(43,-85),   "MN" = c(46,-94),
               "MS" = c(33,-90),   "MO" = c(38,-92),   "MT" = c(47,-110),  "NE" = c(41,-98),
               "NV" = c(38,-117),  "NH" = c(43,-72),   "NJ" = c(40,-75),   "NM" = c(35,-106),
               "NY" = c(42,-75),   "NC" = c(36,-80),   "ND" = c(48,-100),  "OH" = c(40,-83),
               "OK" = c(36,-97),   "OR" = c(45,-122),  "PA" = c(41,-77),   "RI" = c(42,-72),
               "SC" = c(34,-81),   "SD" = c(44,-99),   "TN" = c(36,-87),   "TX" = c(31,-98),
               "UT" = c(40,-112),  "VT" = c(44,-73),   "VA" = c(38,-78),   "WA" = c(47,-121),
               "WV" = c(38,-81),   "WI" = c(44,-90),   "WY" = c(43,-107))

# create latitude and longitude columns
data <- data %>% rowwise() %>% mutate(LAT = latlon[[STATE]][1], LON = latlon[[STATE]][2])

# situs state to region map
regions <- list("east"     = c("AL","CT","DC","DE","GA","MA","MD","ME","MS","NC","NH","NJ","NY","PA","RI","SC","TN","VA","VT"),
                "central"  = c("FL","IA","IL","IN","KY","MI","MN","MO","ND","NE","OH","SD","WI","WV"),
                "west"     = c("AK","AR","AZ","CA","CO","HI","ID","KS","LA","MT","NM","NV","OK","OR","TX","UT","WA","WY"))

# create region column
data <- data %>% mutate(REGION = case_when(STATE %in% regions$east ~ "east",
                                           STATE %in% regions$central ~ "central",
                                           STATE %in% regions$west ~ "west"))

# get internal expense from estimated annual net premium
get_internal_expense <- function(premium) {
  sapply(premium, function(x) {
    for (i in 1:nrow(expense)) {
      if (expense$EST_ANN_NET_PREM_MIN[i] <= x & x < expense$EST_ANN_NET_PREM_MAX[i]) {
        return(expense$INTERNAL_EXPENSE[i])
      }
    }
  })
}

# create new internal expenses column
data <- data %>% mutate(INTERNAL_EXPENSES = get_internal_expense(EST_ANNUALIZED_NET_PREM))

# assume pepm rate = 0.5 for all est annualized premiums
data <- data %>% mutate(PERCENT_PEPM = MAX_LIVES * 0.5 / PREM)

# PEPM in [0,1]
data <- data %>% filter(0 <= PERCENT_PEPM & PERCENT_PEPM <= 1)

# create tolerable loss ratio
data <- data %>% mutate(TLR = 1 - (PERCENT_COMMISSION + PREMIUM_TAX + PERCENT_PEPM + INTERNAL_EXPENSES))

# TLR in [0,1]
data <- data %>% filter(0 <= TLR & TLR <= 1)

# rtn lookup from policy lives, duration & voluntary indicator
data <- data %>% 
  mutate(RTN = case_when(MAX_LIVES  < 100  & TRUE_GROUP_VOL == "T" & POLICY_DURATION <  2 ~ 0.8833,
                         MAX_LIVES  < 100  & TRUE_GROUP_VOL == "T" & POLICY_DURATION <  4 ~ 0.9700,
                         MAX_LIVES  < 100  & TRUE_GROUP_VOL == "T" & POLICY_DURATION >= 4 ~ 1.0200,
                         MAX_LIVES  < 100  & TRUE_GROUP_VOL == "V" & POLICY_DURATION <  2 ~ 0.9733,
                         MAX_LIVES  < 100  & TRUE_GROUP_VOL == "V" & POLICY_DURATION <  4 ~ 1.0185,
                         MAX_LIVES  < 100  & TRUE_GROUP_VOL == "V" & POLICY_DURATION >= 4 ~ 1.0670,
                         MAX_LIVES  < 1000 & TRUE_GROUP_VOL == "T" & POLICY_DURATION <  2 ~ 0.8741,
                         MAX_LIVES  < 1000 & TRUE_GROUP_VOL == "T" & POLICY_DURATION <  4 ~ 0.9514,
                         MAX_LIVES  < 1000 & TRUE_GROUP_VOL == "T" & POLICY_DURATION >= 4 ~ 1.0208,
                         MAX_LIVES  < 1000 & TRUE_GROUP_VOL == "V" & POLICY_DURATION <  2 ~ 1.0104,
                         MAX_LIVES  < 1000 & TRUE_GROUP_VOL == "V" & POLICY_DURATION <  4 ~ 1.0347,
                         MAX_LIVES  < 1000 & TRUE_GROUP_VOL == "V" & POLICY_DURATION >= 4 ~ 1.0670,
                         MAX_LIVES >= 1000 & TRUE_GROUP_VOL == "T" & POLICY_DURATION <  2 ~ 0.9800,
                         MAX_LIVES >= 1000 & TRUE_GROUP_VOL == "T" & POLICY_DURATION <  4 ~ 1.0000,
                         MAX_LIVES >= 1000 & TRUE_GROUP_VOL == "T" & POLICY_DURATION >= 4 ~ 1.0200,
                         MAX_LIVES >= 1000 & TRUE_GROUP_VOL == "V" & POLICY_DURATION <  2 ~ 0.9800,
                         MAX_LIVES >= 1000 & TRUE_GROUP_VOL == "V" & POLICY_DURATION <  4 ~ 1.0000,
                         MAX_LIVES >= 1000 & TRUE_GROUP_VOL == "V" & POLICY_DURATION >= 4 ~ 1.0200))

# calculate needed premium
data <- data %>% mutate(NEEDED_PREMIUM = PREM / RTN)

# calculate expected claims
data <- data %>% mutate(EXPECTED_CLAIMS = NEEDED_PREMIUM * TLR)

# calcualte actual claims
data <- data %>% mutate(ACTUAL_CLAIMS = PAID_CLAIMS + IBNR)

# calculate actual to expected ratio
data <- data %>% mutate(ACTUAL_TO_EXPECTED = ACTUAL_CLAIMS / EXPECTED_CLAIMS)

# re-order columns
data <- data %>% select("GROUP_ID", "DIST_ID", "REP_ID",
                        "REG_OFFICE", "STATE", "LAT", "LON", "REGION", "INDUSTRY", "SUB_INDUSTRY",
                        "AVG_SALARY", "AVG_AGE", "PCT_FEMALE",
                        "POLICY_EFFECTIVE_DATE", "POLICY_DURATION",
                        "COVG_CODE", "TRUE_GROUP_VOL", "ACTIVE_TERMED", "LTD_INDICATOR",
                        "MAX_LIVES", "PREM", "NEEDED_PREMIUM", "RTN", "EST_ANNUALIZED_NET_PREM", "INC_YEAR",
                        "PAID_COMMISSION", "PAID_CLAIMS", "IBNR", "ACTUAL_CLAIMS", "EXPECTED_CLAIMS", "ACTUAL_TO_EXPECTED",
                        "PERCENT_COMMISSION", "PREMIUM_TAX", "INTERNAL_EXPENSES", "PERCENT_PEPM", "TLR")

# write data to Excel
write_xlsx(data, path = output, col_names = TRUE)

