library("readxl")
library("writexl")
library("dplyr")
library("FactoMineR")
library("factoextra")
library("viridis")

# phase 2 excel
input <- "C:/Users/its-student/Desktop/Phase2In.xlsx"

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

# 2018 is the only full year of data
data <- data %>% filter(INC_YEAR == "2018")

# drop incurred year after filtering
data <- data %>% select(-c(INC_MONTH, INC_YEAR, INC_MONTHYEAR))

# positive lives and policy duration
data <- data %>% filter(MAX_LIVES > 0, POLICY_DURATION >= 0)

# track number of claims paid (chargebacks decrease count)
data <- data %>% mutate(CLAIM_COUNT = case_when(PAID_CLAIMS  > 0 ~  1,
                                                PAID_CLAIMS == 0 ~  0,
                                                PAID_CLAIMS  < 0 ~ -1))

# group claims on annual basis
data <- data %>% group_by(GROUP_ID, DIST_ID, REP_ID,
                          COVG_CODE, TRUE_GROUP_VOL,
                          POLICY_EFFECTIVE_DATE, 
                          REG_OFFICE, SIC, STATE, 
                          MAX_LIVES, ACTIVE_TERMED, LTD_INDICATOR, POLICY_DURATION) %>% 
  summarise(PREM = sum(PREM), EST_ANNUALIZED_NET_PREM = mean(EST_ANNUALIZED_NET_PREM),
            PAID_COMMISSION = sum(PAID_COMMISSION), PAID_CLAIMS = sum(PAID_CLAIMS), CLAIM_COUNT = sum(CLAIM_COUNT), IBNR = sum(IBNR))

# positive premiums, paid claims and reserves
data <- data %>% filter(EST_ANNUALIZED_NET_PREM >= 0, PREM >= 0, PAID_CLAIMS >= 0, IBNR >= 0)

# left outer-join on industry code
data <- merge(x = data, y = sic, by.x = "SIC", by.y = "SIC_CODE", all.x = TRUE)

# normalize industry title by removing spaces and commas
data <- data %>% mutate(INDUSTRY = stringr::str_replace_all(INDUSTRY, c("," = "", " " = "")))

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


# famd = pca (numerics) + mca (factors)
data <- data %>% select("REGION", "INDUSTRY",
  "COVG_CODE", "LTD_INDICATOR", 
  "AVG_SALARY", "AVG_AGE", "PCT_FEMALE",
  "TRUE_GROUP_VOL", "LTD_INDICATOR", "ACTIVE_TERMED",
  "MAX_LIVES", "POLICY_DURATION", "PREM", "EST_ANNUALIZED_NET_PREM",
  "RTN",
  "PAID_COMMISSION", "PAID_CLAIMS", "CLAIM_COUNT", "IBNR",
  "PERCENT_COMMISSION", "PREMIUM_TAX", "INTERNAL_EXPENSES", "PERCENT_PEPM")

# data correlation
data.cor <- cor(select_if(data, is.numeric))

# heat map of correlation
heatmap(x = data.cor, symm = TRUE, 
        margins = c(15, 10), col = viridis(256))

# use FactoMineR for PCA
data.famd <- FAMD(data, graph = FALSE)

# data.famd$quali.var
summary(data.famd)

# scree plot
fviz_screeplot(data.famd, addlabels = TRUE)

# contribution to Dim (red dashed line = average value)
fviz_contrib(data.famd, "var", axes = 1)
fviz_contrib(data.famd, "var", axes = 2)

# quantitative plot
fviz_famd_var(data.famd, "quanti.var", 
              col.var = "cos2", gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"),
              repel = TRUE, ggtheme = theme_minimal())

# qualitative plot
fviz_famd_var(data.famd, "quali.var", 
              col.var = "cos2", gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"), 
              repel = TRUE, ggtheme = theme_minimal())

# mca plot of individuals for true group indicator
fviz_mca_ind(data.famd, label = "none", 
             habillage = c("TRUE_GROUP_VOL"),
             addEllipses = TRUE,
             pointsize = 0,
             ggtheme = theme_minimal())