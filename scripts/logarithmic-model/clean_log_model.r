# install
install.packages("readxl")
install.packages("writexl")
install.packages("dplyr")
install.packages("ggplot2")

# load packages
library("readxl")
library("writexl")
library("dplyr")
library("ggplot2")

# ======================================== S C R U B B I N G ========================================

# set file names
input <- "/home/carl/Desktop/RtnData.xlsx"
output <- "/home/carl/Desktop/RtnDataClean.xlsx"

# read in the data
data <- read_excel(input, sheet = "Data")
sic <- read_excel(input, sheet = "SIC")

# left outer-join on industry code
rtnData <- merge(x = data, y = sic, by.x = "INDUSTRY_CODE", by.y = "SIC_CODE", all.x = TRUE)

# drop VAL_DATE column
rtnData <- rtnData %>% select(-c("VAL_DATE"))

# remove NBOC and Amendment rows
rtnData <- rtnData %>% filter(EVENT_TYPE != "NBOC")
rtnData <- rtnData %>% filter(EVENT_TYPE != "Amendment")

# use only positive premiums
rtnData <- rtnData %>% filter(NEEDED_PREMIUM > 0 & FINAL_QUOTED_PREMIUM > 0)

# use plans with positive enrolled
rtnData <- rtnData %>% filter(ENROLLED_LIVES > 0)

# standardize renewals
rtnData <- rtnData %>% mutate(RENEWAL_INSTANCE = case_when(EVENT_TYPE == "New Business" ~ "0",
                                                           RENEWAL_INSTANCE == "1st Renewal" ~ "1",
                                                           RENEWAL_INSTANCE == "2nd Renewal" ~ "2",
                                                           RENEWAL_INSTANCE == "2nd+" ~ "3",
                                                           RENEWAL_INSTANCE == "3rd Renewal +" ~ "3",
                                                           TRUE ~ as.character(RENEWAL_INSTANCE)))
rtnData <- rtnData %>% filter(!is.na(RENEWAL_INSTANCE))

# convert renewals to numeric
rtnData$RENEWAL_INSTANCE <- as.numeric(rtnData$RENEWAL_INSTANCE)

# format date timestamps
rtnData$RAE_EFF_DATE <- as.Date(rtnData$RAE_EFF_DATE, "%y-%m-%d")
rtnData$COV_EFF_DATE <- as.Date(rtnData$COV_EFF_DATE, "%y-%m-%d")

# set date range >= 2017
rtnData <- rtnData %>% filter(RAE_EFF_DATE >= as.Date("2017-01-01"))
rtnData <- rtnData %>% filter(COV_EFF_DATE >= as.Date("2017-01-01"))

# clean up sales office names
rtnData <- rtnData %>% mutate(SALES_OFFICE = tolower(SALES_OFFICE))
rtnData <- rtnData %>% mutate(SALES_OFFICE = case_when(SALES_OFFICE == "central philadelphia" ~ "philadelphia",
                                                       SALES_OFFICE == "chicago/indy/milwaukee" ~ "chicago",
                                                       SALES_OFFICE == "home" ~ "ft wayne",
                                                       SALES_OFFICE == "miami/orlando/tampa" ~ "miami",
                                                       SALES_OFFICE == "washington dc" ~ "washington d.c.",
                                                       TRUE ~ as.character(SALES_OFFICE)))

# map sales office to sales region
rtnData <- rtnData %>% mutate(UW_REGION = tolower(UW_REGION))
regions <- list("home" = c("ft wayne"),
                "central" = c("chicago","cincinnati","cleveland","detroit","ft lauderdale","indianapolis","miami","milwaukee","minneapolis","omaha","orlando","pittsburgh","st. louis","tampa"),
                "east" = c("atlanta","boston","charlotte","long island","nashville","new jersey","new york","parsnippany","philadelphia","portland, me","rochester","washington d.c.","white plains"),
                "west" = c("dallas","denver","houston","kansas city","los angeles","orange county","phoenix","portland, or","sacramento","san diego","san francisco","seattle"))

rtnData <- rtnData %>% mutate(UW_REGION = case_when(SALES_OFFICE %in% regions$home ~ "home",
                                                    SALES_OFFICE %in% regions$central ~ "central",
                                                    SALES_OFFICE %in% regions$east ~ "east",
                                                    SALES_OFFICE %in% regions$west ~ "west",
                                                    TRUE ~ as.character(UW_REGION)))

# define coverage types
coverageType <- list("Life" = c("Basic Life","Dependent Life","Life","Life EE Paid","Life ER Paid","Optional Child Life","Optional Life","Optional Spouse Life","Voluntary Child Life","Voluntary Life","Voluntary Life - Unismoke","Voluntary Life - Unismoke ALT Plan","Voluntary Spouse Life"),
                      "STD" = c("NY DBL","State Disability - DBL","State Disability - TDB","STD","STD ASO","STD ATP","STD Core/Buy Up","STD EE Paid","STD ER Paid","STD Spec Worksite","STD True Zero","SW STD Alternate","Voluntary STD"),
                      "LTD" = c("LTD","LTD Core/Buy Up","LTD EE Paid","LTD ER Paid","LTD Spec Worksite","LTD True Zero","Optional LTD","Voluntary LTD"),
                      "AD&D" = c("AD&D","AD&D EE Paid","AD&D ER Paid","Basic AD&D","Optional AD&D","Voluntary AD&D","Voluntary Spouse AD&D"),
                      "Dental" = c("Dental","Dental HMO","Self-Funded Dental","Voluntary Dental","Voluntary Dental HMO"))

# aggregate coverage names into coverage types
rtnData <- rtnData %>% mutate(COVERAGE_TYPE = case_when(COVERAGE_NAME %in% coverageType$Life ~ "Life",
                                                        COVERAGE_NAME %in% coverageType$STD ~ "STD",
                                                        COVERAGE_NAME %in% coverageType$LTD ~ "LTD",
                                                        COVERAGE_NAME %in% coverageType$`AD&D` ~ "AD&D",
                                                        COVERAGE_NAME %in% coverageType$Dental ~ "Dental",
                                                        TRUE ~ "Misc"))

# remove misc coverage types (i.e. Implementation Credit, Vision, Voluntary Vision)
rtnData <- rtnData %>% filter(COVERAGE_TYPE != "Misc")

# classify groups into size by eligible lives
rtnData <- rtnData %>% mutate(GROUP_SIZE = case_when(ELIGIBLE_LIVES < 100 ~ "S",
                                                     ELIGIBLE_LIVES >= 100 & ELIGIBLE_LIVES < 1000 ~ "M",
                                                     ELIGIBLE_LIVES >= 1000 & ELIGIBLE_LIVES < 10000 ~ "L",
                                                     ELIGIBLE_LIVES >= 10000 ~ "XL",
                                                     TRUE ~ "NA"))

# add rtn column
rtnData <- rtnData %>% mutate(RTN = FINAL_QUOTED_PREMIUM / NEEDED_PREMIUM)

# write rtn data to Excel
#write_xlsx(rtnData, path = output, col_names = TRUE)

# ======================================== M O D E L I N G ========================================

# shift renewal instance for logarithmic model because log(0) = -inf
rtnData <- rtnData %>% mutate(RENEWAL_INSTANCE = RENEWAL_INSTANCE + 1.0)

# create modeling factors based on coverage type, region & industry
model.factors <- as.data.table(rtnData %>% group_by(COVERAGE_TYPE, UW_REGION, INDUSTRY) %>% summarise())

# initialize model list
models <- list()

# build out models for all factors (3 regions) x (5 coverage types) x (10 industries) = 150 combinations
for (i in 1:nrow(model.factors)) {
  # pull in relevant data for each factor
  model.data <- rtnData %>% filter(COVERAGE_TYPE == model.factors[i, "COVERAGE_TYPE"][[1]],
                                   UW_REGION == model.factors[i, "UW_REGION"][[1]],
                                   INDUSTRY == model.factors[i, "INDUSTRY"][[1]]) %>% select(RENEWAL_INSTANCE, RTN, NEEDED_PREMIUM)
  # save regression to model list
  models[[i]] <- lm(RTN ~ log(RENEWAL_INSTANCE), data = model.data, weights = NEEDED_PREMIUM)
}

# example using short-term disability in the service industry from the east region
std.east.service <- rtnData %>% filter(COVERAGE_TYPE == "STD", UW_REGION == "east", INDUSTRY == "Services") %>%
  select(RENEWAL_INSTANCE, RTN, NEEDED_PREMIUM)

# create a logarithmic model weighted by needed premium
std.east.service.fit <- lm(RTN ~ log(RENEWAL_INSTANCE), data = std.east.service, weights = NEEDED_PREMIUM)

# generate sampling points for model
x <- data.frame(RENEWAL_INSTANCE = seq(from = range(std.east.service$RENEWAL_INSTANCE)[1],
                                       to = range(std.east.service$RENEWAL_INSTANCE)[2], 
                                       length.out = 100))

# get prediction errors
errors <- predict(std.east.service.fit, newdata = x, se.fit = TRUE)

# create alpha confidence interval
alpha <- 0.99
x$lci <- errors$fit - qnorm(alpha) * errors$se.fit
x$fit <- errors$fit
x$uci <- errors$fit + qnorm(alpha) * errors$se.fit

# plot points, model and confidence interval
ggplot(x, aes(x = RENEWAL_INSTANCE, y = fit)) +
  xlab("Renewal") + ylab("RTN") +
  theme_bw() + geom_line() + geom_smooth(aes(ymin = lci, ymax = uci), stat = "identity") +
  geom_point(data = std.east.service, aes(x = RENEWAL_INSTANCE, y = RTN, size = NEEDED_PREMIUM))

