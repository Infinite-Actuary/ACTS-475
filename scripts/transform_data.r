# install
install.packages("readxl")
install.packages("writexl")
install.packages("dplyr")

# load packages
library("readxl")
library("writexl")
library("dplyr")

# ======================================== S C R U B B I N G ========================================

# set file names
fileName <- "~/RtnData.xlsx"

# read in the data
dataSheet <- read_excel(input, sheet = "Data")
sicSheet <- read_excel(input, sheet = "SIC")

# left outer-join on industry code
rtnData <- merge(x = dataSheet, y = sicSheet, by.x = "INDUSTRY_CODE", by.y = "SIC_CODE", all.x = TRUE)

# drop VAL_DATE column
rtnData <- rtnData %>% select(-c("VAL_DATE"))

# remove NBOC and Amendment rows
rtnData <- rtnData %>% filter(EVENT_TYPE != "NBOC")
rtnData <- rtnData %>% filter(EVENT_TYPE != "Amendment")

# use only positive premiums
rtnData <- rtnData %>% filter(NEEDED_PREMIUM > 0 & FINAL_QUOTED_PREMIUM > 0)

# use plans with positive enrolled
rtnData <- rtnData %>% filter(ENROLLED_LIVES > 0)

# eligible should be >= enrolled
rtnData <- rtnData %>% mutate(ELIGIBLE_LIVES = case_when(ELIGIBLE_LIVES < ENROLLED_LIVES ~ ENROLLED_LIVES,
                                                         TRUE ~ ELIGIBLE_LIVES))

# standardize renewals
rtnData <- rtnData %>% mutate(RENEWAL_INSTANCE = case_when(EVENT_TYPE == "New Business" ~ "0",
                                                           RENEWAL_INSTANCE == "1st Renewal" ~ "1",
                                                           RENEWAL_INSTANCE == "2nd Renewal" ~ "2",
                                                           RENEWAL_INSTANCE == "2nd+" ~ "3",
                                                           RENEWAL_INSTANCE == "3rd Renewal +" ~ "3",
                                                           TRUE ~ as.character(RENEWAL_INSTANCE)))
                                                           
# remove blank renewal instances                                                           
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

# sales office to sales region map
rtnData <- rtnData %>% mutate(UW_REGION = tolower(UW_REGION))
regions <- list("home" = c("ft wayne"),
                "central" = c("chicago","cincinnati","cleveland","detroit","ft lauderdale","indianapolis","miami","milwaukee","minneapolis","omaha","orlando","pittsburgh","st. louis","tampa"),
                "east" = c("atlanta","boston","charlotte","long island","nashville","new jersey","new york","parsnippany","philadelphia","portland, me","rochester","washington d.c.","white plains"),
                "west" = c("dallas","denver","houston","kansas city","los angeles","orange county","phoenix","portland, or","sacramento","san diego","san francisco","seattle"))

# map region to office
rtnData <- rtnData %>% mutate(UW_REGION = case_when(#SALES_OFFICE %in% regions$home ~ "home",
                                                    SALES_OFFICE %in% regions$central ~ "central",
                                                    SALES_OFFICE %in% regions$east ~ "east",
                                                    SALES_OFFICE %in% regions$west ~ "west",
                                                    TRUE ~ "misc"))

# remove misc regions (i.e. home)
rtnData <- rtnData %>% filter(UW_REGION != "misc")

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

# normalize industry titles by removing spaces and commas
rtnData <- rtnData %>% mutate(INDUSTRY = stringr::str_replace_all(INDUSTRY, c("," = "", " " = "")))

# add rtn column
rtnData <- rtnData %>% mutate(RTN = FINAL_QUOTED_PREMIUM / NEEDED_PREMIUM)

# convert categorical columns to factor
rtnData <- rtnData %>% mutate_at(vars(COVERAGE_TYPE, GROUP_SIZE, INDUSTRY, RENEWAL_INSTANCE, UW_REGION), as.factor)

# remove extraneous info and rebuild data
data <- rtnData %>% select(c(RTN, COVERAGE_TYPE, ELIGIBLE_LIVES, INDUSTRY, NEEDED_PREMIUM, RENEWAL_INSTANCE, UW_REGION))
  
# omit rows with NA values
data <- na.omit(data)

# save data to disk
save(data, rtnData, rtn.breaks, file = "~/rtn_data.rdata")

# write rtn data to Excel
#write_xlsx(rtnData, path = "~/rtnDataClean.xlsl", col_names = TRUE)
