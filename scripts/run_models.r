# install
install.packages("dplyr")
install.packages("ggplot2")

# load packages
library("dplyr")
library("ggplot2")

# load in data model was trained on
load(file = "~/data.RData")

# load models and training index created on hcc
load(file = "~/model-rf.RData")

# fetch the data that was used to train by passing in the training index
training.data <- data[training.index, ]

# fetch the data that was used for validation
validation.data <- data[-training.index, ]

# summarize accuracy of models
results <- resamples(list(knn = fit.knn, rf = fit.rf))
summary(results)

# compare model accuracy & cohen kappa
dotplot(results)

# display model
print(fit.rf)

# ======================================== ACTUAL VS PREDICTED ========================================

# sampling size
n <- 1:100

# fetch random sample from data
sampling <- data %>% sample_n(length(n))

# run sample through model to get predictions
sampling.predictions <- predict(fit.rf, sampling)

# fetch actual and predicted values
actual <- sampling$RTN
predicted <- sampling.predictions

# scatter plot (actual vs predicted) of random sample
ggplot(data.frame(n, actual, predicted)) +
  geom_point(aes(x = n, y = actual, col = "Actual")) +
  scale_x_discrete(limits = n) +
  theme(axis.text.x = element_blank()) +
  xlab("Sample") + ylab("RTN") +
  geom_point(aes(x = n, y = predicted, col = "Predicted"))

# ======================================== RECONSTRUCT RTN ============================================

# set renewal instances
renewal <- c(0, 1, 2, 3)

# select a specific group
group <- data.frame(COVERAGE_TYPE = "Life", 
                ELIGIBLE_LIVES = 2500, 
                INDUSTRY = "Services", 
                NEEDED_PREMIUM = 50000, 
                UW_REGION = "east")

# create a group row for each renewal and populate the renewal value
group <- bind_rows(replicate(length(renewal), group, simplify = FALSE))
group$RENEWAL_INSTANCE <- renewal

# make predictions for the group
group.predictions <- predict(fit.rf, group)

# plot only the predicted points vs renewal
plot(renewal, group.predictions, main = "STD Service East", xlab = "Renewal", ylab = "RTN")
grid(NA, 10, lwd = 2) # grid only in y-direction
lines(renewal, group.predictions)
