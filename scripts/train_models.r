# load packages
library("dplyr")
library("caret") # machine learning interface
library("randomForest")

# ======================================== T R A I N I N G ========================================

# load in data
load(file = "~/data.rdata")

# use index to partition data for training (90%) & validation
training.index <- createDataPartition(data$RTN, p = 0.90, list = FALSE)

# use p% for training the model
training.data <- data[training.index, ]

# use (1-p)% for data validation
validation.data <- data[-training.index, ]

# k-fold cross validation
control <- trainControl(method = "cv", number = 10)

# metric for evaluating model
metric <- "Rsquared"

# RF: Random Forest
fit.rf <- train(RTN ~ ., data = training.data, method = "rf", metric = metric, trControl = control)

# kNN: k-Nearest Neighbors
fit.knn <- train(RTN ~ ., data = training.data, method = "knn", metric = metric, trControl = control)

# save model to disk
save(fit.rf, fit.knn, training.index, file = "model.RData")
