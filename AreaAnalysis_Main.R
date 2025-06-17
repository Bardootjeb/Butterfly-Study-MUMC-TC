# Script for Data Analysis – Bachelor Thesis
# 
# Author: Bart Bruijnen
# Institution: Maastricht University, FHML
# Supervisors:  Prof. Dr. S.M.E. Engelen, PhD. W. Lasten
# Date: 07-07-2025
#
# This R script contains all code used for the statistical analysis and visualization 
# of the data collected for the bachelor thesis project. 
#
# Before running this script, please ensure that all required R packages are installed.
# The following commands can be used to install and load the necessary packages.

# This script supports the analysis presented in the thesis entitled 
# "The Effect of Confounding Factors on Volatile Organic Compound Composition 
# of Human Exhaled Breath". It includes data preprocessing, statistical testing, 
# and the creation of relevant figures and tables.

# Create a vector containing the required packages
packages <- c("ggplot2", "readxl", "lubridate", "pROC","writexl")
installed <- packages %in% rownames(installed.packages())
if (any(!installed)) {
  install.packages(packages[!installed])
} else {
  message("Packages are already installed")
}

# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Step 4. Calculating Youden Index

library(pROC)
library(writexl)

# Load original file
df <- read_excel("BreathPatients.xlsx")

# Clean column names to lowercase and simplify
names(df) <- tolower(gsub("[^[:alnum:]_]", "_", names(df)))

# Recode Tumor Classification → Binary (Benign = 0, Malignant = 1)
df$tumor_bin <- ifelse(tolower(df$tumor_classification) == "malignant", 1,
                   ifelse(tolower(df$tumor_classification) == "benign", 0, NA))

# Recode Fasting → Binary (Ja/Yes = 1, Nee/No = 0)
df$fasting_bin <- ifelse(tolower(df$fasting) %in% c("ja", "yes"), 1,
                         ifelse(tolower(df$fasting) %in% c("nee", "no"), 0, NA))

# Insert sex_bin after sex
df$sex_bin <- ifelse(tolower(df$sex) == "m", 1,
                     ifelse(tolower(df$sex) == "v", 2, NA))


# Insert bmi_bin after bmi
bmi_index <- which(names(df) == "bmi")
df$bmi_bin <- cut(df$bmi,
                  breaks = c(0, 24.9, 29.9, Inf),
                  labels = 1:3,
                  right = TRUE,
                  include.lowest = TRUE)

# Insert age_bin after age
age_index <- which(names(df) == "age")
df$age_bin <- cut(df$age,
                  breaks = c(0, 29, 39, 49, 59, 69, 79, 89, Inf),
                  labels = 1:8,
                  right = TRUE,
                  include.lowest = TRUE)

# Insert smoke_bin after smoking_history
smoke_index <- which(names(df) == "smoking_history")
df$smoke_bin <- ifelse(tolower(df$smoking_history) == "non", 1,
                       ifelse(tolower(df$smoking_history) == "ex", 2,
                              ifelse(tolower(df$smoking_history) == "current", 3, NA)))

# Save cleaned version
write_xlsx(df, "BreathPatients_bin.xlsx")

MultiClassAnalysis("BreathPatients_bin.xlsx", "RAW_Breath_Patients_ROC_MCA.xlsx")

# Load ROC results
results <- read_excel("RAW_Breath_Patients_ROC_MCA.xlsx")

# Filter by Youden Index > 0.3 and Specificity >= 0.65
filtered <- results[results$Youden_Index > 0.3 & results$Specificity >= 0.65, ]

# Remove any rows with NA in the Class column
filtered <- filtered[!is.na(filtered$Class), ]

# Identify best-performing marker for each Area-Class pair (max specificity)
combinations <- unique(filtered[, c("Area", "Class")])
top_specificity <- data.frame()

for (i in seq_len(nrow(combinations))) {
  area_i <- combinations$Area[i]
  class_i <- combinations$Class[i]
  subset_data <- filtered[filtered$Area == area_i & filtered$Class == class_i, ]
  top_row <- subset_data[which.max(subset_data$Specificity), ]
  top_specificity <- rbind(top_specificity, top_row)
}

# Sort by specificity
top_specificity <- top_specificity[order(-top_specificity$Specificity), ]
print(top_specificity)

# Save
write_xlsx(top_specificity, "Breath_Patients_TopMarkers.xlsx")

ConfounderEffectAnalysis(
  full_data_path = "BreathPatients_bin.xlsx",
  top_markers_path = "Breath_Patients_TopMarkers.xlsx"
)





