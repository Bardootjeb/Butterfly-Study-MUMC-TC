# ==============================================================================
# Script for Data Analysis – Bachelor Thesis
#
# Author: Bart Bruijnen  
# Institution: Maastricht University, FHML  
# Supervisors: Prof. Dr. S.M.E. Engelen, PhD W. Lasten  
# Date: 07-07-2025
#
# Description:
# This R script contains all code used for the statistical analysis and visualization 
# of the data collected for the bachelor thesis project.
#
# This script supports the analyses presented in the thesis entitled:
# "The Effect of Confounding Factors on the Volatile Organic Compound (VOC) Composition 
# of Human Exhaled Breath".
#
# It includes:
# - Data preprocessing
# - Binary or Multiclass ROC analysis
# - Selection of top discriminatory markers
# - Evaluation of potential confounding effects
# - Export of results for reporting
# ==============================================================================
# Before running this script, ensure all required R packages are installed.
# Use the following command to install missing packages automatically.

# Create a vector containing the required packages. Check if installed and if not, install.
packages <- c("readxl", "dplyr", "pROC","writexl")
installed <- packages %in% rownames(installed.packages())
if (any(!installed)) {
  install.packages(packages[!installed])
} else {
  message("Packages are already installed")
}
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Step 1. Data Preprocessing

# In this step, we prepare the raw dataset for downstream analysis.
# This includes:
# - Importing necessary libraries and external R functions
# - Loading the original Excel file
# - Standardizing column names
# - Recoding categorical variables into binary or ordinal factors
# - Saving the cleaned dataset to a new Excel file
# Source files and require data

# Set working directory to local environment (adjust as needed)
setwd("~/Maastricht University/Biomedical Sciences/BMS year 3/BBS3006 - Thesis & Internship/Butterfly")

# Load required R packages
source("AreaAnalysis_Functions.R")      # Loads custom analysis functions
library(readxl)                         # For reading Excel files (.xlsx)
library(dplyr)                          # For data manipulation (e.g., filtering, mutating)
library(pROC)                           # For ROC curve analysis
library(writexl)                        # For writing Excel files (.xlsx)

# Load Excel spreadsheet with patient metadata and breath VOC intensities 
# and turn into dataframe
rawdata <- read_excel("Breath_Rawdata.xlsx") 

# Clean column names to lowercase and simplify.
names(rawdata) <- tolower(gsub("[^[:alnum:]_]", "_", names(rawdata))) 
# Replaces spaces/special characters with underscores (e.g., "Tumor Classification" → "tumor_classification")

# Recode Tumor Classification → Binary (Benign = 0, Malignant = 1). 
rawdata$tumor_bin <- ifelse(tolower(rawdata$tumor_classification) == "malignant", 1,
                   ifelse(tolower(rawdata$tumor_classification) == "benign", 0, NA)) 
# Numeric binary column `tumor_bin` with values 0 (benign), 1 (malignant), NA otherwise

# Recode Fasting → Binary (Ja/Yes = 1, Nee/No = 0)
rawdata$fasting_bin <- ifelse(tolower(rawdata$fasting) %in% c("ja", "yes"), 1,
                         ifelse(tolower(rawdata$fasting) %in% c("nee", "no"), 0, NA))
# Numeric binary column `fasting_bin` with 0 (Non-fasted) or 1 (Fasted), NA otherwise

# Insert sex_bin after sex
rawdata$sex_bin <- ifelse(tolower(rawdata$sex) == "m", 1,
                     ifelse(tolower(rawdata$sex) == "v", 2, NA))
# Numeric binary column `sex_bin` with 1 (Male) or 2 (Female), NA otherwise

# Recode BMI to ordinal categories:
#   1 = normal weight (0–24.9)
#   2 = overweight (25–29.9)
#   3 = obese (30+)
bmi_bin <- which(names(rawdata) == "bmi")  # Find index of "bmi" column
rawdata$bmi_bin <- cut(
  rawdata$bmi,
  breaks = c(0, 24.9, 29.9, Inf),
  labels = 1:3,
  right = TRUE,
  include.lowest = TRUE
) # Factor column `bmi_bin` with levels 1–3

# Recode Age to ordinal bins by decade (starting from <30)
#   1 = 0–29, 2 = 30–39, ..., 8 = 90+
age_bin <- which(names(rawdata) == "age")  # Find index of "age" column
rawdata$age_bin <- cut(
  rawdata$age,
  breaks = c(0, 29, 39, 49, 59, 69, 79, 89, Inf),
  labels = 1:8,
  right = TRUE,
  include.lowest = TRUE
) # Factor column `age_bin` with levels 1–8 (age groups by decade)

# Recode Smoking history to ordinal categories:
#   1 = non-smoker, 2 = ex-smoker, 3 = current smoker
smoke_bin <- which(names(rawdata) == "smoking_history")
rawdata$smoke_bin <- ifelse(
  tolower(rawdata$smoking_history) == "non", 1,
  ifelse(tolower(rawdata$smoking_history) == "ex", 2,
         ifelse(tolower(rawdata$smoking_history) == "current", 3, NA))
) # Numeric column `smoke_bin` with levels 1–3

# Save cleaned version
write_xlsx(rawdata, "Breath_Metadata.xlsx")

# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Step 2. Binary or Multiclass ROC analysis

# This function performs ROC analysis for each area (VOC compound) across 
# multiple diagnostic classes. It reads the metadata which includes the 
# classification labels & potential confounding factors, and generates ROC curves 
# for each marker against each class using a one-vs-rest approach.
# For each class-marker combination, it calculates performance metrics including AUC, 
# sensitivity, specificity, and Youden Index. The resulting performance metrics are 
# saved for further filtering and selection of top markers.
ROCClassAnalysis("Breath_Metadata.xlsx", "Breath_RCA.xlsx")

# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Step 3. Selection of top discriminatory markers

# Load ROC results from Excel file into a data frame called 'results'
results <- read_excel("Breath_RCA.xlsx")

# Filter the results based on Youden Index > 0.3 and Specificity >= 0.65
# This selects markers with moderate discriminatory power and acceptable specificity
filtered <- results[results$Youden_Index > 0.3 & results$Specificity >= 0.65, ]

# Remove any rows where the 'Class' column contains NA values
# Ensures only valid class labels are included for downstream analysis
filtered <- filtered[!is.na(filtered$Class), ]

# Sort the filtered results by Youden Index in descending order
filtered <- filtered[order(-filtered$Youden_Index), ]

# Print the sorted filtered results to the console
print(filtered)

# Save the sorted filtered results to an Excel file
write_xlsx(filtered, "Final Areas.xlsx")

# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Step 4. Evaluation of potential confounding effects

# This function assesses whether the top discriminatory markers (areas/VOCs) identified
# might be influenced by potential confounding variables. It takes two Excel files as input:
# - full_data_path: a dataset including all patients with binary-coded confounders and factor variables
# - top_areas_path: the selected top markers (areas) from prior ROC analysis

# For each grouping variable identified by the suffix "_bin" or factor type, 
# the function tests whether the distribution of each selected marker differs significantly 
# across the groups defined by the potential confounder.

# Depending on the number of groups and data normality, it chooses appropriate statistical tests:
# - For two groups: t-test (if normal) or Wilcoxon rank-sum test (if non-normal)
# - For more than two groups: ANOVA (if normal) or Kruskal-Wallis test (if non-normal)

# It outputs an Excel file with separate sheets per confounder group summarizing:
# - The statistical test used
# - p-values for differences in marker distributions between groups
# - Whether the difference is statistically significant (p < 0.05)
# - Whether the data passed normality assumptions

# This helps identify markers potentially influenced by confounders, guiding further interpretation.
ConfounderEffectAnalysis(
  full_data_path = "Breath_Metadata.xlsx",
  top_areas_path = "Final Areas.xlsx",
  output_file = "Breath_CEA.xlsx"
)
# ==============================================================================
# End of script
print("All analyses completed successfully.")
print("Results have been saved to the specified output files.")
print("Thank you for using this analysis pipeline.")