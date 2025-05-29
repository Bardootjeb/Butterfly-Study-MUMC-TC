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


# Check if BiocManager is installed; install it if not
if (!requireNamespace("BiocManager")) {
  install.packages("BiocManager")
  BiocManager::install("biomaRt")
} else {
  message("BiocManager is already installed")
}

# Create vector containing all the packages
packages <- c("ggplot2", "readxl", "lubridate")

# Check if required packages is installed; install it if not
if (!requireNamespace(packages)) {
  BiocManager::install(packages)
} else {
  message("Packages are already installed")
}

install.packages("lubridate")
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Step 1. Preparatory Analysis

# Source files and require data
source("Thesis_Functions.R")
library(readxl)
library(lubridate)
library(ggplot2)


# Ensure the output directory exists
if (!dir.exists("Output")) {
  dir.create("Output")
}

# Import data from excel file and transform data into dataframe
weekdays <- read_excel("RH TEMP.xlsx")
weekdays <- as.data.frame(weekdays)
weekdays$Day <- as.factor(weekdays$Day)

# Assign numeric vectors to the selected collumns
titles <- c("RH Max (%)", "Temp Max (°C)", "RH Min (%)", "Temp Min (°C)")
weekdays[titles] <- lapply(weekdays[titles], 
                          function(x) as.numeric(gsub(",", ".", gsub(" ", "", x))))
# Extract numeric values from the dataframe
RH_Temp <- weekdays[3:6]
# Summary statistics
summary(RH_Temp) 
# Calculate the SDs
SD <- sapply(RH_Temp, sd, na.rm = TRUE)

# Rename collumn name
names(weekdays)[names(weekdays) == "Time of day"] <- "Time.of.Day"
# Turn "Time of Day" from character into a factor of time
weekdays$Time.of.Day <- hms::as_hms(weekdays$Time.of.Day)

save.pdf(function(){
}, "Gene Expression")

# Plot the RH
plot_range(weekdays, "RH Max (%)", "RH Min (%)", "Relative Humidity (%)", 
           "Relative Humidity Max and Min per Measurement")

# Plot the temp
plot_range(weekdays, "Temp Max (°C)", "Temp Min (°C)", "Temperature (°C)", 
           "Temperature Max and Min per Measurement")

weekdays$Time.Category <- cut(
  as.numeric(weekdays$Time.of.Day),
  breaks = c(0, 10*3600, 14*3600, 24*3600), # in seconds
  labels = c("08", "12", "16"),
  include.lowest = TRUE, right = FALSE
)

table(weekdays$Time.Category, weekdays$Day)

friedman.test(`RH Max (%)` ~ Time.Category | Day, data = weekdays)


# For RH Max
aov_max <- aov(`RH Max (%)` ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_max)

# For RH Min
aov_min <- aov(`RH Min (%)` ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_min)

weekdays$RH_Range <- weekdays$`RH Max (%)` - weekdays$`RH Min (%)`
aov_range <- aov(RH_Range ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_range)

weekdays$RH_Mean <- (weekdays$`RH Max (%)` + weekdays$`RH Min (%)`) / 2
aov_mean <- aov(RH_Mean ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_mean)
