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

# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Step 1. Preparatory Analysis

# Set working directory
setwd("~/Maastricht University/Biomedical Sciences/BMS year 3/BBS3006 - Thesis & Internship/Butterfly")
# Note: This only works if you were to have the same exact working directory

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
weekdays <- read_excel("HOLDING.xlsx")
weekdays <- as.data.frame(weekdays)
weekdays$Day <- as.factor(weekdays$Day)

# Assign numeric vectors to the selected collumns
titles <- c("RH Max (%)", "Temp Max (°C)", "RH Min (%)", "Temp Min (°C)")
weekdays[titles] <- lapply(weekdays[titles], 
                          function(x) as.numeric(gsub(",", ".", gsub(" ", "", x))))

# Create new values: RH_Mean and range
weekdays$RH_Mean <- (weekdays$`RH Max (%)` + weekdays$`RH Min (%)`) / 2
weekdays$RH_Range <- weekdays$`RH Max (%)` - weekdays$`RH Min (%)`

# Create new values: Temperature mean and range
weekdays$Temp_Mean <- (weekdays$`Temp Max (°C)` + weekdays$`Temp Min (°C)`) / 2
weekdays$Temp_Range <- weekdays$`Temp Max (°C)` - weekdays$`Temp Min (°C)`

# Extract numeric values from the dataframe 
RH_Temp <- weekdays[3:11]
# Summary statistics
summary(RH_Temp) 

# Rename collumn name
names(weekdays)[names(weekdays) == "Time of day"] <- "Time.of.Day"

# Transforming the data to fit analysis
weekdays$Time.Category <- cut(
  as.numeric(hms::as_hms(weekdays$Time.of.Day)), # Turn "Time of Day" from 
  #character into a factor of time and then into numeric data
  breaks = c(0, 10*3600, 14*3600, 24*3600),  # 0-10h, 10-14h, 14-24h
  labels = c("08", "12", "16"), # create three categories
  include.lowest = TRUE, right = FALSE
)


# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Step 2. Plot the RH and temperature range using the plot_range function 
# & save to PDF using the save.pdf function.

## The functions can be found in the "Thesis_Functions.R" file

save.pdf(function(){
}, "Gene Expression")

# Plot the RH
plot_range(weekdays, "RH_Mean", "Mean Relative Humidity", 
           "Relative Humidity Range per Measurement")
# Plot the temp
plot_range(weekdays, "Temp_Mean", "Mean Temperature", 
           "Temperature Range per Measurement")

# Old function. Im keeping it for now to see the trend per day

# Plot the RH
plot_range2(weekdays, "RH Max (%)", "RH Min (%)", "Relative Humidity (%)",
             "Relative Humidity Range per Measurement")
# Plot the temp
plot_range2(weekdays, "Temp Max (°C)", "Temp Min (°C)", "Temperature (°C)", 
           "Temperature Range per Measurement")


# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Step 3. Statistical Analysis

# Create table to check if every data point has a value = 1
table(weekdays$Time.Category, weekdays$Day)

# I'm aiming for 14 days of results, but in the meantime, I will have to use this
# test, which works with smaller data sets
friedman.test(`RH_Mean` ~ Time.Category | Day, data = weekdays)

# For RH Max
aov_max <- aov(`RH Max (%)` ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_max)

# For RH Min
aov_min <- aov(`RH Min (%)` ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_min)

# For RH Range
aov_range <- aov(RH_Range ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_range)

# For RH Mean
aov_mean <- aov(RH_Mean ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_mean)

# Now we do the same for the temperature

# Testing the changes for Temp
friedman.test(Temp_Mean ~ Time.Category | Day, data = weekdays)

# And repeating the steps for temperature
aov_temp_max <- aov(`Temp Max (°C)` ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_temp_max)

aov_temp_min <- aov(`Temp Min (°C)` ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_temp_min)

aov_temp_mean <- aov(Temp_Mean ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_temp_mean)

aov_temp_range <- aov(Temp_Range ~ Time.of.Day + Error(Day/Time.of.Day), data = weekdays)
summary(aov_temp_range)
