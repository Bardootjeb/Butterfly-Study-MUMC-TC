# Environmental Data Analysis for Exhaled Breath VOC Studies

This project contains an R script designed to analyze and visualize room temperature and relative humidity (RH) data. 
The data was collected as part of a bachelor’s thesis investigating how environmental factors might affect the volatile organic compound (VOC) profiles in human exhaled breath.
The analysis focuses on variations in temperature and humidity throughout the workday at the sampling location, which is important because these factors can influence VOC measurements.

## Project Files

- **Thesis_Main.R** — Main script that performs data analysis and visualization.  
- **Thesis_Functions.R** — Helper functions used by the main script, such as plotting and saving outputs.  
- **HOLDING.xlsx** — Excel file containing the raw environmental data.  
- **Output/** — Folder where generated plot files (PDFs) are saved.  

## Getting Started

### Prerequisites

Make sure you have R and RStudio installed on your computer.

### Installing Required Packages

The analysis depends on several R packages: `ggplot2`, `readxl`, and `lubridate`. To install any missing packages, run the following in R:

```r
packages <- c("ggplot2", "readxl", "lubridate")
installed <- packages %in% rownames(installed.packages())
if (any(!installed)) {
  install.packages(packages[!installed])
}

```

## Running the Script
Download or clone this repository to your local machine.
Place HOLDING.xlsx in the root directory of the cloned repository.
Open Thesis_Main.R in RStudio.
Set the working directory to the folder where the repository is saved, for example:

```R

setwd("~/Maastricht University/Biomedical Sciences/BMS year 3/BBS3006 - Thesis & Internship/Butterfly")
# NOTE: Change this path to your local repository path
```

Execute the entire script

# Script Functionality
The script performs the following key steps:

## Data Preparation:

Imports room temperature and relative humidity data from HOLDING.xlsx.

Converts relevant columns to numeric format.

Creates new variables for mean and range of RH and temperature.

Categorizes time of day into intervals (e.g., morning, midday, afternoon).

## Visualization:

Generates boxplots to show how mean RH and temperature vary across times of day.

Saves plots as PDFs in the Output/ folder.

## Statistical Analysis:

Conducts Friedman tests and repeated measures ANOVA to assess how time of day affects environmental conditions.
Evaluates the impact on max, min, mean, and range values for both RH and temperature.

## Why This Matters
Environmental factors such as temperature and humidity can affect the stability and measurement of VOCs in breath samples. 
High humidity, for example, can interfere with analytical instruments and lead to inaccurate VOC detection. 
By analyzing these variables, this project helps clarify how environmental conditions may confound VOC studies and supports more reliable breath analysis.
