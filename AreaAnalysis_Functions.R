# ==============================================================================
# Functions for VOC Data Analysis - Bachelor Thesis
# Author: Bart Bruijnen
# Institution: Maastricht University, FHML
# Date: 07-07-2025
#
# This R script contains custom functions used in the main analysis pipeline for
# the bachelor thesis titled:
# "The Effect of Confounding Factors on the Volatile Organic Compound (VOC)
# Composition of Human Exhaled Breath".
#
# Functions included:
# - ROClassAnalysis(): Performs binary or multiclass ROC curve analyses
#   for VOC markers and saves results.
# - ConfounderEffectAnalysis(): Tests the influence of potential confounders on
#   selected markers using appropriate statistical tests.
#
# Usage:
# Source this file into your main script using source("AreaAnalysis_Functions.R").
# Then call the functions with appropriate file paths as arguments.
#
# Ensure required packages (readxl, writexl, pROC) are installed before running.
# ==============================================================================

# ----------------------------
# Function: ROCClassAnalysis
# ----------------------------
# Performs ROC analysis (binary or multiclass) for each VOC area marker in a dataset.
# - file_path: path to Excel file with data
# - file_name: name of output Excel file to save ROC results
#
# Prompts user to specify the grouping variable (response) column name.
# Saves the ROC metrics (AUC, Youden Index, Sensitivity, Specificity, etc.) to output.
# Supports binary and one-vs-rest multiclass ROC computation.
ROCClassAnalysis <- function(file_path, file_name) {
  
  # Load dataset from Excel file
  df <- read_excel(file_path)
  print("Available columns:")
  print(names(df))  # Show column names for user reference
  
  # Prompt user to input the name of the grouping variable (response)
  response_col <- readline(prompt = "Enter the name of the primary grouping variable (e.g., Group): ")
  
  # Check if user input matches any column names
  if (!(response_col %in% names(df))) {
    stop("Selected grouping variable not found in dataset.")
  }
  
  # Select predictor columns starting with "area_" (VOC markers/compounds)
  predictor_vars <- grep("^area_", names(df), value = TRUE)
  
  # Initialize empty data frame to collect ROC results containing the values 
  # shown below with their data type after.
  all_results <- data.frame(
    Area = character(),
    Class = character(),
    AUC = numeric(),
    Youden_Index = numeric(),
    Optimal_Cutoff = numeric(),
    Sensitivity = numeric(),
    Specificity = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Extract the response variable column using the user-specified column name
  response_raw <- df[[response_col]]
  # Identify all unique classes or categories present in the response variable
  classes <- unique(response_raw)
  # Count the total number of unique classes to determine if it's binary or multiclass classification
  n_classes <- length(classes)
  
  # Check if at least two classes are present
  if (n_classes < 2) {
    stop("Not enough classes in selected grouping variable.")
  }
  
  # Add header row indicating the grouping variable used
  all_results <- rbind(
    all_results,
    data.frame(
      Area = paste0("--- ROC for Response: ", response_col, " ---"),
      Class = NA,
      AUC = NA,
      Youden_Index = NA,
      Optimal_Cutoff = NA,
      Sensitivity = NA,
      Specificity = NA,
      stringsAsFactors = FALSE
    )
  )
  
  # Loop through each VOC area marker
  for (predictor_col in predictor_vars) {
    predictor <- df[[predictor_col]]
    
    # Skip if predictor is not numeric, otherwise analysis won't work
    if (!is.numeric(predictor)) next
    
    # Binary classification case
    if (n_classes == 2) {
      response <- as.numeric(as.character(response_raw))
      
      # Compute ROC curve using pROC package
      roc_obj <- pROC::roc(response = response, predictor = predictor)
      
      # Extract sensitivities, specificities, and thresholds
      sens <- roc_obj$sensitivities
      spec <- roc_obj$specificities
      thr <- roc_obj$thresholds
      
      # Calculate Youden Index and find optimal threshold index
      youden <- sens + spec - 1
      max_idx <- which.max(youden)
      
      # Append results for current VOC marker/compound and bind the values
      # based on the max youden
      all_results <- rbind(
        all_results,
        data.frame(
          Area = predictor_col,
          Class = paste0("comparison"),
          AUC = as.numeric(roc_obj$auc),
          Youden_Index = youden[max_idx],
          Optimal_Cutoff = thr[max_idx],
          Sensitivity = sens[max_idx],
          Specificity = spec[max_idx],
          stringsAsFactors = FALSE
        )
      )
    } else {
      # Multiclass: one-vs-rest approach
      for (cls in classes) {
        # Create binary response for current class vs rest
        response <- ifelse(response_raw == cls, 1, 0)
        
        # Compute ROC curve
        roc_obj <- pROC::roc(response = response, predictor = predictor)
        
        # Extract metrics
        sens <- roc_obj$sensitivities
        spec <- roc_obj$specificities
        thr <- roc_obj$thresholds
        youden <- sens + spec - 1
        max_idx <- which.max(youden)
        
        # Append results per class-marker combination
        all_results <- rbind(
          all_results,
          data.frame(
            Area = predictor_col,
            Class = paste0(cls, " vs rest"),
            AUC = as.numeric(roc_obj$auc),
            Youden_Index = youden[max_idx],
            Optimal_Cutoff = thr[max_idx],
            Sensitivity = sens[max_idx],
            Specificity = spec[max_idx],
            stringsAsFactors = FALSE
          )
        )
      }
    }
  }
  
  # Save all results to Excel file
  write_xlsx(all_results, file_name)
  cat("ROC analysis complete and saved to:", file_name, "\n")
}
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# -----------------------------------
# Function: ConfounderEffectAnalysis
# -----------------------------------
# Evaluates if selected VOC markers (areas) differ significantly across levels of
# potential confounding variables.
#
# Inputs:
# - full_data_path: Excel file with full dataset (binary-coded confounders + numeric areas)
# - top_areas_path: Excel file with selected top markers (areas) from ROC analysis
# - output_file: name of output Excel file for confounder analysis results
#
# For each confounder grouping variable ending in "_bin" or factor,
# performs appropriate tests based on group count and normality:
# - Two groups: t-test (normal) or Wilcoxon rank-sum (non-normal)
# - Multiple groups: ANOVA (normal) or Kruskal-Wallis (non-normal)
#
# Saves results as Excel sheets per confounder.
ConfounderEffectAnalysis <- function(full_data_path, top_areas_path, output_file) {
  
  # Load the full dataset with potential confounders
  full_data <- read_excel(full_data_path)
  # Load top markers (areas)
  top_markers <- read_excel(top_areas_path)
  # Extract unique area names from top markers
  selected_areas <- unique(top_markers$Area)
  # Identify confounder grouping variables: those ending with '_bin' or factor type
  grouping_vars <- names(full_data)[grepl("_bin$", names(full_data)) | sapply(full_data, is.factor)]
  
  # Error if no grouping variables found
  if (length(grouping_vars) == 0) {
    stop("No grouping variables ending in '_bin' or factors found.")
  }
  
  all_results <- list()  # Store results per confounder group
  
  # Loop through each grouping variable (potential confounder)
  for (group in grouping_vars) {
    
    # Prepare empty data frame to store test results for current (potential) confounder
    group_results <- data.frame(
      Area = character(),
      Group = character(),
      Test = character(),
      P_Value = numeric(),
      Significant = character(),
      Normal_Distribution = character(),
      stringsAsFactors = FALSE
    )
    
    # Extract the values of the current grouping variable (e.g., age_bin, sex_bin) from the dataset
    group_vals <- full_data[[group]]
    
    # Skip confounder if less than 2 unique groups present
    if (length(unique(na.omit(group_vals))) < 2) next
    
    # Test each selected area (marker) for differences across groups
    for (area in selected_areas) {
      
      # Skip if area not in full dataset
      if (!(area %in% names(full_data))) {
        warning(paste("Skipping", area, "- not found in full dataset"))
        next
      }
      
      # Extract the values of the current VOC marker (area) from the dataset for 
      # statistical testing
      predictor <- full_data[[area]]
      
      # Skip non-numeric predictors
      if (!is.numeric(predictor)) next
      
      # Remove missing values
      df_subset <- na.omit(data.frame(group = group_vals, predictor = predictor))
      
      # Skip if fewer than 6 observations remain
      if (nrow(df_subset) < 6) next
      
      # Identify the unique groups or categories within the current (potential) 
      # confounder variable (e.g., sex_bin, age_bin)
      group_levels <- unique(df_subset$group)
      
      
      # Check normality for each group using Shapiro-Wilk test
      normality <- sapply(group_levels, function(g) {
        vals <- df_subset$predictor[df_subset$group == g]
        if (length(vals) < 3) return(NA)  # Not enough data for test
        shapiro.test(vals)$p.value > 0.05
      })
      
      # Determine if all groups within the current confounder variable follow a normal distribution
      # This will inform whether to use parametric (e.g., t-test, ANOVA) or non-parametric tests
      all_normal <- all(normality == TRUE, na.rm = TRUE)
      
      # Initialize placeholders for the type of statistical test and the resulting p-value
      test_type <- NA  # Will store which test is used (e.g., "t-test", "Wilcoxon", etc.)
      p_val <- NA      # Will store the computed p-value from the test
      
      # Choose test based on number of groups and normality
      if (length(group_levels) == 2) {
        if (all_normal) {
          test_type <- "t-test"
          p_val <- t.test(predictor ~ group, data = df_subset)$p.value
        } else {
          test_type <- "Wilcoxon"
          p_val <- wilcox.test(predictor ~ group, data = df_subset)$p.value
        }
      } else {
        if (all_normal) {
          test_type <- "ANOVA"
          p_val <- summary(aov(predictor ~ group, data = df_subset))[[1]][["Pr(>F)"]][1]
        } else {
          test_type <- "Kruskal-Wallis"
          p_val <- kruskal.test(predictor ~ group, data = df_subset)$p.value
        }
      }
      
      # Append results for current area-group test
      group_results <- rbind(group_results, data.frame(
        Area = area,
        Group = group,
        Test = test_type,
        P_Value = round(p_val, 5),
        Significant = ifelse(!is.na(p_val) & p_val < 0.05, "Yes", "No"),
        Normal_Distribution = ifelse(all_normal, "Yes", "No"),
        stringsAsFactors = FALSE
      ))
      
      # Print test summary to console
      cat(area, "vs", group, "→", test_type, "| p =", round(p_val, 4), "\n")
    }
    
    # Save results per confounder if any tests were performed
    if (nrow(group_results) > 0) {
      all_results[[group]] <- group_results
    }
  }
  
  # Write all confounder test results to separate sheets in an Excel file
  write_xlsx(all_results, output_file)
  # Return a confirmation message with the output file path
  return(paste("Confounder effect analysis complete. Results saved to:", output_file))
}
# ==============================================================================
# End of AreaAnalysis_Functions.R

# You can now source this file in your main analysis script using:
# source("AreaAnalysis_Functions.R")