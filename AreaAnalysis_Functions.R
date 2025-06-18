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
# Evaluates whether selected VOC markers (areas) differ significantly across levels 
# of potential confounders using appropriate tests based on group number and normality.
#
# Inputs:
# - full_data_path: path to Excel file containing full dataset (confounders + numeric areas)
# - top_areas_path: path to Excel file with selected top markers (areas) to test
# - output_file: filename for saving the results as Excel sheets
#
# For each grouping variable (ending with "_bin" or factor), this function:
# - Tests normality within groups (Shapiro-Wilk test)
# - Performs t-test/Wilcoxon for two groups; ANOVA/Kruskal-Wallis for multiple groups
# - Computes p-values and flags significance
# - Saves results to Excel with one sheet per confounder variable
ConfounderEffectAnalysis <- function(full_data_path, top_areas_path, output_file) {
  
  # Read full dataset (with confounders and area values)
  full_data <- read_excel(full_data_path)
  # Read top areas/markers identified for testing
  top_markers <- read_excel(top_areas_path)
  # Extract unique area names from top markers
  selected_areas <- unique(top_markers$Area)
  
  # Identify grouping variables by either factor type or names ending with '_bin'
  grouping_vars <- names(full_data)[grepl("_bin$", names(full_data)) | sapply(full_data, is.factor)]
  
  # Stop execution if no suitable grouping variables found
  if (length(grouping_vars) == 0) stop("No grouping variables ending in '_bin' or factors found.")
  
  all_results <- list()  # Initialize list to hold results per confounder
  
  # Loop over each grouping variable (potential confounder)
  for (group in grouping_vars) {
    
    # Extract values of the current grouping variable
    group_vals <- full_data[[group]]
    
    # Skip if fewer than two unique groups exist (no meaningful test)
    if (length(unique(na.omit(group_vals))) < 2) next
    
    results_list <- list()  # Initialize storage for test results per area for this group
    
    # Loop through each selected VOC marker area
    for (area in selected_areas) {
      
      # Skip area if not found in dataset columns
      if (!(area %in% names(full_data))) {
        warning(paste("Skipping", area, "- not found in full dataset"))
        next
      }
      
      # Extract predictor values for the area
      predictor <- full_data[[area]]
      # Continue only if predictor is numeric
      if (!is.numeric(predictor)) next
      
      # Create subset dataframe with non-missing group and predictor values
      df_subset <- na.omit(data.frame(group = group_vals, predictor = predictor))
      
      # Skip analysis if fewer than 6 observations remain
      if (nrow(df_subset) < 6) next
      
      # Identify unique groups within this confounder variable
      group_levels <- unique(df_subset$group)
      
      # Test normality for each group using Shapiro-Wilk test
      normality <- sapply(group_levels, function(g) {
        vals <- df_subset$predictor[df_subset$group == g]
        # Skip normality test if fewer than 3 observations in group
        if (length(vals) < 3) return(NA)
        shapiro.test(vals)$p.value > 0.05  # TRUE if data is normal
      })
      
      # Determine if all groups are normally distributed (ignoring NAs)
      all_normal <- all(normality == TRUE, na.rm = TRUE)
      
      # Initialize placeholders for test metadata and results
      test_type <- effect_label <- NA
      p_val <- test_stat <- effect_size <- NA
      
      # Choose appropriate statistical test based on group count and normality
      if (length(group_levels) == 2) {
        if (all_normal) {
          # Two groups, normal data: independent two-sample t-test
          test_type <- "t-test"
          test <- t.test(predictor ~ group, data = df_subset)
          p_val <- test$p.value
          test_stat <- unname(test$statistic)
          # Compute Cohen's d effect size (parametric)
          d <- cohen.d(df_subset$predictor, df_subset$group)
          effect_size <- unname(d$estimate)
          effect_label <- "Cohen_d"
        } else {
          # Two groups, non-normal data: Wilcoxon rank-sum test
          test_type <- "Wilcoxon"
          test <- wilcox.test(predictor ~ group, data = df_subset)
          p_val <- test$p.value
          test_stat <- unname(test$statistic)
          # Nonparametric effect size with Cohen's d alternative
          d <- cohen.d(df_subset$predictor, df_subset$group, nonparam = TRUE)
          effect_size <- unname(d$estimate)
          effect_label <- "Cohen_d (nonparam)"
        }
      } else {
        # More than two groups
        if (all_normal) {
          # Parametric ANOVA for normal data
          test_type <- "ANOVA"
          model <- aov(predictor ~ group, data = df_subset)
          p_val <- summary(model)[[1]][["Pr(>F)"]][1]
          anova_table <- anova(model)
          test_stat <- anova_table[[1]][1]
          ss_total <- sum(anova_table[[1]][, "Sum Sq"])
          ss_between <- anova_table[["Sum Sq"]][1]
          # Eta-squared effect size: variance explained by group
          effect_size <- ss_between / ss_total
          effect_label <- "Eta_squared"
        } else {
          # Nonparametric Kruskal-Wallis test for non-normal data
          test_type <- "Kruskal-Wallis"
          test <- kruskal.test(predictor ~ group, data = df_subset)
          p_val <- test$p.value
          test_stat <- unname(test$statistic)
          H <- test_stat
          k <- length(group_levels)
          n <- nrow(df_subset)
          # Eta-squared effect size estimate for Kruskal-Wallis
          effect_size <- (H - k + 1) / (n - k)
          effect_label <- "KW_Eta2"
        }
      }
      
      # Create clean group name by removing trailing "_bin" and converting to Title Case
      name <- gsub("_bin$", "", group)
      name <- tools::toTitleCase(nice_group_name)
      
      # Store current test result as a named list with dynamic column names for stats/effect sizes
      this_result <- list(
        Area = area,
        Group = name,
        Test = test_type,
        P_Value = round(p_val, 5)
      )
      this_result[[paste0(test_type, "_statistic")]] <- round(test_stat, 5)
      this_result[[effect_label]] <- round(effect_size, 5)
      
      # Append this result to the results list for current grouping variable
      results_list[[length(results_list) + 1]] <- this_result
      
      # Print concise summary to console for progress monitoring
      cat(area, "vs", name, "→", test_type, "| p =", round(p_val, 5), 
          "|", paste0(test_type, "_statistic"), "=", round(test_stat, 5), 
          "|", effect_label, "=", round(effect_size, 5), "\n")
    }
    
    # After looping all areas, combine results to data frame if any exist
    if (length(results_list) > 0) {
      df_results <- do.call(rbind, lapply(results_list, function(x) as.data.frame(x, stringsAsFactors = FALSE)))
      
      # Convert numeric-looking columns from character to numeric explicitly
      for (col in names(df_results)) {
        if (all(is.na(df_results[[col]]) | grepl("^[0-9\\.]+$", df_results[[col]]))) {
          df_results[[col]] <- as.numeric(df_results[[col]])
        }
      }
      
      # Store final results data frame for this grouping variable in output list
      all_results[[group]] <- df_results
    }
  }
  
  # Write all confounder test results to Excel, one sheet per confounder
  write_xlsx(all_results, output_file)
  
  # Return confirmation message with output file path
  return(paste("Confounder effect analysis complete. Results saved to:", output_file))
}
# ==============================================================================
# End of AreaAnalysis_Functions.R

# You can now source this file in your main analysis script using:
# source("AreaAnalysis_Functions.R")