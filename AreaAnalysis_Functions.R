MultiClassAnalysis <- function(file_path, file_name) {
  
  # Load data
  df <- read_excel(file_path)
  print("Available columns:")
  print(names(df))
  
  # Ask user for response variable
  response_col <- readline(prompt = "Enter the name of the primary grouping variable (e.g., Group): ")
  
  if (!(response_col %in% names(df))) {
    stop("Selected grouping variable not found in dataset.")
  }
  
  # Only keep predictors starting with "area_"
  predictor_vars <- grep("^area_", names(df), value = TRUE)
  
  # Initialize results container
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
  
  response_raw <- df[[response_col]]
  classes <- unique(response_raw)
  n_classes <- length(classes)
  
  if (n_classes < 2) {
    stop("Not enough classes in selected grouping variable.")
  }
  
  # Header
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
  
  for (predictor_col in predictor_vars) {
    predictor <- df[[predictor_col]]
    if (!is.numeric(predictor)) next
    
    if (n_classes == 2) {
      # Standard binary ROC
      response <- as.numeric(as.character(response_raw))
      
      roc_obj <- pROC::roc(response = response, predictor = predictor)
      sens <- roc_obj$sensitivities
      spec <- roc_obj$specificities
      thr <- roc_obj$thresholds
      youden <- sens + spec - 1
      max_idx <- which.max(youden)
      
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
      # Multi-class: one-vs-rest
      for (cls in classes) {
        response <- ifelse(response_raw == cls, 1, 0)
        
        roc_obj <- pROC::roc(response = response, predictor = predictor)
        sens <- roc_obj$sensitivities
        spec <- roc_obj$specificities
        thr <- roc_obj$thresholds
        youden <- sens + spec - 1
        max_idx <- which.max(youden)
        
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
write_xlsx(all_results, file_name)
cat("Multi-class ROC analysis complete and saved to:", file_name, "\n")
}


ConfounderEffectAnalysis <- function(full_data_path, top_markers_path, output_file = "Confounder_Effect_Analysis.xlsx") {
  # Load datasets
  cat("Loading data...\n")
  full_data <- read_excel(full_data_path)
  top_markers <- read_excel(top_markers_path)
  selected_areas <- unique(top_markers$Area)
  
  # Identify grouping variables
  grouping_vars <- names(full_data)[grepl("_bin$", names(full_data)) | sapply(full_data, is.factor)]
  if (length(grouping_vars) == 0) {
    stop("No grouping variables ending in '_bin' or factors found.")
  }
  
  cat(length(selected_areas), "selected areas and", length(grouping_vars), "grouping variables.\n")
  
  all_results <- list()  # list of data.frames by group
  
  for (group in grouping_vars) {
    group_results <- data.frame(
      Area = character(),
      Group = character(),
      Test = character(),
      P_Value = numeric(),
      Significant = character(),
      Normal_Distribution = character(),
      stringsAsFactors = FALSE
    )
    
    group_vals <- full_data[[group]]
    if (length(unique(na.omit(group_vals))) < 2) next
    
    for (area in selected_areas) {
      if (!(area %in% names(full_data))) {
        warning(paste("Skipping", area, "- not found in full dataset"))
        next
      }
      
      predictor <- full_data[[area]]
      if (!is.numeric(predictor)) next
      
      df_subset <- na.omit(data.frame(group = group_vals, predictor = predictor))
      if (nrow(df_subset) < 6) next
      
      group_levels <- unique(df_subset$group)
      
      # Normality test per group
      normality <- sapply(group_levels, function(g) {
        vals <- df_subset$predictor[df_subset$group == g]
        if (length(vals) < 3) return(NA)
        shapiro.test(vals)$p.value > 0.05
      })
      
      all_normal <- all(normality == TRUE, na.rm = TRUE)
      test_type <- NA
      p_val <- NA
      
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
      
      group_results <- rbind(group_results, data.frame(
        Area = area,
        Group = group,
        Test = test_type,
        P_Value = round(p_val, 5),
        Significant = ifelse(!is.na(p_val) & p_val < 0.05, "Yes", "No"),
        Normal_Distribution = ifelse(all_normal, "Yes", "No"),
        stringsAsFactors = FALSE
      ))
      
      cat(area, "vs", group, "→", test_type, "| p =", round(p_val, 4), "\n")
    }
    
    # Store results per group (sheet)
    if (nrow(group_results) > 0) {
      all_results[[group]] <- group_results
    }
  }
  
  # Write each group to its own Excel sheet
  write_xlsx(all_results, output_file)
  cat("\n Confounder effect analysis complete. Results saved to:", output_file, "\n")
}

