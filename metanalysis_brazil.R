library(readxl)
library(metafor)
library(dplyr)
library(writexl)
library(ggplot2)

# ---------------------------
# 1. Read data table
# ---------------------------
df <- read_excel(
  "C:\\path to\\Brasil.xlsx",
  sheet = "Brasil_20",
  na = c("", "NA")
)

# ---------------------------
# 2. Selected moderators
# ---------------------------
mods_cont_selected <- c("Temperature", "Rainfall", "Lat")  
mods_cat_selected  <- c()      

# ---------------------------
# 3. Standardize moderators
# ---------------------------
for (var in mods_cont_selected) {
  colname <- grep(paste0("^", var, "$"), colnames(df), ignore.case = TRUE, value = TRUE)
  if (length(colname) == 1) {
    df[[paste0(var, "_z")]] <- as.numeric(scale(df[[colname]]))
  }
}

# ---------------------------
# 4. Calculate effect sizes
# ---------------------------
Cer10 <- escalc(
  measure = "MD",
  m1i = me, sd1i = sde, n1i = ne,
  m2i = mc, sd2i = sdc, n2i = nc,
  data = df
)

# ---------------------------
# 5. Formula for moderators
# ---------------------------
mods_use <- c(
  paste0(mods_cont_selected, "_z")[paste0(mods_cont_selected, "_z") %in% colnames(df)],
  mods_cat_selected[mods_cat_selected %in% colnames(df)]
)
if(length(mods_use) == 0) stop("No moderators found in the dataframe.")
mods_formula <- as.formula(paste("~", paste(mods_use, collapse = " + ")))

# ---------------------------
# 6. Models
# ---------------------------
Cer10Null <- rma(yi, vi, data = Cer10, method = "REML", na.action = na.omit)
Cer10M    <- rma(yi, vi, mods = mods_formula, data = Cer10, method = "REML", na.action = na.omit)
Cer10FE   <- rma(yi, vi, data = Cer10, method = "FE", na.action = na.omit)

# ---------------------------
# 6b. PET and PEESE estimators
# ---------------------------
pet_res   <- try(rma(yi, vi, mods = ~sei, data = Cer10, method="FE"), silent=TRUE)
peese_res <- try(rma(yi, vi, mods = ~vi, data = Cer10, method="FE"), silent=TRUE)

get_pet_peese <- function(rma_obj, estimator_name) {
  if (inherits(rma_obj, "try-error")) {
    return(data.frame(Estimator = estimator_name, Term=NA, Estimate=NA, SE=NA, Zval=NA, Pval=NA, CI.lb=NA, CI.ub=NA))
  } else {
    res <- get_rma_results(rma_obj)
    res$Estimator <- estimator_name
    res <- res[, c("Estimator","Term","Estimate","SE","Zval","Pval","CI.lb","CI.ub")]
    return(res)
  }
}

pet_df   <- get_pet_peese(pet_res, "PET")
peese_df <- get_pet_peese(peese_res, "PEESE")

# ---------------------------
# 7. Heterogeneity (tau², I², Q, R²%)
# ---------------------------
get_heterogeneity <- function(rma_obj, model_name, tau2_null = NA) {
  R2 <- if (!is.na(tau2_null) && tau2_null > 0) {
    100 * (tau2_null - rma_obj$tau2) / tau2_null
  } else {
    NA
  }
  
  interpretation <- if (!is.na(R2)) {
    paste0("The moderators explained approximately ", round(R2, 1), "% of the heterogeneity.")
  } else {
    "Not applicable for this model."
  }
  
  data.frame(
    Model = model_name,
    tau2  = rma_obj$tau2,
    I2    = rma_obj$I2,
    H2    = rma_obj$H2,
    QE    = rma_obj$QE,
    QEp   = rma_obj$QEp,
    R2_percent = R2,
    Interpretation = interpretation
  )
}

hetero_null <- get_heterogeneity(Cer10Null, "Null")
hetero_mod  <- get_heterogeneity(Cer10M, "Moderators", tau2_null = Cer10Null$tau2)
hetero_fe   <- get_heterogeneity(Cer10FE, "Fixed Effect")

hetero_df <- rbind(hetero_null, hetero_mod, hetero_fe)

# ---------------------------
# 8. Egger’s test (publication bias)
# ---------------------------
egger_test <- try(regtest(Cer10Null, model = "rma", predictor = "sei"), silent = TRUE)
if (inherits(egger_test, "try-error")) {
  egger_res <- data.frame(Test = "Egger", Intercept = NA, Zval = NA, Pval = NA)
} else {
  if (!is.null(egger_test$intercept)) {
    egger_res <- data.frame(Test = "Egger", Intercept = egger_test$intercept, Zval = egger_test$zval, Pval = egger_test$pval)
  } else {
    est <- if(!is.null(egger_test$coef)) egger_test$coef else NA
    p   <- if(!is.null(egger_test$pval)) egger_test$pval else NA
    egger_res <- data.frame(Test = "Egger", Intercept = est, Zval = NA, Pval = p)
  }
}

# ---------------------------
# 9. Fail-safe N
# ---------------------------
valid_idx <- complete.cases(Cer10$yi, Cer10$vi)
fsn_res <- try(fsn(Cer10$yi[valid_idx], Cer10$vi[valid_idx], type = "Rosenthal", alpha = 0.05), silent = TRUE)
if (inherits(fsn_res, "try-error") || is.null(fsn_res$fsn.r)) {
  fsn_df <- data.frame(Observed_Studies = sum(valid_idx), FailSafeN = NA, Observed_P = NA,
                       Target_P = 0.05, Type = "Rosenthal", Interpretation = "Fail-safe N could not be computed.")
} else {
  interpretation_fsn <- paste0("Approximately ", format(fsn_res$fsn.r, big.mark = "."),
                               " null studies would be required to nullify the observed effect (p ≥ ", fsn_res$alpha.target, ").")
  fsn_df <- data.frame(Observed_Studies = fsn_res$k.obs, FailSafeN = fsn_res$fsn.r,
                       Observed_P = fsn_res$alpha.obs, Target_P = fsn_res$alpha.target,
                       Type = fsn_res$type, Interpretation = interpretation_fsn)
}

# ---------------------------
# 10. Trim-and-fill
# ---------------------------
trimfill_res <- try(trimfill(Cer10Null), silent=TRUE)
if (inherits(trimfill_res, "try-error")) {
  trimfill_df <- data.frame(Test="TrimFill", Estimate=NA, CI.lb=NA, CI.ub=NA, k0=NA)
} else {
  trimfill_df <- data.frame(Test="TrimFill", Estimate=trimfill_res$b,
                            CI.lb=trimfill_res$ci.lb, CI.ub=trimfill_res$ci.ub, k0=trimfill_res$k0)
}

# ---------------------------
# 11. Function to extract RMA results
# ---------------------------
get_rma_results <- function(rma_obj) {
  data.frame(Term = rownames(rma_obj$beta),
             Estimate = rma_obj$beta[,1],
             SE = rma_obj$se,
             Zval = rma_obj$zval,
             Pval = rma_obj$pval,
             CI.lb = rma_obj$ci.lb,
             CI.ub = rma_obj$ci.ub)
}

res_null_df <- get_rma_results(Cer10Null)
res_mod_df  <- get_rma_results(Cer10M)
res_fe_df   <- get_rma_results(Cer10FE)

# ---------------------------
# 12. Save Excel file
# ---------------------------
out_base <- "C:\\path to\\Brasil20F.xlsx"
dir.create(dirname(out_base), recursive = TRUE, showWarnings = FALSE)

write_xlsx(
  list(
    Null_Model = res_null_df,
    Moderator_Model = res_mod_df,
    Fixed_Effect_Model = res_fe_df,
    PET = pet_df,
    PEESE = peese_df,
    Heterogeneity = hetero_df,
    Egger_Test = egger_res,
    FailSafeN = fsn_df,
    TrimFill = trimfill_df
  ),
  path = out_base
)

# ---------------------------
# 13. Funnel plots in TIFF format (normal and contour-enhanced)
# ---------------------------
# Normal funnel plot
tiff(paste0(out_base, "_FunnelPlot_Normal.tif"), width=2000, height=1500, res=300)
funnel(Cer10Null, xlab="Mean Difference", ylab="Standard Error", cex=.8)
dev.off()

# Contour-enhanced funnel plot
tiff(paste0(out_base, "_FunnelPlot_Contour.tif"), width=2000, height=1500, res=300)
funnel(Cer10Null, xlab="Mean Difference", ylab="Standard Error", cex=.8,
       back="white", shade=c(0.9,0.95,0.99), legend=TRUE)
dev.off()

if (!inherits(trimfill_res, "try-error")) {
  # Trim-and-fill normal
  tiff(paste0(out_base, "_FunnelPlot_TrimFill_Normal.tif"), width=2000, height=1500, res=300)
  funnel(trimfill_res, xlab="Mean Difference", ylab="Standard Error", cex=.8)
  dev.off()
  
  # Trim-and-fill contour-enhanced
  tiff(paste0(out_base, "_FunnelPlot_TrimFill_Contour.tif"), width=2000, height=1500, res=300)
  funnel(trimfill_res, xlab="Mean Difference", ylab="Standard Error", cex=.8,
         back="white", shade=c(0.9,0.95,0.99), legend=TRUE)
  dev.off()
}

# ---------------------------
# 14. Leave-One-Out Analysis (with tau² and I²)
# ---------------------------
k_studies <- if(!is.null(Cer10Null$k)) Cer10Null$k else length(Cer10$yi)

loo_df <- data.frame(
  Study = character(0), 
  Estimate = numeric(0), 
  SE = numeric(0),
  Zval = numeric(0), 
  Pval = numeric(0), 
  CI.lb = numeric(0),
  CI.ub = numeric(0), 
  Tau2 = numeric(0),
  I2 = numeric(0),
  Diff_vs_Full = numeric(0),
  stringsAsFactors = FALSE
)

if (k_studies > 2) {
  for (i in seq_len(k_studies)) {
    data_loo <- Cer10[-i, ]
    res_loo <- try(rma(yi, vi, data = data_loo, method="REML"), silent=TRUE)
    if (!inherits(res_loo, "try-error")) {
      loo_df <- rbind(loo_df, data.frame(
        Study = if(!is.null(rownames(Cer10))) rownames(Cer10)[i] else paste0("Study_", i),
        Estimate = coef(res_loo),
        SE = res_loo$se,
        Zval = res_loo$zval,
        Pval = res_loo$pval,
        CI.lb = res_loo$ci.lb,
        CI.ub = res_loo$ci.ub,
        Tau2 = res_loo$tau2,
        I2 = res_loo$I2,
        Diff_vs_Full = coef(res_loo) - coef(Cer10Null)
      ))
    }
  }
}

# ---------------------------
# 15. Add Leave-One-Out results to Excel
# ---------------------------
write_xlsx(
  list(
    Null_Model = res_null_df,
    Moderator_Model = res_mod_df,
    Fixed_Effect_Model = res_fe_df,
    PET = pet_df,
    PEESE = peese_df,
    Heterogeneity = hetero_df,
    Egger_Test = egger_res,
    FailSafeN = fsn_df,
    TrimFill = trimfill_df,
    LeaveOneOut = if(nrow(loo_df)>0) loo_df else data.frame(Message = "Leave-one-out not available")
  ),
  path = out_base
)

# ---------------------------
# 16. Leave-One-Out plot (with tau²)
# ---------------------------
if (nrow(loo_df) > 0 && !all(is.na(loo_df$Estimate))) {
  
  tiff(paste0(out_base, "_LeaveOneOut.tif"), width=2500, height=1800, res=300)
  par(mar=c(5, max(4, 0.6 * nchar(max(as.character(loo_df$Study), na.rm=TRUE))), 4, 2) + 0.1)
  ypos <- seq_len(nrow(loo_df))
  
  # Main plot: estimated effect with CI
  plot(loo_df$Estimate, ypos, xlab = "Estimated effect (leave-one-out)", ylab = "",
       pch = 19, cex = 0.8, xlim = range(c(loo_df$CI.lb, loo_df$CI.ub, coef(Cer10Null)), na.rm=TRUE))
  for (i in seq_len(nrow(loo_df))) segments(loo_df$CI.lb[i], ypos[i], loo_df$CI.ub[i], ypos[i])
  
  # Add tau² bars (on a secondary axis)
  par(new=TRUE)
  plot(loo_df$Tau2, ypos, type="h", axes=FALSE, xlab="", ylab="", col="grey60", lwd=5, xlim=range(c(loo_df$Tau2, Cer10Null$tau2), na.rm=TRUE))
  
  axis(2, at = ypos, labels = loo_df$Study, las=1, cex.axis=0.8)
  axis(3, col="grey60", col.axis="grey40")
  mtext("Tau²", side=3, line=2, col="grey40")
  
  abline(v = coef(Cer10Null), col = "red", lwd = 2, lty = 2)
  title("Leave-one-out sensitivity analysis (with Tau²)")
  dev.off()
}

# ---------------------------
# 17. Histogram of percentage difference (linear vs trim-and-fill)
# ---------------------------
library(ggplot2)

if (!inherits(trimfill_res, "try-error") && !is.null(trimfill_res$yi.filled) && length(trimfill_res$yi.filled) > 0) {
  
  df_diff <- data.frame(
    linear = Cer10Null$yi,
    trim_fill = trimfill_res$yi.filled[1:length(Cer10Null$yi)]
  )
  
  df_diff <- df_diff %>%
    mutate(percent_diff = 100 * (linear - trim_fill) / trim_fill)
  
  tiff(paste0(out_base, "_PercentDiff_Linear_TrimFill.tif"), width=2000, height=1500, res=300)
  ggplot(df_diff, aes(x = percent_diff)) +
    geom_histogram(binwidth = 1, fill = "gray40", color = "black") +
    labs(
      x = "Percent difference between linear and trim-and-fill",
      y = "Count",
      title = "Comparison between linear and trim-and-fill estimates"
    ) +
    theme_minimal()
  dev.off()
  
  # ---------------------------
  # 18. Statistical summary of percentage difference
  # ---------------------------
  summary_diff <- data.frame(
    Mean = mean(df_diff$percent_diff, na.rm=TRUE),
    SD = sd(df_diff$percent_diff, na.rm=TRUE),
    Min = min(df_diff$percent_diff, na.rm=TRUE),
    Max = max(df_diff$percent_diff, na.rm=TRUE),
    Median = median(df_diff$percent_diff, na.rm=TRUE)
  )
  
} else {
  message("Trim-and-fill did not add studies; histogram and summary will not be generated.")
  summary_diff <- data.frame(Message="Trim-and-fill did not add studies; PercentDiff not available")
}

# ---------------------------
# 19. Update Excel file with summary table
# ---------------------------
write_xlsx(
  list(
    Null_Model = res_null_df,
    Moderator_Model = res_mod_df,
    Fixed_Effect_Model = res_fe_df,
    PET = pet_df,
    PEESE = peese_df,
    Heterogeneity = hetero_df,
    Egger_Test = egger_res,
    FailSafeN = fsn_df,
    TrimFill = trimfill_df,
    LeaveOneOut = if(nrow(loo_df)>0) loo_df else data.frame(Message = "Leave-one-out not available"),
    PercentDiff_Summary = summary_diff
  ),
  path = out_base  
)

