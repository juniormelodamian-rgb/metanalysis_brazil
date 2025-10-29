# ===========================================================
# Packages
# ===========================================================
required_pkgs <- c("readxl","writexl","metafor","dplyr","rlang")
new_pkgs <- required_pkgs[!(required_pkgs %in% installed.packages()[,"Package"])]
if(length(new_pkgs)) install.packages(new_pkgs)

library(readxl)
library(writexl)
library(metafor)
library(dplyr)
library(rlang)

# ===========================================================
# Helper Functions
# ===========================================================
parse_depth <- function(depth_str){
  s <- gsub("cm|\\s", "", depth_str)
  s <- gsub("–|—|−", "-", s)
  parts <- unlist(strsplit(s, "-"))
  if(length(parts) < 2) return(c(NA, NA))
  as.numeric(parts[1:2])
}

calc_lnRR <- function(m_t, sd_t, n_t, m_c, sd_c, n_c, zero_adj = 1e-6){
  m_t[m_t == 0] <- zero_adj
  m_c[m_c == 0] <- zero_adj
  v <- (sd_t^2)/(n_t * m_t^2) + (sd_c^2)/(n_c * m_c^2)
  yi <- log(m_t / m_c)
  list(yi = yi, vi = v)
}

prepare_es <- function(dat,
                       col_m_t = "treatment_mean",
                       col_sd_t = "treatment_sd",
                       col_n_t = "treatment_n",
                       col_m_c = "control_mean",
                       col_sd_c = "control_sd",
                       col_n_c = "control_n",
                       col_depth = "depth",
                       col_study = "study_id"){
  
  needed <- c(col_m_t, col_sd_t, col_n_t, col_m_c, col_sd_c, col_n_c, col_depth, col_study)
  missing <- setdiff(needed, names(dat))
  if(length(missing) > 0)
    stop("Missing columns: ", paste(missing, collapse=", "))
  
  dat2 <- dat %>%
    mutate(
      depth_raw = !!sym(col_depth),
      study_id  = !!sym(col_study)
    ) %>%
    rowwise() %>%
    mutate(
      depth_start = parse_depth(depth_raw)[1],
      depth_end   = parse_depth(depth_raw)[2]
    ) %>% ungroup()
  
  ln <- calc_lnRR(dat2[[col_m_t]], dat2[[col_sd_t]], dat2[[col_n_t]],
                  dat2[[col_m_c]], dat2[[col_sd_c]], dat2[[col_n_c]])
  
  dat2$yi <- ln$yi
  dat2$vi <- ln$vi
  dat2
}

lnRR_to_perc <- function(estimate, ci_lb, ci_ub){
  c(PercChange=(exp(estimate)-1)*100,
    PercChange_CI_lb=(exp(ci_lb)-1)*100,
    PercChange_CI_ub=(exp(ci_ub)-1)*100)
}

# ===========================================================
# 1) Read Input Sheet
# ===========================================================
dat <- read_excel("C:\\path to\\NV_grassland.xlsx",
                  sheet = "Degraded")

# ===========================================================
# 2) Compute lnRR and Prepare Dataset
# ===========================================================
es <- prepare_es(dat)
es_clean <- es %>% filter(!is.na(study_id) & !is.na(depth_raw)) %>%
  mutate(study_id = factor(study_id),
         depth_raw = factor(depth_raw))

cat("Valid records:", nrow(es_clean), "\n")

# ===========================================================
# 3) Meta-analysis by Biome (intercept-only)
# ===========================================================
biome_levels <- unique(es_clean$Biome)
results_by_Biome <- list()

for(b in biome_levels){
  sub_es <- es_clean %>% filter(Biome == b)
  
  if(nrow(sub_es) > 1){
    res <- rma.mv(
      yi, vi,
      random = ~1 | study_id/depth_raw,
      data = sub_es,
      method = "REML"
    )
    results_by_Biome[[b]] <- res
  } else {
    message("Biome =", b, " skipped: insufficient studies (n <= 1)")
  }
}

# ===========================================================
# 4) Export Biome-level Intercept Results
# ===========================================================
biome_summary_list <- lapply(names(results_by_Biome), function(b){
  res <- results_by_Biome[[b]]
  df <- data.frame(
    Biome = b,
    term  = "(Intercept)",
    estimate = coef(res),
    se       = res$se,
    zval     = res$zval,
    pval     = res$pval,
    ci_lb    = res$ci.lb,
    ci_ub    = res$ci.ub
  )
  perc_df <- t(mapply(lnRR_to_perc, df$estimate, df$ci_lb, df$ci_ub))
  cbind(df, perc_df)
})

biome_summary_df <- do.call(rbind, biome_summary_list)

write_xlsx(list(results_by_Biome = biome_summary_df),
           path = "Pastdegraded.xlsx")

cat("\n>>> Biome-level intercept-only results saved to 'Pastdegraded.xlsx'\n")
