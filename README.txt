# Meta-analysis of Soil Carbon in Brazil

This repository contains the **R scripts** used for the meta-analysis supporting the revision of the manuscript submitted to *Nature Climate Change*.  
The analysis quantifies the impact of agricultural management practices and environmental gradients on **soil organic carbon (SOC)** across Brazilian regions.

All analyses, models, diagnostics, and figures were generated in R using a **fully reproducible workflow**.

---

## 📁 Repository Structure

```
metanalysis_brazil/
│
├── metanalysis_brazil.R         # Main meta-analysis script (Mean Difference models)
├── metanalysis_multilevel.R     # Multilevel meta-analysis (lnRR, Amazon case example)
├── Brasil.xlsx                  # Input dataset (not included if confidential)
├── Brasil20F.xlsx               # Example output file for the 0–20 cm layer
├── README.md                    # This documentation file
└── Figures (.tif)               # Funnel plots, forest plots, and diagnostics
```

---

## 📦 R Package Dependencies

Both scripts require the following CRAN packages:

```r
library(readxl)
library(writexl)
library(metafor)
library(dplyr)
library(ggplot2)
```

**Recommended R version:** ≥ 4.3.0  
All functions follow the `metafor` syntax and are fully compatible with Windows, macOS, and Linux systems.

---

## ⚙️ Running the Analyses

### 1️⃣ Mean Difference (MD) Model
Script: **`metanalysis_brazil.R`**

Adjust the input file path in line 9:
```r
df <- read_excel("C:/path/to/Brasil.xlsx", sheet = "Brasil_20")
```

Run the script in R or RStudio (Ctrl + A → Ctrl + Enter).  
The script automatically generates:

- Random-, fixed-, and moderator-effect models (REML method).  
- PET and PEESE bias estimators.  
- Heterogeneity metrics (τ², I², Q-test, R²%).  
- Egger’s regression, Fail-safe N, and Trim-and-Fill tests.  
- Leave-one-out sensitivity plots and bias diagnostics.  

Outputs are automatically saved as Excel and TIFF files:
```
Brasil20F.xlsx
Brasil20F.xlsx_FunnelPlot_Normal.tif
Brasil20F.xlsx_FunnelPlot_Contour.tif
Brasil20F.xlsx_LeaveOneOut.tif
Brasil20F.xlsx_PercentDiff_Linear_TrimFill.tif
```

---

### 2️⃣ Multilevel Model (lnRR)
Script: **`metanalysis_multilevel.R`**

This script applies a **multilevel meta-analysis** using log response ratios (lnRR)  
as effect sizes, with random effects structured as `study_id/depth`.

It automatically:
- Calculates lnRR and sampling variances.  
- Fits null, random, and moderator multilevel models.  
- Computes percentage changes (ΔSOC).  
- Evaluates heterogeneity (τ², I², R²).  
- Produces Trim-and-Fill and Fail-safe N summaries.

Example outputs:
```
Amazon_1_2f.xlsx
Atlantic_Forest_trimfill_failsafe_summary.xlsx
```

---

## 📚 Methodological Notes

- Effect sizes were computed as either **Mean Differences (MD)** or **log Response Ratios (lnRR)**.  
- Moderator variables (Temperature, Rainfall, Latitude) were z-standardized before modeling.  
- All random-effect models were estimated using the **Restricted Maximum Likelihood (REML)** method.  
- Heterogeneity, bias diagnostics, and precision follow **Borenstein et al. (2011)** and **Stanley & Doucouliagos (2012)**.

---

## 🔁 Reproducibility

All results, heterogeneity measures, and plots are generated directly by the scripts —  
no manual post-processing is performed.

To reproduce the analyses:
1. Clone this repository.  
2. Adjust the file paths in each script.  
3. Run the scripts in sequence.

---

## 📄 License

**License:** Free use and reproduction are permitted with appropriate citation.  
If used in a publication, please cite the corresponding article and this repository.
