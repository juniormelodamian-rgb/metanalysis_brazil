Description

This repository contains the final R script used for the meta-analysis supporting the revision of the manuscript submitted to Nature Climate Change.
The analysis quantifies the impact of agricultural practices and environmental gradients on soil organic carbon (SOC) across Brazilian regions.

All models, diagnostics, and figures were generated using R in a fully reproducible workflow.



Repository structure
metanalysis_brazil/
│
├── metanalysis_brazil.R       # Main analysis script
├── Brasil.xlsx                # Input dataset (not included if confidential)
├── Brasil20F.xlsx             # Output file with model results for the 0-20 cm. To others layers, please change the code
├── README.md                  # This documentation file
└── Figures and diagnostic plots (.tif) automatically generated



R package dependencies

The analysis requires the following CRAN packages:

library(readxl)
library(metafor)
library(dplyr)
library(writexl)
library(ggplot2)


Recommended R version: ≥ 4.3.0
All functions follow metafor syntax and are compatible with Windows and Linux systems.



Running the script

Adjust the path to the input Excel file in line 9:

df <- read_excel("C:/path/to/Brasil.xlsx", sheet = "Brasil_20")


Run the entire script (Ctrl + A → Ctrl + Enter in RStudio).

The script automatically generates:

Random-effects, fixed-effects, and moderator models (REML method).

PET and PEESE estimators for publication bias adjustment.

Heterogeneity metrics (τ², I², Q-test, and explained variance R²%).

Egger’s regression test, Rosenthal’s Fail-safe N, and Trim-and-Fill correction.

Leave-one-out sensitivity analysis and corresponding plots.

Percentage difference histogram between linear and trim-and-fill estimates.

All results are exported to:

Database.xlsx



Output files
Output	Description
Brasil20F.xlsx	Main Excel file containing all tables and model results.
Brasil20F.xlsx_FunnelPlot_Normal.tif	Standard funnel plot.
Brasil20F.xlsx_FunnelPlot_Contour.tif	Contour-enhanced funnel plot.
Brasil20F.xlsx_LeaveOneOut.tif	Leave-one-out analysis plot (with τ²).
Brasil20F.xlsx_PercentDiff_Linear_TrimFill.tif	Histogram of percent differences between linear and trim-and-fill results.



Methodological notes

Effect sizes were computed as Mean Differences (MD) between experimental and control groups.

Moderator variables (Temperature, Rainfall, Lat) were z-standardized before modeling.

The random-effects model was estimated using Restricted Maximum Likelihood (REML).


Heterogeneity, bias diagnostics, and effect precision follow the recommendations of
Borenstein et al. (2011) and Stanley & Doucouliagos (2012).

Reproducibility

All model outputs, heterogeneity measures, and diagnostic plots are generated automatically.
To reproduce the analysis, simply run the script after adjusting the data path.

For transparency, no manual post-processing was applied — all numbers and figures derive directly from the R workflow.

📄 License

MIT License — free use and reproduction are permitted with appropriate citation.
