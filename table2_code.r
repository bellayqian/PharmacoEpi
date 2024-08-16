# dependencies ------------------------------------------------------------
library(docxtractr)
library(janitor)
library(writexl)
library(dplyr)
library(tidyverse)
library(survival)

setwd("~/Desktop/有事/Harvard/Research/PharmacoEpi/Korina_DPP4i/2024-0625 GLP1RA/")
# read-in data ------------------------------------------------------------
worddoc_before <-
  read_docx("./Word doc/before psm/GLP1_vs_DPP4i_dental_outcomes_6.29update_202406291600.docx")

worddoc_after <-
  read_docx("./Word doc/after psm/GLP1_vs_DPP4i_dental_outcomes_6.29update_202406291617.docx")

# variable names ----------------------------------------------------------
results_table_before <- docx_extract_tbl(worddoc_before, 5, header = FALSE)
results_table_before[results_table_before == ''] <- NA

results_table_before <- results_table_before %>% remove_empty("rows")

for (i in 1:nrow(results_table_before)) {
  if (is.na(results_table_before[i, 'V1'])) {
    results_table_before[i, 'V1'] <- results_table_before[i-1, 'V1']
  }
}

results_table_export_variables <- results_table_before %>% 
  filter(V2 == "Risk analysis")

# -------------------------------------------------------------------------
## edit file path-before
#docx_describe_tbls(worddoc_before)

results_table_before <- docx_extract_tbl(worddoc_before, 8, header = FALSE)

results_table_before[results_table_before == ''] <- NA

results_table_before <- results_table_before %>% remove_empty("rows")

for (i in 1:nrow(results_table_before)) {
  if (is.na(results_table_before[i, 'V1'])) {
    results_table_before[i, 'V1'] <- results_table_before[i-1, 'V1']
  }
}
# remove mostly missing values
results_table_before <- 
  results_table_before %>% 
  slice(2:n())
?slice

for (i in 1:nrow(results_table_before)) {
  if (is.na(results_table_before[i, 'V2'])) {
    results_table_before[i, 'V2'] <- results_table_before[i-1, 'V2']
  }
}
## edit file path-after
#docx_describe_tbls(worddoc_after)

results_table_after <- docx_extract_tbl(worddoc_after, 8, header = FALSE)

results_table_after[results_table_after == ''] <- NA

results_table_after <- results_table_after %>% remove_empty("rows")

for (i in 1:nrow(results_table_after)) {
  if (is.na(results_table_after[i, 'V1'])) {
    results_table_after[i, 'V1'] <- results_table_after[i-1, 'V1']
  }
}
# remove mostly missing values
results_table_after <- 
  results_table_after %>% 
  slice(2:n())

for (i in 1:nrow(results_table_after)) {
  if (is.na(results_table_after[i, 'V2'])) {
    results_table_after[i, 'V2'] <- results_table_after[i-1, 'V2']
  }
}
# numbers in active/control arms ------------------------------------------
## before 
# active included number, events
results_table_export_active_before <- results_table_before %>% 
  filter(grepl("^Risk analysis", V2)) %>% 
  filter(V3 == 1)
# control included number, events
results_table_export_control_before <- results_table_before %>% 
  filter(grepl("^Risk analysis", V2)) %>% 
  filter(V3 == 2)
# excluded (active+control)
# Assume the data frame is called 'results_table' and the column containing the sentence is named 'V3'
results_table_export_before <- results_table_before %>% 
  # filter(str_detect(V2, "^Kaplan"))
  filter(V2 == "Kaplan - Meier survival analysis") %>% 
  filter(grepl('patients in Cohort 1', V3)) 


## after
# active included number, events
results_table_export_active_after <- results_table_after %>% 
  filter(grepl("^Risk analysis", V2)) %>% 
  filter(V3 == 1)
# control included number, events
results_table_export_control_after <- results_table_after %>% 
  filter(grepl("^Risk analysis", V2)) %>% 
  filter(V3 == 2)
# excluded (active+control)
# Assume the data frame is called 'results_table' and the column containing the sentence is named 'V3'
results_table_export_after <- results_table_after %>% 
  filter(V2 == "Kaplan - Meier survival analysis") %>% 
  filter(grepl('patients in Cohort 1', V3)) 


# kaplan-meier ------------------------------------------------------------
## before
# RD
results_table_export_rd_before <- results_table_before %>% 
  filter(grepl("^Risk analysis", V2)) %>%  
  filter(V3 == "Risk Difference")
# RR
results_table_export_rr_before <- results_table_before %>% 
  filter(grepl("^Risk analysis", V2)) %>%  
  filter(V3 == "Risk Ratio")
# OR
results_table_export_or_before <- results_table_before %>% 
  filter(grepl("^Risk analysis", V2)) %>%  
  filter(V3 == "Odds Ratio") 
# HR
results_table_export_risk_before <- results_table_before %>% 
  filter(grepl("^Kaplan", V2)) %>%  
  filter(V3 == "Hazard Ratio and Proportionality")
# log rank pvalue
# results_table_export_log_before <- results_table_before %>% 
#   filter(grepl("^Kaplan", V2)) %>%  
#   filter(V3 == "Log-Rank Test")

# RR pvalue
results_table_export_log_before <- results_table_before %>%
  filter(grepl("^Risk analysis", V2)) %>%  
  filter(V3 == "Risk Difference")

## after
# RD
results_table_export_rd_after <- results_table_after %>% 
  filter(grepl("^Risk analysis", V2)) %>%  
  filter(V3 == "Risk Difference")
# RR
results_table_export_rr_after <- results_table_after %>% 
  filter(grepl("^Risk analysis", V2)) %>%  
  filter(V3 == "Risk Ratio")
# OR
results_table_export_or_after <- results_table_after %>% 
  filter(grepl("^Risk analysis", V2)) %>%  
  filter(V3 == "Odds Ratio")
# HR
results_table_export_risk_after <- results_table_after %>% 
  filter(grepl("^Kaplan", V2)) %>%  
  filter(V3 == "Hazard Ratio and Proportionality")
# log rank pvalue
# results_table_export_log_after <- results_table_after %>% 
#   filter(grepl("^Kaplan", V2)) %>%  
#   filter(V3 == "Log-Rank Test")
# RR pvalue
results_table_export_log_after <- results_table_after %>%
  filter(grepl("^Risk analysis", V2)) %>%  
  filter(V3 == "Risk Difference")

# turn them all into one excel file_comparison
## before
results_control_event_string_before <- gsub(",", "", results_table_export_control_before$V6)
results_active_event_string_before <- gsub(",", "", results_table_export_active_before$V6)
results_event_total_before <- as.numeric(results_control_event_string_before) + as.numeric(results_active_event_string_before)
results_event_total_before <- as.vector(results_event_total_before)
#cohort1 <- as.vector(cohort1)
#cohort2 <- as.vector(cohort2)

# turn them all into one excel file_comparison
## after
# results_control_event_string_after <- gsub(",", "", results_table_export_control_after$V6)
# results_active_event_string_after <- gsub(",", "", results_table_export_active_after$V6)
# results_event_total_after <- as.numeric(results_control_event_string_after) + as.numeric(results_active_event_string_after)
# results_event_total_after <- as.vector(results_event_total_after)
# 
# df_all <- data.frame(results_table_export_active_before$V5, results_table_export_active_before$V6,
#                      results_table_export_control_before$V5, results_table_export_control_before$V6,
#                      results_table_export_rr_before$V4, results_table_export_rr_before$V5,
#                      results_table_export_risk_before$V4, results_table_export_risk_before$V5, results_table_export_log_before$V6,
#                      results_table_export_active_after$V5, results_table_export_active_after$V6,
#                      results_table_export_control_after$V5, results_table_export_control_after$V6,
#                      results_table_export_rr_after$V4, results_table_export_rr_after$V5,
#                      results_table_export_risk_after$V4, results_table_export_risk_after$V5, results_table_export_log_after$V6
# )
# colnames(df_all) <- c("Included", "Events", "Included",	"Events",
#                       "Risk Ratio",	"95% CI",	"Hazard Ratio",	"95% CI", "P-value",
#                       "Included", "Events", "Included",	"Events",
#                       "Risk Ratio",	"95% CI",	"Hazard Ratio",	"95% CI", "P-value"
# )

results_table_export_active_before$V1 <- gsub("^\\d+\\s+", "", results_table_export_active_before$V1)

pvalue_before <- round(p.adjust(results_table_export_log_before$V7, method = "BH"), 3)
pvalue_after <- round(p.adjust(results_table_export_log_after$V7, method = "BH"), 3)

e_val <- function(hr) {
  hr <- ifelse(hr < 1, 1 / hr, hr)
  e_value <- (hr + sqrt(hr^2 - 1)) / hr
  return(e_value)
}
evalue_before <- round(e_val(as.numeric(results_table_export_risk_before$V4)), 3)
evalue_after <- round(e_val(as.numeric(results_table_export_risk_after$V4)), 3)

df_all <- data.frame(results_table_export_active_before$V1, #Outcomes
                     results_table_export_active_before$V5, #activeIncluded_before
                     results_table_export_active_before$V6, #activeEvents_before
                     results_table_export_control_before$V5, #controlIncluded_before
                     results_table_export_control_before$V6, #controlEvents_before
                     results_table_export_rr_before$V4, #RiskRatio_before
                     results_table_export_rr_before$V5, #RR 95%CI
                     results_table_export_risk_before$V4, #HazardRatio_before
                     results_table_export_risk_before$V5, #HR 95%CI
                     results_table_export_log_before$V6, #P-val
                     pvalue_before, #Adjusted P-val
                     evalue_before, #E-val
                     results_table_export_active_after$V5, #activeIncluded_after
                     results_table_export_active_after$V6, #activeEvents_after
                     results_table_export_control_after$V5, #controlIncluded_after
                     results_table_export_control_after$V6, #controlEvents_after
                     results_table_export_rr_after$V4, #RiskRatio_after
                     results_table_export_rr_after$V5, #RR 95%CI
                     results_table_export_risk_after$V4, #HazardRatio_after
                     results_table_export_risk_after$V5, #HR 95%CI
                     results_table_export_log_after$V6, #P-val
                     pvalue_after, #Adjusted P-val
                     evalue_after #E-val
)
colnames(df_all) <- c("Outcomes",
                      "activeIncluded_before", "activeEvents_before", 
                      "controlIncluded_before", "controlEvents_before",
                      "RiskRatio_before", "CI_RR_before", 
                      "HazardRatio_before", "CI_HR_before", "pval1", "adjusted_pval1", 
                      "eval1",
                      "activeIncluded_after", "activeEvents_after", 
                      "controlIncluded_after", "controlEvents_after",
                      "RiskRatio_after", "CI_RR_after", 
                      "HazardRatio_after", "CI_HR_after", "pval2", "adjusted_pval2",
                      "eval2"
)

# Table Making ------------------------------------------------------------
library(readxl)
library(kableExtra)
df_processed <- df_all %>%
  rename(Outcome = Outcomes,
         'GLP1 Before PSM Events' = activeEvents_before,
         'DPP4i Before PSM Events' = controlEvents_before,
         'RR Before PSM' = RiskRatio_before,
         '95% RR CI Before PSM' = CI_RR_before,
         'HR Before PSM' = HazardRatio_before, 
         '95% HR CI Before PSM' = CI_HR_before,
         'P-value Before PSM' = pval1,
         'Adjusted P-value Before PSM' = adjusted_pval1,
         'E-value Before PSM' = eval1,
         'GLP1 After PSM Events' = activeEvents_after,
         'DPP4i After PSM Events' = controlEvents_after,
         'RR After PSM' = RiskRatio_after,
         '95% RR CI After PSM' = CI_RR_after, # Corrected to After PSM
         'HR After PSM' = HazardRatio_after,
         '95% HR CI After PSM' = CI_HR_after, # Corrected to After PSM
         'P-value After PSM' = pval2,
         'Adjusted P-value After PSM' = adjusted_pval2,
         'E-value After PSM' = eval2) %>%
  select(Outcome, 'GLP1 Before PSM Events', 'DPP4i Before PSM Events',
         'RR Before PSM', '95% RR CI Before PSM', 'HR Before PSM', '95% HR CI Before PSM', 
         'P-value Before PSM', 'Adjusted P-value Before PSM', 'E-value Before PSM',
         'GLP1 After PSM Events', 'DPP4i After PSM Events',
         'RR After PSM', '95% RR CI After PSM', 'HR After PSM', '95% HR CI After PSM', 
         'P-value After PSM', 'Adjusted P-value After PSM', 'E-value After PSM')

# export ------------------------------------------------------------------
# numbers, OR, RR_before
write_xlsx(df_processed,"../2024-0625 GLP1RA/Excel doc/OR,RR.xlsx")



