library(readxl)
library(tidyverse)
library(writexl)

setwd("~/Desktop/有事/Harvard/Research/PharmacoEpi/Korina_DPP4i/2024-0625 GLP1RA/Excel doc/")
# data input --------------------------------------------------------------
before <- readxl::read_excel("./before psm.xlsx")
after <- readxl::read_excel("./after psm.xlsx")

# before ------------------------------------------------------------------
colnames(before) <- c(
  'NA1',
  'exposure',
  'variable',
  'description',
  'mean_sd',
  'patients',
  'pct',
  'pvalue',
  'std_diff'
)
# another way: 
for (i in 3:nrow(before)) {
  if (is.na(before[i, 'variable'])) {
    before[i, 'variable'] <- before[i-1, 'variable']
  }
}

for (i in 3:nrow(before)) {
  if (is.na(before[i, 'description'])) {
    before[i, 'description'] <- before[i-1, 'description']
  }
}
# formatting patients and pct ---------------------------------------------
before <- before %>%
  mutate(std_diff = round(as.numeric(std_diff), 4))

before <- before %>%
  mutate(
    n_patients_and_pct = paste0(scales::comma_format()(as.numeric(patients)), " (", as.numeric(pct) * 100, "%)")
  )

before_with <- before %>% 
  filter(
    exposure == 1
  )
# write_xlsx(before_with,"/Users/huangpinjia/Desktop/before+with.xlsx")

before_without <- before %>% 
  filter(
    exposure == 2
  )
# write_xlsx(before_without,"/Users/huangpinjia/Desktop/before+without.xlsx")

# after -------------------------------------------------------------------
colnames(after) <- c(
  'NA1',
  'exposure',
  'variable',
  'description',
  'mean_sd',
  'patients',
  'pct',
  'pvalue',
  'std_diff'
)
# another way: 
for (i in 3:nrow(after)) {
  if (is.na(after[i, 'variable'])) {
    after[i, 'variable'] <- after[i-1, 'variable']
  }
}

for (i in 3:nrow(after)) {
  if (is.na(after[i, 'description'])) {
    after[i, 'description'] <- after[i-1, 'description']
  }
}

# formatting patients and pct ---------------------------------------------
after <- after %>%
  mutate(std_diff = round(as.numeric(std_diff), 4))

after <- after %>% 
  mutate(
    n_patients_and_pct = paste0(scales::comma_format()(as.numeric(patients)), " (", as.numeric(pct) * 100, "%)")
  )

after_with <- after %>% 
  filter(
    exposure == 1
  )

after_without <- after %>% 
  filter(
    exposure == 2
  )

# data output -------------------------------------------------------------
df <- data.frame(before_with$description,before_with$n_patients_and_pct, before_without$n_patients_and_pct, before_with$std_diff,
                 after_with$n_patients_and_pct, after_without$n_patients_and_pct, after_with$std_diff)

colnames(df) <- c("Description","GLP1 before", "DPP4i before", "std diff before", "GLP1 after",	"DPP4i after", "std diff after")

# export ------------------------------------------------------------------
# before
write_xlsx(df,"table1.xlsx")
