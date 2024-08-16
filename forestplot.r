# First, install and load required packages if you haven't already
# install.packages(c("ggplot2", "dplyr"))
library(ggplot2)
library(dplyr)
library(readxl)
library(forestplot)

setwd("~/Desktop/有事/Harvard/Research/PharmacoEpi/Korina_DPP4i/")

SGL2i <- read_xlsx("./2024-0624 SGLT2i//Excel doc/OR,RR_.xlsx")
GLP1 <- read_xlsx("./2024-0625 GLP1RA/Excel doc/OR,RR_.xlsx")

# Function to clean and prepare data
clean_data <- function(df) {
  df %>%
    filter(!is.na(Outcome) & Outcome != "NA" & !grepl("outcomes", Outcome)) %>%
    select(Outcome, `After PSM`, ...8, ...9, ...10) %>%
    rename(Events_1 = `After PSM`,
           Events_2 = ...8,
           RR = ...9,
           CI = ...10) %>%
    filter(!is.na(RR) & RR != "RR") %>%
    mutate(Lower = as.numeric(gsub("\\(|,.*", "", CI)),
           Upper = as.numeric(gsub(".*,|\\)", "", CI)),
           RR = as.numeric(RR))
}
SGL2i_clean <- clean_data(SGL2i)
GLP1_clean <- clean_data(GLP1)

# Function to create table text
create_tabletext <- function(data, title) {
  cbind(
    c(title, "Dental outcomes", data$Outcome),
    c("", "RR (95%CI)", sprintf("%.3f (%.3f, %.3f)", data$RR, data$Lower, data$Upper))
  )
}
GLP1_tabletext <- create_tabletext(GLP1_clean, "GLP1 vs DPP4i")
SGL2i_tabletext <- create_tabletext(SGL2i_clean, "SGLT2i vs DPP4i")

# Function to create forest plot
create_forestplot <- function(data, tabletext) {
  forestplot(
    labeltext = tabletext,
    mean = c(NA, NA, data$RR),
    lower = c(NA, NA, data$Lower),
    upper = c(NA, NA, data$Upper),
    is.summary = c(TRUE, TRUE, rep(FALSE, nrow(data))),
    xlab = "Risk Ratio (95% CI)",
    zero = 1,
    boxsize = 0.2,
    ci.vertices = TRUE,
    ci.vertices.height = 0.1,
    col = fpColors(box = "black", lines = "black", zero = "black"),
    xticks = c(0.25, 0.5, 0.71, 1.0, 1.41),
    txt_gp = fpTxtGp(label = gpar(fontface = 2),
                     ticks = gpar(cex = 0.8),
                     title = gpar(cex = 1),
                     xlab = gpar(cex = 0.8))
  )
}

# Create and save plots
pdf("forest_plot.pdf", width = 10, height = 8)
create_forestplot(GLP1_clean, GLP1_tabletext)
create_forestplot(SGL2i_clean, SGL2i_tabletext)
dev.off()