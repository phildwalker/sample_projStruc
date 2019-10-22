# Program to kick off the RMD file (creates the HTML log and creates the daily metrics data) -pw
# Tue Sep 24 10:18:17 2019 ------------------------------

# setwd("S:/TSR_CSR/.....xxx")

# Start Fresh Environment ------------
suppressWarnings(source("./Scripts/Clean_Environ.R"))

# getwd()
# setwd("S:/TSR_CSR/.....xxx")
# getwd()

# Load Libraries ----------------------------------------------------------

# For running HTML/RMD
suppressWarnings(suppressPackageStartupMessages(library(knitr)))
suppressWarnings(suppressPackageStartupMessages(library(rmarkdown)))
suppressWarnings(suppressPackageStartupMessages(library(here)))

# For data cleaning
suppressWarnings(suppressPackageStartupMessages(library(tidyverse)))

# Imports -----------------------------------------------------------------

# Holidays
# Pulls in holiday calendar, and creates the variables for "today" and "last business day"  
suppressWarnings(source("./Scripts/BusinessDays.R"))

# Emails script
suppressWarnings(source("./Scripts/Emails.R"))

# Write xlsx sript
suppressWarnings(source("./Scripts/writeXLSX.R"))

# Read in excel file which contains the maps/ look-ups
SkillList <- readxl::read_excel("./DataImports/Imports_Lookups.xlsx", sheet = "QC_List") %>% # List of Skills to QC
  filter(QC_Y_N == "Y") 
AssocList <- readxl::read_excel("./DataImports/Imports_Lookups.xlsx", sheet = "AssocList") # Associate list
QC_lookup <- readxl::read_excel("./DataImports/Imports_Lookups.xlsx", sheet = "QC_Perc") # Look up table for what to percentage to QC
Notify <- readxl::read_excel("./DataImports/Imports_Lookups.xlsx", sheet = "QC_Email") # Look up table who to send the report to

# Run Report (after initialize imports/libraries) -------------------------

OutputName <- paste0("NonFincQC_",Today,".html")

Sys.setenv(RSTUDIO_PANDOC="C:/Program Files/RStudio/bin/pandoc")
rmarkdown::render(input="./Scripts/BuildReport.Rmd",
                  output_file= here("Outputs", "NonFincQC.html"))

file.copy(from=here("Outputs", "NonFincQC.html"),
          to=here("Outputs", "html", OutputName),
          overwrite = TRUE)

