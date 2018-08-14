
# =========
# Retail4BA
# =========

# Load libraries
library(dplyr)
library(readxl)
library(ggplot2)

# ========================================
# Analysis-by-cohorts-vintage-example.xlsx
# ========================================

# ----------------------------------
# Store Size (ss)
# ----------------------------------

# Read in Store Size ('General info') worksheet
ss <- read_xlsx("xlsx_retail4BA/Analysis-by-cohorts-vintage-example.xlsx", sheet = "General info", skip = 2)

# Create valid column names for 'Store Size' table
h1 <- colnames(ss) %>%
  make.names()
names(ss) <- gsub("\\.\\.","\\.",h1)
rm(h1)

# ----------------------------------
# Sales by Store (sbs)
# ----------------------------------

# Read in 'Sales by Store' worksheet
sbs <- read_xlsx("xlsx_retail4BA/Analysis-by-cohorts-vintage-example.xlsx", sheet = "Sales by Store", skip = 2)

# Create valid column names for 'Sales by Store' table
h1 <- colnames(sbs) %>%
  make.names()
names(sbs) <- gsub("\\.\\.","\\.",h1)
rm(h1)

# ----------------------------------
# FootFall (ff)
# ----------------------------------

# Read in 'FootFall' worksheet
ff <- read_xlsx("xlsx_retail4BA/Analysis-by-cohorts-vintage-example.xlsx", sheet = "FootFall", skip = 2)

# ----------------------------------
# Year of opening (yoo)
# ----------------------------------

# Read in 'Year of opening' worksheet
yoo <- read_xlsx("xlsx_retail4BA/Analysis-by-cohorts-vintage-example.xlsx", sheet = "Year of opening", skip = 2)


# ----------------------------------
# Exploratory Analysis
# ----------------------------------

# Number of stores from 'Store Size' table
nrow(ss)
# [1] 88

# Store Numbers unique? yes
storeNum <- unique(ss$Store.Number.new)
length(storeNum)
# [1] 88   

# Sales by Store table describes same stores? yes
storeNum2 <- unique(sbs$Store.Number.new)
length(storeNum2)
# [1] 88
setdiff(storeNum,storeNum2)
# numeric(0)

# ----------------------------------
# Store #1132 in Breslau (s1132)
# ----------------------------------

# Get Store #1132 in Breslau sales information in 2010
s1132 <- sbs[(sbs$Store.Number.new == 1132 & sbs$Year == 2010),]

# Calculate Sales Density for Store #1132 in Euros 
round(s1132$Net.sales.In.EUR / s1132$Space.In.sq.m)
# [1] 289

# =============================================================
# =============================================================

# ----------------------------------------
# Category Space Totals (catSpaceTotals) 
# Analysis-by-cohorts-vintage-example.xlsx
# ----------------------------------------

# Get total square meter store space
# for 5 product categories
catSpaceTotals <- colSums(ss[,4:8], na.rm = T)
catSpaceTotals <- as_tibble(catSpaceTotals)

# Create tibble 'catSpaceTotals' with approp colnames
catName <- gsub("\\.In\\.sq\\.m","",rownames(catSpaceTotals))
catName <- as_tibble(catName)
catSpaceTotals <- cbind(catName, catSpaceTotals)
colnames(catSpaceTotals) <- c("Category","CatSpcTotal")
rm(catName)

# Plot category space totals
ggplot(catSpaceTotals, aes(x = Category, y = CatSpcTotal )) +
  geom_bar(stat = "identity") +
  ggtitle("Category Space Totals \n (all stores, in square meters)")

# ----------------------------------









