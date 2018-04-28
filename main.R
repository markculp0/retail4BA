
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
ss <- read_xlsx("Analysis-by-cohorts-vintage-example.xlsx", sheet = "General info", skip = 2)

# Create valid column names for 'Store Size' table
h1 <- colnames(ss) %>%
  make.names()
names(ss) <- gsub("\\.\\.","\\.",h1)
rm(h1)

# ----------------------------------
# Sales by Store (sbs)
# ----------------------------------

# Read in 'Sales by Store' worksheet
sbs <- read_xlsx("Analysis-by-cohorts-vintage-example.xlsx", sheet = "Sales by Store", skip = 2)

# Create valid column names for 'Sales by Store' table
h1 <- colnames(sbs) %>%
  make.names()
names(sbs) <- gsub("\\.\\.","\\.",h1)
rm(h1)

# ----------------------------------
# FootFall (ff)
# ----------------------------------

# Read in 'FootFall' worksheet
ff <- read_xlsx("Analysis-by-cohorts-vintage-example.xlsx", sheet = "FootFall", skip = 2)

# ----------------------------------
# Year of opening (yoo)
# ----------------------------------

# Read in 'Year of opening' worksheet
yoo <- read_xlsx("Analysis-by-cohorts-vintage-example.xlsx", sheet = "Year of opening", skip = 2)

# ==================================
# Product-Range-Management.xlsx
# ==================================

# ----------------------------------
# Product Range Performance (prp)
# ----------------------------------

# Read in 'Product Range Performance worksheet
prp <- read_xlsx("Product-Range-Management.xlsx", sheet = "Basic Option", skip = 2)

# Create valid column names for 'Product Range 
# Performance' table
h1 <- colnames(prp) %>%
  make.names()
names(prp) <- gsub("\\.\\.","\\.",h1)
rm(h1)
prp <- rename(prp, Category = X__1)


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

# ==================================
# Product-Range-Management.xlsx
# ==================================

# ----------------------------------
# Category Space (catSpace) from
# Product-Range-Management.xlsx
# ----------------------------------

# Get category space in square meters
catSpace <- prp[1:6,1:2] %>%
  rename(Space = Space.In.sq.m)

# Plot category space totals
ggplot(catSpace, aes(x = Category, y = Space )) +
  geom_bar(stat = "identity") +
  ggtitle("Category Space Totals \n (all stores, square meters)")

# ----------------------------------
# Sales by Category (salesByCat) from
# Product-Range-Management.xlsx
# ----------------------------------

# Get sales by category 
#  (USD in thousands)
salesByCat <- prp[1:6,c(1,3)] %>%
  rename(Sales = Sales.generated.In.thousands.of.USD)

# Plot sales totals by category
ggplot(salesByCat, aes(x = Category, y = Sales )) +
  geom_bar(stat = "identity") +
  ggtitle("Sales by Category \n (all stores, USD in thousands)")

# ----------------------------------
# Sales Density (salesDens) 
# by product category 
# from Product-Range-Management.xlsx
# ----------------------------------

# Calculate sales density by category
salesDens <- salesByCat[,2] / catSpace[,2] * 1000
salesDens <-  cbind(salesByCat[,1], salesDens[,1]) %>%
  rename(Sales_Density = 'salesDens[, 1]')
  
# Plot sales density by category
ggplot(salesDens, aes(x = Category, y = Sales_Density )) +
  geom_bar(stat = "identity") +
  ggtitle("Sales Density by Category \n    (USD per sq meter)")

# ----------------------------------
# Sales Margin (margins) 
# by product category 
# from Product-Range-Management.xlsx
# ----------------------------------

# Get margin breakdown by category
margins <- prp[1:6,c(1,5)] %>%
  rename(Margins = X.Margin)

# Plot margins by category
ggplot(margins, aes(x = Category, y = Margins )) +
  geom_bar(stat = "identity") +
  ggtitle("Margins by Category \n      (percent)")

# ----------------------------------
# Margin Density (marginDens) 
# by product category 
# from Product-Range-Management.xlsx
# ----------------------------------

# Calculate margin density by product category 
marginDens <- salesDens[,2] * margins[,2] %>%
  rename(Margin_Density = Margins)

# Reformat as a tibble w 2 cols
marginDens <- as_tibble(cbind(salesDens[,1], marginDens[,1])) %>%
  rename(Category = V1, Margin_Density = V2)

# Plot margin density by category
ggplot(marginDens, aes(x = Category, y = Margin_Density )) +
  geom_bar(stat = "identity") +
  ggtitle("Margin Density by Category \n    (USD per sq meter)")

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









