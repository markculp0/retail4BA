
# ========================
# Product Range Management
# Product-Range-Management.xlsx
# ========================

# Load libraries
library(dplyr)
library(readxl)
library(ggplot2)

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
# Category Space (catSpace) from
# Product-Range-Management.xlsx
# ----------------------------------

# Get category space in square meters
catSpace <- prp[1:6,1:2] %>%
  rename(Space = Space.In.sq.m)

# Get each categories % of current total space
catSpacePrcnt <- round(catSpace[,2] / sum(catSpace[,2]),2)

# Combine in tibble
catSpace <- as_tibble(cbind(catSpace,catSpacePrcnt[,1])) %>%
  rename(SpacePrcnt='catSpacePrcnt[, 1]')
rm(catSpacePrcnt)

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

# ----------------------------------
# Margin Contributed (marginContrib) 
# by product category 
# from Product-Range-Management.xlsx
# ----------------------------------

# Calculate each category's margin
# contributed from current sales
marginContrib <- salesByCat[,2] * margins[,2]

# Calculate each category's % of 
# current total contibuted margin
marginContribPcnt <- round(marginContrib[,1] / sum(marginContrib[,1]),2)

# Combine as one tibble
marginContrib <- as_tibble(cbind(margins[,1],marginContrib))
marginContrib <- as_tibble(cbind(marginContrib,marginContribPcnt)) 
colnames(marginContrib) <- c("Category","CurMargin","CurPercent")
rm(marginContribPcnt)

# Plot margins contributed from current sales by category
ggplot(marginContrib, aes(x = Category, y = CurMargin )) +
  geom_bar(stat = "identity") +
  ggtitle("Margins Contributed \n from Current Sales \n by Category (USD thousands)")

# ----------------------------------
# Compare Contributed Margins to 
# Existing Space Allocations
# (margin2space) by product category 
# from Product-Range-Management.xlsx
# ----------------------------------

# Compare % of contributed margins 
# to % of existing space allocated to category 
margin2space <- round(marginContrib[,3] / catSpace[,3],2)

# Show over performing and under performing
# space efficiencies by category
margin2SpaceOverUnder <- margin2space - 1

# Combine as tibble
margin2space <- as_tibble(cbind(prp[1:6,1], margin2space[,1])) %>%
  rename(MarginVsSpace = 'margin2space[, 1]')
margin2space <- as_tibble(cbind(margin2space,margin2SpaceOverUnder[,1])) %>%
  rename(OverUnder = 'margin2SpaceOverUnder[, 1]')
margin2space$Color <- ifelse(margin2space$OverUnder < 0, "negative", "positive")
rm(margin2SpaceOverUnder)

# Plot over/under performing category space allocations
ggplot(margin2space, aes(x = Category, y = OverUnder, fill = Color)) +
  geom_bar(stat = "identity") +
  scale_fill_manual(values=c(positive="skyblue",negative="pink")) +
  geom_text(aes(label=OverUnder)) +
  ggtitle("Over/Under Performing \n Product Space Allocations \n by Category (percent)")

# ----------------------------------
# Calculate Recommended
# Space Reallocation
# (newSpace) by product category 
# from Product-Range-Management.xlsx
# ----------------------------------

# Calculate new space allocation based on performance
newSpace <- catSpace$Space * margin2space$MarginVsSpace

# Rounding adjustments
newSpace[1] <- newSpace[1] + 1
newSpace[3] <- newSpace[3] - 0.5
newSpace[6] <- newSpace[6] + 2
sum(newSpace)

# Calc new space allocation percentages
newSpacePrct <- round(newSpace / 1250, 2)

# Combine as tibble
newSpace <- as_tibble(cbind(prp[1:6,1],newSpace))
newSpace <- as_tibble(cbind(newSpace, newSpacePrct))
colnames(newSpace) <- c("Category", "NewSpace", "NewSpacePrct")
rm(newSpacePrct)

# ----------------------------------
# Calculate Potential Margin
# Generated using new
# Space Reallocation
# (newMarginContrib) by product category 
# from Product-Range-Management.xlsx
# ----------------------------------

# Convert to double 
marginDens$Margin_Density <- as.double(marginDens$Margin_Density)

# Calc new potential margin generated
newMargin <- round((newSpace$NewSpace * marginDens[,2]) / 1000)
colnames(newMargin) <- "NewMargin"
