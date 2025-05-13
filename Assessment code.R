# Install and load necessary packages
if (!require(ltm)) install.packages("ltm")
if (!require(openxlsx)) install.packages("openxlsx")
if (!require(dplyr)) install.packages("dplyr")
if (!require(tidyr)) install.packages("tidyr")
if (!require(tibble)) install.packages("tibble")  # for column_to_rownames

library(ltm)
library(openxlsx)
library(dplyr)
library(tidyr)
library(tibble)

# Load and clean your data
my_data <- read.csv("Simulated_Adaptive_Test_Data (2).csv")
my_data$section <- trimws(my_data$section)  # Remove leading/trailing spaces

# Choose the section
section_name <- "Grammar"

# Filter and select relevant columns
section_data <- my_data %>%
  filter(section == section_name & !is.na(score)) %>%
  dplyr::select(user_id, item_id, score)

# Reshape data to wide format
wide_data <- section_data %>%
  pivot_wider(names_from = item_id, values_from = score, values_fill = NA) %>%
  column_to_rownames("user_id")

# Convert to data.frame
wide_data <- as.data.frame(wide_data)

# Filter out users with all 0 or all 1
user_scores <- rowSums(wide_data, na.rm = TRUE)
num_items <- ncol(wide_data)
wide_data_clean <- wide_data[user_scores > 0 & user_scores < num_items, ]

# Remove items with only 1 response category (needed for Rasch model)
item_valid <- sapply(wide_data_clean, function(x) length(unique(na.omit(x))) > 1)
wide_data_clean <- wide_data_clean[, item_valid]

# Fit the Rasch model
rasch_model <- rasch(wide_data_clean)

# Extract person abilities and item difficulties
abilities <- factor.scores(rasch_model)$score.dat
item_difficulties <- coef(rasch_model)

# Prepare data frames
ability_df <- data.frame(
  UserID = rownames(wide_data_clean),
  Ability = round(abilities$z1, 3),
  Section = section_name
)

difficulty_df <- data.frame(
  ItemID = rownames(item_difficulties),
  Difficulty = round(item_difficulties[, "Dffclt"], 3),
  Section = section_name
)

# Write to Excel
write.xlsx(ability_df, "User_final_Abilities.xlsx")
write.xlsx(difficulty_df, "Item_final_Difficulties.xlsx")
