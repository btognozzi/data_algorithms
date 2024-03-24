library(tidyverse)
library(openxlsx)

# Function to read CSV
read_csv_with_skip <- function(file_path, skip_lines) {
  df <- read.csv(file_path, skip = skip_lines, header = TRUE,
                 na.string=c("","null","NaN","X"), stringsAsFactors = FALSE)
  return(df)
}

# Function to check control values
check_control <- function(df, row_index, lower_limit, upper_limit, message) {
  control_row <- df[row_index, ]
  if (is.na(control_row[, 5])) {
    return(paste(message, "has passed"))
  } else if (control_row[, 5] >= lower_limit && control_row[, 5] <= upper_limit) {
    return(paste(message, "has passed"))
  } else {
    return(paste("Review ", message, ", may need Medical Director approval", sep = ''))
  }
}

# Function to check NTC
check_NTC <- function(df, df1) {
  df_last_row <- df[nrow(df), 5]
  df1_last_row <- df1[nrow(df1), 5]
  if (all(is.na(df_last_row) && is.na(df1_last_row))) {
    return("NTC has passed")
  } else {
    return("NTC has not passed")
  }
}

# Reading in the FAM CSV
file <- readline(prompt = "Enter File Path for FAM CSV:")
df <- read_csv_with_skip(file, 25)

# Subset data
FAM_df <- subset(df, select = c('No.', 'Color', 'Name', 'Type', 'Ct'))
wt_df <- FAM_df %>% filter(FAM_df$No. <= 36)
mut_df <- FAM_df %>% filter(FAM_df$No. >= 37)

# Creating the standard dfs
wt_std <- wt_df %>% filter(No. <= 4)
mut_std <- mut_df %>% filter(No. <= 40)
std_conc <- log(c(50, 500, 5000, 50000), base = 10)

wt_std$log_CN <- std_conc
mut_std$log_CN <- std_conc

# Creating Linear Models for the wt/mut
lm_wt <- lm(Ct ~ log_CN, data = wt_std)
lm_mut <- lm(Ct ~ log_CN, data = mut_std)

# Add the linear regression formula and R-squared value to the plot - wt
wt_formula_text <- as.character(round(coef(lm_wt)[1], 3))
wt_formula_text <- paste(wt_formula_text, "+", as.character(coef(lm_wt)[2]), '* x')
wt_formula_text <- paste("y =", wt_formula_text)

# Calculate R-squared value - wt
wt_r_squared <- summary(lm_wt)$r.squared
wt_r_squared_text <- paste("R^2 =", round(wt_r_squared, 4))


# Add the linear regression formula and R-squared value to the plot - mut
mut_formula_text <- as.character(round(coef(lm_mut)[1], 3))
mut_formula_text <- paste(mut_formula_text, "+", as.character(coef(lm_mut)[2]), '* x')
mut_formula_text <- paste("y =", mut_formula_text)

# Calculate R-squared value - mut
mut_r_squared <- summary(lm_mut)$r.squared
mut_r_squared_text <- paste("R^2 =", round(mut_r_squared, 4))

# Making Graphs 
wt_graph <- ggplot(wt_std, aes(x = log_CN, y = Ct)) + ggtitle("JAK2 WT Standard Graph") + geom_point() + 
                xlim(1, 5) + ylim(22, 38) + geom_smooth(method = "lm", se = FALSE, formula = y ~ x) + 
                annotate("text", x = 3.75, y = 32, label= wt_formula_text) + 
                annotate("text", x = 3.75, y = 31.4, label = wt_r_squared_text)

mut_graph <- ggplot(mut_std, aes(x = log_CN, y = Ct)) + ggtitle("JAK2 MUT Standard Graph") + geom_point() + 
                xlim(1, 5) + ylim(22, 38) + geom_smooth(method = "lm", se = FALSE, formula = y ~ x, color = "red") + 
                annotate("text", x = 3.75, y = 32, label= mut_formula_text) + 
                annotate("text", x = 3.75, y = 31.4, label = mut_r_squared_text)

# Setting up Slope and Intercept variables
wt_intercept <- coef(lm_wt)[1]
wt_slope <- coef(lm_wt)[2]

mut_intercept <- coef(lm_mut)[1]
mut_slope <- coef(lm_mut)[2]

# Calculating log_CN values for the samples
wt_df$log_CN <- ifelse(wt_df$No. >= 5, (wt_df$Ct - wt_intercept) / wt_slope, NA)
mut_df$log_CN <- ifelse(mut_df$No. >= 41, (mut_df$Ct - mut_intercept) / mut_slope, NA)

# Calculating the number of copies
wt_df$WT_Copies <- ifelse(!is.na(wt_df$log_CN), round(10^(wt_df$log_CN)), 0)
mut_df$MUT_Copies <- ifelse(!is.na(mut_df$log_CN), round(10^(mut_df$log_CN)), 0)

# Combining the copy numbers w/ Sample ID names
final_df <- cbind(wt_df$Name, as.integer(wt_df$WT_Copies), as.integer(mut_df$MUT_Copies))
colnames(final_df) <- c("Sample ID", "WT_Copies", "MUT_Copies")
final_df <- final_df[-(1:4),]

# Converting final_df to integers for the copies
final_df <- data.frame(final_df, stringsAsFactors = FALSE)
final_df$WT_Copies <- as.integer(final_df$WT_Copies)
final_df$MUT_Copies <- as.integer(final_df$MUT_Copies)

#Creating the total copies column
final_df$Total_Copies <- final_df$WT_Copies + final_df$MUT_Copies

#Creating the % JAK2 column
final_df$JAK2_Percent <- ifelse(is.na(wt_df[5:nrow(wt_df), 5]) | is.na(mut_df[5:nrow(mut_df), 5]), 0, 
                                 round((final_df$MUT_Copies / final_df$Total_Copies) * 100, 3))

# Reading in the HEX CSV for IC Ct data
file2 <- readline(prompt = 'Enter File Path For HEX CSV:')

df1 <- read_csv_with_skip(file2, 25)

# Subset HEX data
HEX_df <- subset(df1, select = c('No.', 'Color', 'Name', 'Type', 'Ct'))
hex <- subset(HEX_df, select = c('No.', 'Ct'))
wt_hex <- hex %>% filter(hex$No. <= 36)
mut_hex <- hex %>% filter(hex$No. >= 37)

# Merging and renaming IC Ct and Ct columns
wt_result_df <- merge(x = wt_df, y = wt_hex, by = 'No.')
wt_result_df <- rename(wt_result_df, IC_Ct = Ct.y, Ct = Ct.x)

mut_result_df <- merge(x = mut_df, y = mut_hex, by = 'No.')
mut_result_df <- rename(mut_result_df, IC_Ct = Ct.y, Ct = Ct.x)

# Create and save Excel file
wb <- createWorkbook()
sheets <- c("FAM Data", "HEX Data", "WT Results", "MUT Results", "WT Standard", "MUT Standard", "JAK2 Results")
for (sheet_name in sheets) {
  addWorksheet(wb, sheet_name)
}

# Write data to the first sheet
writeDataTable(wb, sheet = "FAM Data", x = df)

# Write data to the second sheet
writeDataTable(wb, sheet = "HEX Data", x = df1)

# Write data to the third sheet
writeDataTable(wb, sheet = "WT Results", x = wt_result_df)

# Write data to the fourth sheet
writeDataTable(wb, sheet = "MUT Results", x = mut_result_df)

# Write data to the fifth sheet
writeDataTable(wb, sheet = "WT Standard", x = wt_std)

# Write data to the sixth sheet
writeDataTable(wb, sheet = "MUT Standard", x = mut_std)

# Write data to the seventh sheet
writeDataTable(wb, sheet = "JAK2 Results", x = final_df)

# Save the workbook to a file
save_file_path <- readline(prompt = 'Enter File Path For Output Excel File:')
saveWorkbook(wb, save_file_path)

# Export graphs as a pdf
save_wt_graph_path <- readline(prompt = 'Enter File Path For WT Graph:')
save_mut_graph_path <- readline(prompt = 'Enter File Path For MUT Graph:')
pdf(save_wt_graph_path)
plot(wt_graph)
dev.off()

pdf(save_mut_graph_path)
plot(mut_graph)
dev.off()

# Checking slopes and R-squared values
if (-3.81 <= wt_slope && wt_slope <= -3.07 && 
    -3.81 <= mut_slope && mut_slope <= -3.07 &&
    wt_r_squared >= 0.98 && mut_r_squared >= 0.98) {
  print("Slopes and R-squared values are valid")
} else {
  print("Slopes and/or R-squared values did not pass")
}

# Checking over internal controls
invalid_wt_indices <- which(!(25 <= wt_result_df$IC_Ct & wt_result_df$IC_Ct <= 37.79))
invalid_mut_indices <- which(!(25 <= mut_result_df$IC_Ct & mut_result_df$IC_Ct <= 37.79))

if (length(invalid_wt_indices) == 0 && length(invalid_mut_indices) == 0) {
  print("Internal Controls passed")
} else {
  print("Internal Controls not passed")
}

# Check control values - CHANGE DFs/modify the control idexes
hpc_low_limit <- readline(prompt = 'Enter Lower Limit for HPC:')
hpc_high_limit <- readline(prompt = 'Enter Upper Limit for HPC:')
lpc_low_limit <- readline(prompt = 'Enter Lower Limit for LPC:')
lpc_high_limit <- readline(prompt = 'Enter Upper Limit for LPC:')

print(check_control(mut_result_df, nrow(mut_result_df) - 3, as.double(hpc_low_limit), as.double(hpc_high_limit), "HPC"))
print(check_control(mut_result_df, nrow(mut_result_df) - 2, as.double(lpc_low_limit), as.double(lpc_high_limit), "LPC"))
print(check_control(mut_result_df, nrow(mut_result_df) - 1, 0, 0, "NC"))
print(check_NTC(mut_result_df, wt_result_df))

# Define a function to interpret sample results
interpret_sample <- function(sample_id, jak2_percent) {
  if (jak2_percent >= 0.8) {
    cat(sample_id, "is positive and has a JAK2 percentage of", jak2_percent, "\n")
  } else if (jak2_percent >= 0.1) {
    cat(sample_id, "is indeterminate and has a JAK2 percentage of", jak2_percent, "\n")
  } else if (jak2_percent >= 0) {
    cat(sample_id, "is not detected\n")
  } else {
    cat(sample_id, "is invalid\n")
  }
}

# Iterate through rows excluding control rows
for (i in 1:(nrow(final_df))) {
  sample_id <- final_df$Sample.ID[i]
  jak2_percent <- final_df$JAK2_Percent[i]
  interpret_sample(sample_id, jak2_percent)
}