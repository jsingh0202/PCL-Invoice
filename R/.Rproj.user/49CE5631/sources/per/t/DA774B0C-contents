library(readxl)
library(openxlsx)
library(tcltk)

file <- file.choose()
sheet_names <- excel_sheets(file)

start <- which(tolower(trimws(sheet_names)) == "wr1")
end   <- which(tolower(trimws(sheet_names)) == "fixed fee")
print(sprintf("Starting at sheet: %d", start))
print(sprintf("Ending at sheet: %d", end))

if (length(start) == 0 || length(end) == 0) {
  stop("Could not find 'WR1' or 'FIXED FEE' sheet.")
}
if (start > end) {
  stop("'FIXED FEE' appears before 'WR1'.")
}

target_sheets <- sheet_names[start:end]
cat("Reading sheets:", paste(target_sheets, collapse = ", "), "\n")

sheet_data <- lapply(target_sheets, function(sheet) {
  df <- read_excel(file, sheet = sheet)
  
  # 1. Remove fully empty rows
  df <- df[!apply(df, 1, function(row) all(is.na(row) | trimws(row) == "")), ]
  
  # 2. Keep columns where ANY cell matches one of the patterns
  pattern <- "Description|Total Contract Value|% Complete|Total Progress to Date|Previously Billed|Current Billing|Balance"
  
  keep_cols <- sapply(df, function(col) {
    any(grepl(pattern, col, ignore.case = TRUE))
  })
  
  df <- df[, keep_cols, drop = FALSE]
  return(df)
})
names(sheet_data) <- target_sheets
# output_folder <- tk_choose.dir(caption = "Select output folder")
# if (is.na(output_folder)) stop("No folder selected.")
# output_file <- file.path(output_folder, "filtered_sheets.xlsx")
# output_folder <- tk_choose.dir(caption = "Select output folder")
if (!dir.exists("out")) {
  dir.create("out")
}
output_file <- file.path("out", "filtered_sheets.xlsx")

wb <- createWorkbook()
for (name in names(sheet_data)) {
  addWorksheet(wb, name)
  writeData(wb, name, sheet_data[[name]])
}
saveWorkbook(wb, output_file, overwrite = TRUE)
cat("Saved output to:", output_file, "\n")