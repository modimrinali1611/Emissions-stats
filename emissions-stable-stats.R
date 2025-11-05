library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)
library(purrr)

# Helper: find stable periods within a group for a specific species
find_stable_periods <- function(data, species_col, min_duration_hours = 3, percent_tolerance = 0.10) {
  data <- data %>% arrange(`Time-UTC`)
  total_hours <- as.numeric(difftime(max(data$`Time-UTC`), min(data$`Time-UTC`), units = "hours"))
  duration_threshold <- max(min_duration_hours, 0.10 * total_hours)
  
  stable_periods <- list()
  i <- 1
  while (i < nrow(data)) {
    j <- i
    while (j <= nrow(data)) {
      values <- data[[species_col]][i:j]
      
      if (any(is.na(values))) {
        j <- j + 1
        next
      }
      
      mean_val <- mean(values)
      if (is.na(mean_val) || mean_val == 0) {
        j <- j + 1
        next
      }
      
      if (all(abs(values - mean_val) / mean_val <= percent_tolerance)) {
        duration <- as.numeric(difftime(data$`Time-UTC`[j], data$`Time-UTC`[i], units = "hours"))
        if (duration >= duration_threshold) {
          stable_periods[[length(stable_periods) + 1]] <- data[i:j, ]
          i <- j + 1
          break
        }
      }
      j <- j + 1
    }
    i <- i + 1
  }
  
  return(stable_periods)
}

# Calculate sustained stats per group per species
calculate_sustained_stats <- function(df, species_col) {
  df %>%
    group_by(Configurations_Broader, block_by_stages) %>%
    group_map(~ {
      stable_list <- find_stable_periods(.x, species_col)
      
      if (length(stable_list) == 0) {
        return(tibble(
          Configurations_Broader = .y$Configurations_Broader,
          block_by_stages = .y$block_by_stages,
          Sustained_Min = NA_real_,
          Sustained_Max = NA_real_,
          Sustained_Avg = NA_real_,
          Sustained_Min_Start_OH = NA_real_,
          Sustained_Min_End_OH = NA_real_,
          species = species_col
        ))
      }
      
      period_means <- sapply(stable_list, function(p) mean(p[[species_col]], na.rm = TRUE))
      period_starts <- sapply(stable_list, function(p) min(p$`Time-UTC`))
      period_ends <- sapply(stable_list, function(p) max(p$`Time-UTC`))
      
      min_index <- which.min(period_means)
      max_index <- which.max(period_means)
      
      tibble(
        Configurations_Broader = .y$Configurations_Broader,
        block_by_stages = .y$block_by_stages,
        Sustained_Min = round(period_means[min_index], 2),
        Sustained_Max = round(period_means[max_index], 2),
        Sustained_Avg = round(mean(period_means), 2),
        species = species_col
      )
    }) %>%
    bind_rows()
}

# Summary across species
get_ww_stage_block_summary <- function(df, species_list) {
  # Base summary
  summary_df <- df %>%
    filter(!is.na(Configurations_Broader)) %>%
    group_by(Configurations_Broader, block_by_stages) %>%
    summarise(
      start_time = min(`Time-UTC`),
      end_time = max(`Time-UTC`),
      duration_hr = round(as.numeric(difftime(end_time, start_time, units = "hours")), 2),
      count = n(),
      expected_duration = count - 1,
      mismatch = duration_hr != expected_duration,
      start_oh = `Operating-hrs`[which.min(`Time-UTC`)],
      end_oh = `Operating-hrs`[which.max(`Time-UTC`)],
      .groups = "drop"
    )
  
  # Loop through species and bind
  all_stable_stats <- lapply(species_list, function(sp) {
    calculate_sustained_stats(df, species_col = sp)
  }) %>%
    bind_rows()
  
  # Merge and return
  summary_df %>%
    left_join(all_stable_stats, by = c("Configurations_Broader", "block_by_stages")) %>%
    arrange(Configurations_Broader, block_by_stages, species)
}


file_path <- "emissions_data_ops.xlsx"
emissions_data <- read_excel(file_path, sheet = "emissions_ww_fg_op_data")

#species_list <- c("PZ", "AM93-MNPZ", "Ethylendiamine", "AM98-1-M-PZ", "AM95-FPZ", "Ammonia", "Acetonitrile", "Acetaldehyde","Formaldehyde","Methanol","Aminoacetaldehyde/acetamide", "DinitrosoPZ","NitronitrosoPZ","AM94-Nitrosamine-EDA","MMA","DMA","Acetone","Butanal/butanone")
species_list <- c("PZ", "AM93-MNPZ", "Ammonia", "AM94-Nitrosoamine-EDA","Nitrosamines","NitroPZ","DinitroPZ","AM96-Piperazin-2-one","AM97-Piprazin-2-ol","AM99", "AM100", "AM101-PZ-1-Carboxlic acid", "AM102", "C5H10N2", "C6H12N2","C4H9N2+", "AM105", "AM106","AM107","AM104-Acetone",
                  "Methylethylenediamine","Simple-amines", "Pyrrole",	"MePyrrole",	"Pyrazine","C1-Imidazole/Dyhdropyrazine","C2-Imidazoles",	"C3-pyrazines","C6H9NO","C7H8N2",	"C4H6O2",	"C8H8N2", "C8H10N2")
final_summary <- get_ww_stage_block_summary(emissions_data, species_list)

# Write to Excel with multiple sheets
write.xlsx(final_summary, file = "Emis_stats_emissions_rem.xlsx", rowNames = FALSE)
