
#R script for german Q-Exam excel exports to analyze exam metrics


library(readxl)
library(dplyr)
library(purrr)
library(stringr)
library(tidyr)


# Path to your folder
folder_path <- "..."

# List all Excel files
excel_files <- list.files(folder_path, pattern = "\\.xlsx$", full.names = TRUE)

# Function to read the sheet and add filename
read_item_analyse <- function(file_path) {
  # Read everything first
  df <- read_excel(file_path, sheet = "Item-Analyse", skip = 1, col_types = "text")
  
  # Count student columns dynamically
  student_cols <- str_detect(names(df), "^Student\\*in")  # matches columns like Student*in 1, Student*in 2, ...
  n_students <- sum(student_cols)
  
  # Keep first 7 columns and clean
  df %>%
    select(1:7) %>%
    rename(
      item.nr = 1,
      item.id = 2,
      fach = 3,
      autor = 4,
      schwierigkeit = 5,
      trenns = 6,
      cronbach.auswirk = 7
    ) %>%
    filter(autor == "....") %>% #filter to your exam questions
    mutate(
      source_file = basename(file_path),
      n = n_students  # add number of students as a column
    )
}


# Read all files and combine
all_data <- excel_files %>%
  map_dfr(read_item_analyse) %>%
  mutate(
    across(c(schwierigkeit, trenns, cronbach.auswirk), as.numeric),  # convert to numeric
    across(c(schwierigkeit, trenns), round, 2),                      # round 2 decimals
    cronbach.auswirk = round(cronbach.auswirk, 5)                    # round Cronbach to 3 decimals
  )

# Check result
all_data %>%
  print(n = Inf)

#which(duplicated(all_data$item.id))


# Function to read, clean, and rename each sheet
read_distraktoren <- function(file_path) {
  read_excel(file_path, sheet = "Distraktorenanalyse", col_names = FALSE) %>%  # read as text
    select(1:4) %>%  # keep only the first 4 columns
    rename(
      text  = 1,
      abs   = 2,
      rel   = 3,
      trenn = 4
    ) %>%
    mutate(source_file = basename(file_path))  # add filename
}

# Read all files and combine
all_distraktoren <- excel_files %>%
  map_dfr(read_distraktoren)

all_distraktoren %>%
  print(n = Inf)


all_distraktoren <- all_distraktoren %>%
  mutate(
    Frage = if_else(
      str_detect(text, "^Frage\\s+Nr\\."),
      as.integer(str_extract(text, "\\d+")),
      NA_integer_
    ),
    FragenID = if_else(
      str_detect(text, "^Fragen-ID"),
      as.integer(str_extract(text, "\\d+")),
      NA_integer_
    )
  ) %>%
  group_by(source_file) %>%    
  fill(Frage, FragenID, .direction = "down") %>%
  ungroup()


# Prepare wide distractors
wide_distractors <- all_distraktoren %>%
  filter(str_detect(text, "^\\d+\\.")) %>%
  semi_join(
    all_data %>% mutate(item.id = as.integer(item.id)) %>% select(item.id, source_file),
    by = c("FragenID" = "item.id", "source_file" = "source_file")
  ) %>%
  # temporarily convert all to character for pivot_longer
  transmute(
    FragenID,
    source_file,
    Text  = str_trim(str_remove(text, "^\\d+\\.\\s*")),
    abs   = as.character(abs),
    rel   = as.character(rel),
    trenn = as.character(trenn)
  ) %>%
  group_by(FragenID, source_file) %>%
  mutate(Antwort = row_number()) %>%
  ungroup() %>%
  pivot_longer(
    cols = c(Text, abs, rel, trenn),
    names_to = "metric",
    values_to = "value"
  ) %>%
  mutate(
    col = if_else(
      metric == "Text",
      paste0("answer_text_", Antwort),
      paste0(metric, "_", Antwort)
    )
  ) %>%
  group_by(FragenID, source_file, col) %>%
  summarise(value = first(value), .groups = "drop") %>%
  pivot_wider(
    names_from = col,
    values_from = value
  ) %>%
  # convert metrics back to numeric
  mutate(
    across(matches("^(abs|rel|trenn)_"), ~ round(as.numeric(.x), 3))
  )

# Join wide distractors to all_data
all_data_with_distractors <- all_data %>%
  mutate(item.id = as.integer(item.id)) %>%
  left_join(
    wide_distractors,
    by = c("item.id" = "FragenID", "source_file" = "source_file")
  )



