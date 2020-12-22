read_new_qa_file <- function() {
  
  new_file <-  "C:\\Users\\2286432b\\OneDrive - University of Glasgow\\Teenage pregnancy - docs in development\\Review docs - data extraction in progress\\quality assessment.csv"
  
  dl <-  "C:\\Users\\2286432b\\Downloads"
  
  
  files <- list.files(dl)
  
  file <- files %>% 
    str_subset("review_40742_quality") %>% 
    tibble(files = .) %>% 
    mutate(file_num = str_extract(files, "(?<=\\()\\d+(?=\\)\\.csv)") %>% as.integer()) %>% 
    top_n(1, file_num) %>% 
    pull(files) %>% 
    paste0(dl, "\\", .)
  
  file.copy(file, new_file, overwrite = TRUE, recursive = FALSE)
  
}

add_row_to_docx <- function(df_row, doc_temp = doc_temp) {
  doc_temp %>% body_add_par("")
  
  df_row %>% 
    pivot_longer(Reference:`Reviewed domains`) %>%
    body_add_table(doc_temp, ., 
                   header = FALSE,
                   style = "Plain Table 1",
                   first_row = FALSE,
                   first_column = TRUE)
  
  
}

add_ded_table_to_docx <- function(df_row, doc_temp = doc_notes) {
  if (nrow(df_row) > 0) {
    df_row %>% 
      body_add_table(doc_temp, .,
                     style = "Plain Table 1",
                     first_row = TRUE,
                     first_column = TRUE)
  } else {
    body_add_par(doc_temp, "No high-quality studies found")
  }
}

high_qual_by_cause <- function(data = findings_by_qual, ..., quality = c("all", "high", "low")) {
  
  tryCatch({
  quality <- match.arg(quality)
  }, error = function(e) stop("quality must be \"high\", \"low\" or \"all\"")
  )
  
  rowAny <- function(x) rowSums(x) < ncol(data %>% select(...))
  
  df <-   data %>% 
    group_by(rating) %>% 
    group_map(function(df, y){
      studies <- df %>% 
        select(Reference, score, ..., Title, `Published Year`) %>% 
        filter(rowAny(across(.cols = c(...), .fns =  ~ is.na(.x)))) %>%
        arrange(desc(`Published Year`))
      list(quality = y, studies = studies)
    })
  
  if (quality == "all") {
    map(df, ~ {
      cat(paste0("Studies of ", .x$quality$rating, " rating: ", nrow(.x$studies), "\n"))
      print(.x$studies)
      cat("\n")
    })
    cat(paste0("Total studies: ", nrow(df[[1]]$studies) + nrow(df[[2]]$studies)))
  } else {
    df_out <- map_dfr(df, ~ tibble(rating = .x$quality$rating, data = list(.x$studies))) %>% 
      filter(rating == quality) %>%
      unnest(cols = data) %>% 
      select(- rating)
      cat(paste0("Studies of ", quality, " rating: ", nrow(df_out), "\n"))
      print(df_out)
      cat("\n")
  }
  
}

