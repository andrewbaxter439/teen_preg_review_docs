library(tidyverse)
library(readxl)

endnote <- read_tsv('included_studies.txt', col_names = c("Authors", "Year", "DOI", "Title"))

excel_sheet <- read_xlsx('Data extraction - effect or not.xlsx', sheet = "Transposed")[5:100,] %>% 
  filter(!is.na(Reference))

excel_inc <- excel_sheet %>% transmute(au1 = str_extract(Reference, "^\\w*"),
                                       year = str_extract(Reference, "\\d{4,5}\\S{0,3}")) %>% 
  filter(!is.na(au1)) 

endnote_inc <- endnote %>% transmute(au1 = str_extract(Authors, "^\\w*"),
                                     year = as.character(Year))

covidence_inc <- read_csv("included_covidence_07092020.csv") %>% 
  transmute(au1 = str_extract(Authors, "^\\w*"),
            year = as.character(`Published Year`))

excel_inc %>% anti_join(endnote_inc, by = c("au1", "year"))  # RESOLVED
excel_inc %>% anti_join(endnote_inc,.)  #  RESOLVED

endnote_inc %>% anti_join(excel_inc, by = c("au1", "year"))

# Excel and endnote synced

covidence_inc %>% anti_join(excel_inc)  #  RESOLVED
excel_inc %>% anti_join(covidence_inc)  # Allen and Kirby (2001) extracted in Excel but not in covidence. Swann and Trivedi not in covidence

covidence_inc %>% anti_join(endnote_inc)  #  RESOLVED
endnote_inc %>% anti_join(covidence_inc)  # Allen and Kirby (2001) in endnote but not in covidence. Swann and Trivedi not in covidence



# finding by category -----------------------------------------------------

find_cat <- function(sheet, category) {
  sheet %>% 
    arrange(Reference) %>% 
    filter(!is.na({{category}})) %>% 
    select(Reference) %>% 
    View()
}
