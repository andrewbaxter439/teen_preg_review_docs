library(tidyverse)

covidence_inc <- read_csv("all_covidence_papers.csv")

study_id_key <- covidence_inc %>% 
  mutate(Reference = str_extract(Notes, "(?<=\\(Included\\): ).*?(?=;.*)")) %>% 
  select(Reference, Study, `Covidence #`)
