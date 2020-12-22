library(tidyverse)
library(readxl)
library(export)
library(SPHSUgraphs)

excel_sheet <- read_xlsx('Data extraction - effect or not.xlsx', sheet = "Transposed")[5:105,] %>% 
  filter(!is.na(Reference))


# how many have school-based reported in at least 1 column
excel_sheet %>% select(Reference,
                       school_sti = `School-based STI-focussed education`,
                       school_preg = `School-based pregnancy education`) %>%
  mutate(reports = ifelse(!is.na(school_sti) | !is.na(school_preg), 1, 0)) %>% 
  summarise(tot = sum(reports))

# how many have community-based in one column
excel_sheet %>%
  mutate(reports = ifelse(!is.na(`Community-based STI education`) | !is.na(`Community-based pregnancy education`), 1, 0)) %>% 
  summarise(tot = sum(reports))

# how many have contraception-interventions in one column
excel_sheet %>%
  mutate(reports = ifelse(!is.na(`Contraception access (other)`) | !is.na(`Condom promotion/distribution`) | !is.na(`Advance supply of EC**`) | !is.na(`Contraception initiation follow-up`), 1, 0)) %>% 
  summarise(tot = sum(reports))

# how many have adolescent development in one column
excel_sheet %>%
  mutate(reports = ifelse(!is.na(`Personal development (inc. volunteer work)`) | !is.na(`Early-years intervention`) | !is.na(`Vocational/academic training`), 1, 0)) %>% 
  # filter(reports !=0) %>% 
  # pull(Reference)
  summarise(tot = sum(reports))

# how many have dm-interventions in one column
excel_sheet %>%
  mutate(reports = ifelse(!is.na(`Digital media-based SH intervention`) | !is.na(`Digital-media based intervention (targetted)`), 1, 0)) %>% 
  summarise(tot = sum(reports))

# summarising table -------------------------------------------------------

excel_sheet_counts <- read_xlsx('Data extraction - effect or not.xlsx', sheet = "Evidence by type of exposure")

excel_sheet_counts %>% 
  tally(Total>0)

excel_sheet_counts %>% 
  pull(Categorisation) %>% 
  unique() %>% 
  length()

  

# how many +ves -----------------------------------------------------------

rowAny <- function(x) rowSums(x, na.rm=TRUE) > 0

excel_sheet %>% 
  select(-Notes) %>% 
  group_by(Reference) %>% 
  mutate(across(everything(), ~ ifelse(.x=="+", 1, 0))) %>% 
  filter(rowAny(across(everything(), ~ .x==1))) %>% 
  nrow()

# 64 reviews found at least one +ve result


# nice graphs for each type -----------------------------------------------

excel_sheet_counts %>% 
  filter(Categorisation %in% c("Abs", "Cli", "Con", "Sch", "ConTech") & Intervention_Environment == "Int") %>% 
  mutate(Categorisation = factor(Categorisation, levels = rev(c("Sch", "Abs", "Cli", "Con", "ConTech")), ordered = TRUE)) %>% 
  group_by(Categorisation) %>% 
  mutate(`Intervention or Exposure` = fct_reorder(`Intervention or Exposure`, `Studies showing +ve impact`, .desc = FALSE)) %>% 
  ggplot(aes(`Intervention or Exposure`, `Studies showing +ve impact`)) +
  geom_col(aes(y=1), position="fill" ,fill = "grey") +
  geom_col(fill=sphsu_cols("Cobalt")) + 
  coord_flip() +
  theme_void()


graph2office(last_plot(), "SR_vote_graphs.pptx", width = 2, height = 6, append = TRUE)


excel_sheet_counts %>% 
  filter(Categorisation %in% c("Com", "Dev", "DigInt") & Intervention_Environment == "Int") %>% 
  mutate(Categorisation = factor(Categorisation, levels = rev(c("Com", "Dev", "DigInt")), ordered = TRUE)) %>% 
  group_by(Categorisation) %>% 
  mutate(`Intervention or Exposure` = fct_reorder(`Intervention or Exposure`, `Studies showing +ve impact`, .desc = FALSE)) %>% 
  ggplot(aes(`Intervention or Exposure`, `Studies showing +ve impact`)) +
  geom_col(aes(y=1), position="fill" ,fill = "grey") +
  geom_col(fill=sphsu_cols("Cobalt")) + 
  coord_flip() +
  theme_void()
  # theme(line = element_blank(), 
  #       rect = element_blank(), 
  #       axis.title = element_blank(),
  #       axis.text.x = element_blank(),
  #       axis.text.y = element_blank())

graph2office(last_plot(), "SR_vote_graphs.pptx", width = 2, height = 6, append = TRUE)




excel_sheet_counts %>% 
  filter(Categorisation %in% c("EduInt", "Fam", "Media", "Peer", "RRP", "Soc", "VIS") & Intervention_Environment == "Int") %>% 
  mutate(Categorisation = factor(Categorisation, levels = rev(c("EduInt", "Fam", "Media", "Peer", "RRP", "Soc", "VIS")), ordered = TRUE)) %>% 
  group_by(Categorisation) %>% 
  mutate(`Intervention or Exposure` = fct_reorder(`Intervention or Exposure`, `Studies showing +ve impact`, .desc = FALSE)) %>% 
  ggplot(aes(`Intervention or Exposure`, `Studies showing +ve impact`)) +
  geom_col(aes(y=1), position="fill" ,fill = "grey") +
  geom_col(fill=sphsu_cols("Cobalt")) + 
  coord_flip() +
  theme_void()

graph2office(last_plot(), "SR_vote_graphs.pptx", width = 2, height = 6, append = TRUE)



excel_sheet_counts %>% 
  filter(Categorisation %in% c("Cul", "DigEn", "Eco", "EduEn", "Emp") & Intervention_Environment == "Env") %>% 
  mutate(Categorisation = factor(Categorisation, levels = rev(c("Cul", "DigEn", "Eco", "EduEn", "Emp")), ordered = TRUE)) %>% 
  group_by(Categorisation) %>% 
  mutate(`Intervention or Exposure` = fct_reorder(`Intervention or Exposure`, `Studies showing +ve impact`, .desc = FALSE)) %>% 
  ggplot(aes(`Intervention or Exposure`, `Studies showing +ve impact`)) +
  geom_col(aes(y=1), position="fill" ,fill = "grey") +
  geom_col(fill=sphsu_cols("Cobalt")) + 
  coord_flip() +
  theme_void()

graph2office(last_plot(), "SR_vote_graphs.pptx", width = 2, height = 6, append = TRUE)
