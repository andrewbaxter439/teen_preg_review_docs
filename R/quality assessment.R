library(tidyverse)
library(export)
library(patchwork)
library(SPHSUgraphs)
library(readxl)
library(officer)
library(gt)
source("R/functions.R")
source("R/study id resolver.R")
theme_set(theme_sphsu_light())

#read_new_qa_file()
# Remove Underhills??
quality_assessment <- readr::read_csv("quality assessment_final.csv")


(output_table <- quality_assessment %>% 
  filter(Reviewer == "Andrew") %>% 
  select(`Covidence #`, Study = `Study Id`, matches("^\\d.*")) %>% 
  filter(across(everything(), ~!is.na(.x))) %>% 
  rename_with(~str_replace(.x, "\t", " ")) %>% 
  mutate(across(-Study, factor)) %>% 
  rowwise() %>% 
  mutate(score = sum(ifelse(c_across(-Study)=="low", 1, 0))) %>% 
  arrange(desc(score), Study) %>% 
  mutate(score = factor(score)) %>% 
  mutate(rating = fct_collapse(score,
                               low = c("1", "2", "0"),
                               high = c("3", "4"))) %>% 
  ungroup())

lapply(NA, function(x) {
  
year_v_score <- output_table %>% 
  mutate(year = as.integer(str_extract(Study, "\\d{4}"))
         ) %>% 
  ggplot(aes(year, fct_rev(score))) + 
  geom_jitter(width = 0, height = 0.1, alpha = 0.8) +
  theme_sphsu_light() +
  labs(title = "Reviews by year and score", x = "Year", y = "Score")


score_dist <- output_table %>% 
  ggplot(aes(fct_rev(score))) +
  geom_bar() +
  geom_text(stat = "count", aes(label = ..count.., 
                                vjust = ifelse(..count..>10, 1.5, -1),
                                colour = ..count.. > 10
                                ), 
            # colour = "grey",
            size = 8) +
  theme_sphsu_light() +
  theme(panel.grid.major.x = element_blank(),
        legend.position = "none") +
  scale_colour_manual(values = c("black", "light grey")) + 
  scale_y_continuous(expand = expansion(add = 0)) +
  labs(title = "Number of reviews by score", x = "Score")


rating_dist <- output_table %>% 
  ggplot(aes(fct_rev(rating))) +
  geom_bar() +
  geom_text(stat = "count", aes(label = ..count..
                                ), 
                                vjust = 1.5,
            colour = "light grey",
            size = 8) +
  theme_sphsu_light() +
  theme(panel.grid.major.x = element_blank()) +
  scale_y_continuous(expand = expansion(add = 0)) +
  labs(title = "Number of reviews by high to low rating", x = "Rating")

  year_v_score / (score_dist + rating_dist)
})

# DONE
doc_qual_tab <- read_docx("R/template_de.docx")


output_table %>%
  left_join(study_id_key, by = c("Study", "Covidence #")) %>%
  select(Reference, 3:8) %>%
  mutate(across(2:5, ~ case_when( 
    .x == "low" ~ "Yes",
    .x == "high" ~ "No"
  ))) %>% 
  body_add_table(doc_qual_tab, ., 
                 header = TRUE,
                 style = "Plain Table 1",
                 first_row = TRUE,
                 first_column = FALSE)

print(doc_qual_tab, "Review-quality.docx")

studies_ratings_abstracts <- output_table %>% 
  left_join(study_id_key, by = c("Study", "Covidence #")) %>% 
  left_join(covidence_inc, by = c("Study", "Covidence #")) %>% 
  select(Reference, score, rating, `Published Year`, Title,  Abstract)


# attaching ratings to results and filtering ------------------------------

excel_sheet <- read_excel('Data extraction - effect or not.xlsx', sheet = "Transposed")[-(1:4),] %>% 
  filter(!is.na(Reference))

findings_by_qual <- excel_sheet %>% 
  inner_join(studies_ratings_abstracts, ., by = "Reference")


findings_by_qual %>% high_qual_by_cause(`School-based pregnancy education`, `School-based STI-focussed education`)
findings_by_qual %>% high_qual_by_cause(`Counselling or medical staff one-to-one`)
findings_by_qual %>% high_qual_by_cause(`School-based SH clinic`)
findings_by_qual %>% high_qual_by_cause(`Teenager SH clinic access and use*`)
findings_by_qual %>% high_qual_by_cause(`Condom promotion/distribution`)
findings_by_qual %>% high_qual_by_cause(`Contraception access (other)`)
findings_by_qual %>% high_qual_by_cause(`Personal development (inc. volunteer work)`, `Vocational/academic training`, `Early-years intervention`)
findings_by_qual %>% high_qual_by_cause(`Public information/media campaign`, quality = "low")
findings_by_qual %>% high_qual_by_cause(`Peer-contact sexual health intervention`)
findings_by_qual %>% high_qual_by_cause(`Targeting rapid-repeat pregnancies`)
findings_by_qual %>% high_qual_by_cause(`Family/Community engagement`)
findings_by_qual %>% high_qual_by_cause(`Changing contraceptive technologies`, quality = "all")
findings_by_qual %>% high_qual_by_cause(`Abstinence-based education`)
# All clinic-based
findings_by_qual %>% high_qual_by_cause(`Teenager SH clinic access and use*`, `Counselling or medical staff one-to-one`, `School-based SH clinic`)
# All contraception promotion
findings_by_qual %>% high_qual_by_cause(`Contraception access (other)`, `Contraception initiation follow-up`, `Advance supply of EC**`, `Condom promotion/distribution`)
# All community-based
findings_by_qual %>% high_qual_by_cause(`Community-based pregnancy education`, `Community-based STI education`)
# Digital-media ones
findings_by_qual %>% high_qual_by_cause(`Digital media-based SH intervention`, `Digital-media based intervention (targeted)`)


# outputting doc for further data extraction ------------------------------

# DONE 
# doc_notes <- read_docx("R/template_de.docx")
# 
# 
# map(
#   colnames(findings_by_qual[7:40]),
#   function(.x) {
#     body_add_par(doc_notes, .x, style = "heading 1")
#     
#     filter(findings_by_qual, rating == "high", !is.na(.data[[.x]])) %>% 
#     select(Reference, score, Title) %>% 
#     mutate(years = "", objectives = "", findings = "") %>% 
#     add_ded_table_to_docx()
#   }
# )
# 
# print(doc_notes, "R/notes on school-based preg.docx")
#
# word doc with title, year, rating and study type ------------------------
# 
# doc_temp <- read_docx("R/template.docx")
# 
# findings_by_qual %>% 
#   mutate(across(everything(), as.character)) %>%
#   pivot_longer(`Family/Community engagement`:`Peer-contact sexual health intervention`) %>% 
#   filter(!is.na(value)) %>% 
#   select(-value) %>% 
#   distinct() %>% 
#   pivot_wider(names_from = name, values_from = name) %>% 
#   unite(`Reviewed domains`, `Targeting rapid-repeat pregnancies`:`Employment outside school hours`, sep = ", ", na.rm = TRUE) %>% 
#   select(Reference, Title, `Published Year`, score, rating, `Reviewed domains`) %>% 
#   unite(`Review quality score`, c(score, rating), sep = " - ") %>% 
#   rowwise() %>%
#   group_walk(~ add_row_to_docx(.x, doc_temp = doc_temp))
#   
# print(doc_temp, "Studies and title by rating.docx")
#


# Review summary statistics ----------------------------------------------

findings_by_qual %>% 
  filter(rating == "high") %>% 
  count()  # 61 reviews of high quality


## graphing rob by domain --------------------------------------------------

output_table %>% 
  filter(rating == "high") %>% 
  pivot_longer(`1 – Was it systematic: defining objectives and setting appropriate study eligibility criteria?`:`4 – Was risk of bias and confounding adequately assessed and were results presented to take account of this?`,
               names_to = "domain", values_to = "judgement") %>% 
  mutate(domain = str_wrap(domain, width = 35)) %>% 
  ggplot(aes(fct_rev(domain), fill = judgement)) +
  stat_count() +
  coord_flip() +
  theme_sphsu_light() +
  scale_fill_sphsu(name = "Risk of Bias", breaks = c("low", "high")) +
  xlab("Domain of review quality") +
  scale_y_continuous("Number of reviews", expand = c(0,0)) +
  theme(panel.grid = element_blank(),
        panel.grid.major.x = element_line(colour = "grey"),
        axis.ticks.y = element_blank(),
        axis.title.y = element_text(margin = margin(10, 10, 10, 10)),
        axis.text = element_text(size = 12),
        legend.position = "bottom") 

## if save needed again
ggsave("RoB plot.png", height = 120, width = 155, units = "mm")


output_table %>% 
  filter(rating == "high", `4 – Was risk of bias and confounding adequately assessed and were results presented to take account of this?` == "high")

output_table %>% 
  filter(rating == "low", `1 – Was it systematic: defining objectives and setting appropriate study eligibility criteria?` == "low")


## Graphing in 2 directions --------------------------------------------------

output_table %>% 
  pivot_longer(`1 – Was it systematic: defining objectives and setting appropriate study eligibility criteria?`:`4 – Was risk of bias and confounding adequately assessed and were results presented to take account of this?`,
               names_to = "domain", values_to = "judgement") %>% 
  mutate(domain = str_wrap(domain, width = 35)) %>% 
  group_by(rating, domain, judgement) %>% 
  count() %>% 
  mutate(n = ifelse(rating == "high", n, -n)) %>% 
  ggplot(aes(fct_rev(domain), n, fill = judgement)) +
  geom_bar(stat = "identity") +
  coord_flip() +
  scale_fill_sphsu(name = "Risk of Bias", breaks = c("low", "high")) +
  xlab("Domain of review quality") +
  scale_y_continuous("Number of reviews",
                     expand = c(0,0),
                     breaks = seq(-20, 60, 20),
                     labels = abs(seq(-20, 60, 20))) +
  theme(panel.grid = element_blank(),
        # panel.grid.major.x = element_line(colour = "grey"),
        axis.line.y = element_blank(),
        axis.ticks.y = element_blank(),
        axis.title.y = element_text(margin = margin(10, 10, 10, 10)),
        axis.text = element_text(size = 12),
        legend.position = "bottom") +
  geom_hline(yintercept = 0, size = 1) +
  scale_x_discrete(expand = expansion(add = c(0, 0.9))) +
  geom_text(data = tibble(label = c("Low-quality\nreviews", "High-quality\nreviews"),
                          x = 4.7,
                          y = c(-3, 3),
                          hjust = c(1, 0)), 
            aes(x, y, label = label, hjust = hjust),
            inherit.aes = FALSE,
            lineheight = 1)
  
ggsave("RoB plot-2way.png", height = 120, width = 155, units = "mm")

## Two graphs -------------------------------------

output_table %>% 
  pivot_longer(3:6,
               names_to = "domain",
               values_to = "judgement") %>% 
  mutate(domain = str_wrap(domain, width = 35)) %>% 
  ggplot(aes(fct_rev(domain), fill = judgement)) +
  stat_count() +
  coord_flip() +
  # facet_grid(~ fct_rev(rating),
  #            shrink = TRUE,
  #            scales = "free_x",
  #            space = "free_x",
  #            labeller = labeller("fct_rev(rating)" = 
  #                                  c("low" = "Low-quality\nreviews",
  #                                    "high" = "High-quality\nreviews"))) +
  scale_fill_manual(name = "Risk of Bias", 
                    breaks = c("low", "high"), 
                    values = sphsu_cols("Cobalt", "Rust", names = FALSE)) +
  xlab("Domain of review quality") +
  scale_y_continuous("Number of reviews",
                     expand = c(0,0),
                     breaks = seq(0, 60, 20)) +
  theme(panel.grid = element_blank(),
        axis.line.y = element_blank(),
        axis.ticks.y = element_blank(),
        axis.title.y = element_text(margin = margin(10, 10, 10, 10)),
        axis.text = element_text(size = 12),
        strip.text = element_text(size = 12),
        legend.position = "bottom",
        strip.background = element_rect(fill = "white")) +
  geom_hline(yintercept = 0, size = 1) +
  scale_x_discrete(expand = expansion(add = c(0, 0)))


ggsave("RoB plot-facet.png", height = 120, width = 155, units = "mm")

## dist hist -----------------------------------------

output_table %>% 
  ggplot(aes(fct_rev(score), fill = rating)) +
  geom_bar() +
  theme_sphsu_light() +
  scale_fill_manual(name = "Study quality", 
                    breaks = c("low", "high"), 
                    values = sphsu_cols("Rust", "Cobalt", names = FALSE)) +
  labs(y = "Number of reviews", x = "Score") +
  scale_y_continuous(expand = expansion(c(0,0.05))) +
  theme(panel.grid.major.x = element_blank()) +
  geom_vline(aes(xintercept = 3.5), linetype = 2)

ggsave("RoB plot - histogram.png", height = 95, width = 155, units = "mm")

# study quality summary table ---------------------------------------------

stud_group_count <- read_excel('Data extraction - effect or not.xlsx', sheet = "Evidence by type of exposure") %>% 
  mutate(Categorisation = fct_inorder(Categorisation))



counts_by_cat <- findings_by_qual %>% 
  select(Reference, rating, `Family/Community engagement`:`Peer-contact sexual health intervention`) %>% 
  mutate(across(everything(), as.character)) %>% 
  pivot_longer(-c(Reference, rating), names_to = "cause") %>% 
  filter(!is.na(value)) %>% 
  group_by(rating, cause, value) %>% 
  tally(name = "value_n") %>% 
  ungroup() %>% 
  group_by(rating, cause) %>% 
  pivot_wider(names_from = "value", values_from = "value_n") %>% 
  mutate(across(where(is.numeric), sum, na.rm = TRUE)) %>% 
  rowwise() %>% 
  mutate(rate_n = `+` + `0` + `NA`) %>% 
  select(rating, cause, pos_revs = `+`, no_ev = `0`, rate_n) %>%
  left_join(stud_group_count %>% select(cause = `Intervention or Exposure`, Categorisation, Intervention_Environment), by = "cause") %>% 
  ungroup() %>% 
  mutate(`Intervention or environment change` = ifelse(Intervention_Environment == "Int", "Intervention", "Environment change")) %>% 
  select(-Intervention_Environment)

counts_by_cat %>% 
  group_by(Categorisation)


# outputting summary table 1 ----------------------------------------------------------------

counts_by_cat %>% 
  mutate(neg = 0) %>% 
  pivot_wider(names_from = "rating", values_from = c("pos_revs", "no_ev", "neg", "rate_n"), names_glue = "{rating}_{.value}") %>% 
  mutate(ind = "") %>% 
  select(ind, 1:11) %>% 
  arrange(Categorisation) %>% 
  mutate(across(high_pos_revs:low_rate_n, replace_na, 0)) %>% 
  group_by(`Intervention or environment change`) %>% 
  group_map( ~ {
    .x %>%
      group_by(Categorisation) %>%
      gt(rowname_col = "ind") %>% 
      # tab_row_group('Intervention',`Intervention or environment change` == 'Intervention', others = 'Environment change')
      tab_spanner(
        label = "High quality reviews",
        columns = starts_with("high")
      ) %>% 
      tab_spanner(
        label = "Low quality reviews",
        columns = starts_with("low")
      ) %>% 
      # tab_stubhead(label = .y) %>%
      # tab_header(title = .y) %>%
      cols_label(
        cause = paste0("Hypothesised Cause - ", .y),
        high_pos_revs = "Reviews reporting positive effects",
        high_no_ev = "Reviews reporting no effects",
        high_neg = "Reviews reporting negative effects",
        high_rate_n = "Total reviews",
        low_pos_revs = "Reviews reporting positive effects",
        low_no_ev = "Reviews reporting no effects",
        low_neg = "Reviews reporting negative effects",
        low_rate_n = "Total reviews"
      )
    
  })
