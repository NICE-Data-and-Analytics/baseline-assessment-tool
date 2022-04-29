library(tidyverse)
library(xml2)
library(rvest)
library(rlang)

rec_url <- "https://www.nice.org.uk/guidance/ng205/chapter/Recommendations"
rec_html <- read_html(rec_url)

sections <- rec_html %>% 
    html_elements("div.section h3") %>% 
    html_text2()

section_id <- rec_html %>% 
    html_elements("div.chapter > div.section") %>% 
    html_attr("id") %>% 
    str_replace_all("\\.", "\\\\.")

n <- section_id[[2]]


subsection_name_fn <- function(id, html) {
    css <- paste0("#", id, " h4.title")
    
    html %>% 
        html_elements(css) %>% 
        html_text2()
}

subsection_css <- sprintf("div#%s div.section", n)

subsection_id <- rec_html %>% 
    html_elements(subsection_css) %>% 
    html_attr("id") %>% 
    tibble(id = .) %>% 
    mutate(title = map_chr(id, subsection_name_fn, rec_html))

rec_fn <- function(id, html) {
    css <- sprintf("div#%s div.recommendation_text", id)
    
    table <- html %>% 
        html_elements(css) %>% 
        html_attr("id") %>%
        tibble(id = .)
    
    return(table)
}

subsection_id <- subsection_id %>% 
    mutate(recs = map(id, rec_fn, rec_html))

rec_id <- "ng205_1_2_1"
rec_number_css <- sprintf("div#%s span.paragraph-number", rec_id)

rec_number <- rec_html %>% 
    html_element(rec_number_css) %>% 
    html_text2()

css <- sprintf("div#%s p", rec_id)

rec <- rec_html %>% 
    html_elements(css) %>% 
    html_text2()

chunk <- rec_html %>% 
    html_elements("div#ng205_1_2_1") %>% 
    html_children() %>% 
    as.character()

chunk[[1]] %>% 
    str_remove('^<p class=\"numbered-paragraph\">\r\n') %>% 
    str_remove('(<span class=\"paragraph-number\">)[^>]+(</span>\r\n)') %>% 
    str_remove_all('(<a id=\")[^>]+("></a>)') %>% 
    str_remove_all('(<a class=\"link\")[^>]+(>)') %>% 
    str_remove_all('</a>') %>% 
    str_remove('</p>$') %>% 
    str_trim()

test <- chunk[[2]] %>% 
    str_remove_all('(<a id=\")[^>]+("></a>)') %>% 
    str_remove_all('(<a class=\"link\")[^>]+(>)') %>% 
    str_remove_all('</a>') %>% 
    str_extract_all('(?<=<li class=\"listitem\">\r\n)\\s*<p>[^>]+(?=</p>)') %>% 
    simplify() %>% 
    str_remove('<p>') %>% 
    str_trim() %>% 
    paste("\u2022 ", .)

    str_extract_all('(?=<li class=\"listitem\">\r\n\\s*<p>)[^>]+(?=</p>)')

html_children(section_recs)
xml_structure(section_recs)
# xml2::as_list(section_recs)

rec_id <- section_recs %>% 
    html_elements("div.recommendation_text") %>% 
    html_attr("id")
