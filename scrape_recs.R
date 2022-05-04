library(tidyverse)
library(xml2)
library(rvest)
library(rlang)

rec_url <- "https://www.nice.org.uk/guidance/ng205/chapter/Recommendations"
rec_html <- read_html(rec_url)

section_name_fn <- function(id, html) {
    css <- sprintf("div#%s h3", id)
    
    html %>% 
        html_elements(css) %>% 
        html_text2() %>% 
        str_trim()
}

subsection_name_fn <- function(id, html) {
    css <- paste0("#", id, " h4.title")
    
    html %>% 
        html_elements(css) %>% 
        html_text2()
}

subsection_fn <- function(section_id, html) {
    css <- sprintf("div#%s div.section", section_id)
    
    html %>% 
        html_elements(subsection_css) %>% 
        html_attr("id") %>% 
        tibble(id = .) %>% 
        mutate(title = map_chr(id, subsection_name_fn, rec_html))
}



rec_fn <- function(id, html) {
    css <- sprintf("div#%s div.recommendation_text", id)
    
    table <- html %>% 
        html_elements(css) %>% 
        html_attr("id") %>%
        tibble(id = .)
    
    return(table)
}

rec_number_fn <- function(id, html) {
    rec_number_css <- sprintf("div#%s span.paragraph-number", id)
    
    html %>% 
        html_element(rec_number_css) %>% 
        html_text2()
}

chunk_fn <- function(id, html) {
    rec_css <- paste0("div#", id)
    
    html %>% 
        html_elements(rec_css) %>% 
        html_children() %>% 
        as.character()
}

format_fn <- function(chunk) {
    if (str_detect(chunk, '^<p class=\"numbered-paragraph\">')) {
        chunk %>% 
            str_remove('^<p class=\"numbered-paragraph\">') %>% 
            str_remove('(<span class=\"paragraph-number\">)[^>]+(</span>)') %>% 
            str_remove_all('(<a id=\")[^>]+("></a>)') %>% 
            str_remove_all('(<a class=\"link\")[^>]+(>)') %>% 
            str_remove_all('</a>') %>% 
            str_remove('</p>$') %>% 
            str_remove_all('\n|\r') %>% 
            str_trim()
        
    } else if (str_detect(chunk, '^<div class=\"itemizedlist indented\">')) {
        chunk %>% 
            str_remove_all('(<a id=\")[^>]+("></a>)') %>% 
            str_remove_all('(<a class=\"link\")[^>]+(>)') %>% 
            str_remove_all('</a>') %>% 
            str_extract_all('(?<=<li class=\"listitem\">\r\n)\\s*<p>[^>]+(?=</p>)') %>% 
            simplify() %>% 
            str_remove('<p>') %>% 
            str_trim() %>% 
            paste("\n\u2022 ", ., collapse = "")
    }
}

section_id <- rec_html %>% 
    html_elements("div.chapter > div.section") %>% 
    html_attr("id") %>% 
    str_replace_all("\\.", "\\\\.") %>% 
    tibble(id = .) %>% 
    mutate(name = map(id, section_name_fn, rec_html)) %>% 
    filter(str_starts(name, '\\d')) %>% 
    mutate(subsection = map(id, subsection_fn, rec_html))

section_id <- section_id %>% 
    mutate(subsection = map(subsection,
                            ~ .x %>%
                                mutate(recs = map(id, rec_fn, rec_html)))) %>% 
    mutate(subsection = map(subsection,
                            ~ .x %>%
                                mutate(recs = map(recs, 
                                                  ~ .x %>%
                                                      mutate(number = map_chr(id, rec_number_fn, rec_html),
                                                             chunk = map(id, chunk_fn, rec_html) %>% 
                                                                 map(. %>% map(format_fn)))))))


html_children(section_recs)
xml_structure(section_recs)
# xml2::as_list(section_recs)
