library(tidyverse)
library(xml2)
library(rvest)
library(rlang)

rec_url <- "https://www.nice.org.uk/guidance/ng205/chapter/Recommendations"
rec_html <- read_html(rec_url)

# Get section name
section_name_fn <- function(section_id, html) {
    css <- sprintf("#%s h3", section_id)
    
    html %>% 
        html_element(css) %>% 
        html_text2() %>% 
        str_trim()
}

# Get subsection name
subsection_name_fn <- function(subsection_id, html) {
    css <- sprintf("#%s h4.title", subsection_id)
    
    html %>% 
        html_elements(css) %>% 
        html_text2()
}

# Get all subsections in section
subsection_fn <- function(section_id, html) {
    css <- sprintf("#%s div.section", section_id)
    
    subsection_table <- html %>% 
        html_elements(css) %>%
        # Get subsection html ID
        html_attr("id") %>% 
        tibble(id = .) %>% 
        # Get subsection name
        mutate(title = map_chr(id, subsection_name_fn, html))
    
    if (nrow(subsection_table) == 0) {
        subsection_table <- tibble(id = section_id,
                                   title = "No subsections")
    }
    
    return(subsection_table)
}



rec_fn <- function(subsection_id, html) {
    css <- sprintf("#%s div.recommendation_text", subsection_id)
    
    table <- html %>% 
        html_elements(css) %>% 
        html_attr("id") %>%
        tibble(id = .)
    
    return(table)
}

rec_number_fn <- function(rec_id, html) {
    css <- sprintf("#%s span.paragraph-number", rec_id)
    
    html %>% 
        html_element(css) %>% 
        html_text2()
}

chunk_fn <- function(rec_id, html) {
    css <- paste0("#", rec_id)
    
    html %>% 
        html_elements(css) %>% 
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
    # Get html IDs for all sections, e.g. "ng205_1.2-supporting-positive-relationships"
    html_elements("div.chapter > div.section") %>% 
    html_attr("id") %>% 
    str_replace_all("\\.", "\\\\.") %>% 
    tibble(id = .) %>% 
    # Get section names, e.g. "1.2 Supporting positive relationships"
    mutate(name = map_chr(id, section_name_fn, rec_html)) %>%
    # Keep only recommendation sections
    filter(str_starts(name, '\\d')) %>% 
    # Get html ID and name of all subsections in that section
    mutate(subsection = map(id, subsection_fn, rec_html))

section_id <- section_id %>% 
    # Get html IDs for all recommendations in a subsection
    mutate(subsection = map(subsection,
                            ~ .x %>%
                                mutate(recs = map(id, rec_fn, rec_html)))) %>%
    # 
    mutate(subsection = map(subsection,
                            ~ .x %>%
                                mutate(recs = map(recs, 
                                                  ~ .x %>%
                                                      mutate(number = map_chr(id, rec_number_fn, rec_html),
                                                             chunk = map(id, chunk_fn, rec_html) %>% 
                                                                 map(. %>% map(format_fn))))))) %>% 
    unnest()


# html_children(section_recs)
# xml_structure(section_recs)
# xml2::as_list(section_recs)
