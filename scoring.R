# importing libraries
library(tidyverse)
library(readxl)
library(janitor)

# this read excel function imports everything except the last sheet( grand prize )
read_excel_allsheets_except_last <- function(filename) {
  sheets <- readxl::excel_sheets(filename)
  sheets <- sheets[-c(length(sheets))] # everything but the last sheet 
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  x <- lapply(x, as.data.frame)
  names(x) <- sheets
  x
}

# this cleans up the individual sheets and labels them with their categories 
cleanup_helper <- function(name, ind) {
  current <- sf[[ind]][[name]]
  stored <- row_to_names(current, row_number = 1)
  stored$category <- name
  stored$judge_id <- ind
  stored
}

# this loops through all of the excel files and  cleans + stacks their sheets into a data frame
cleanup <- function(dl, indie) {
  store <- lapply(names(dl), cleanup_helper, ind=indie)
  bind_rows(store)
}

# grab the list of files
fileList <- paste("data/", list.files(path="./data/", pattern=".xlsx"), sep="")

# read in the list of files
sf <- lapply(fileList, read_excel_allsheets_except_last)

# extract data frames from files
sfc <- mapply(cleanup, sf,  indie=c(1,2,3,4,5,6), SIMPLIFY=F)

# clean up some of the data in the files (rename fields, compute score, cleaned photo names)
sfm <- bind_rows(sfc) %>%
  rename("photograph_name"="Photograph Name:", 
         "originality"="Originality (30%):", 
         "technical_quality"="Technical Quality (30%):", 
         "artistic_merit"="Artistic Merit (40%):") %>%
  mutate(originality = as.numeric(originality), 
         technical_quality = as.numeric(technical_quality), 
         artistic_merit=as.numeric(artistic_merit)) %>%
  mutate(score= originality * .3 + technical_quality*.3 + artistic_merit*.4) %>%
  mutate(photograph_name=str_remove_all(photograph_name, "\\*"))  %>%
  select(-c("...5", "...7"))

# drop unnecessary rows, compute the overall score of each photo
sfm <- sfm[!(sfm$photograph_name=="I reordered this one - FYI"),]
sfm <- sfm[!(sfm$photograph_name=="Favorite Category! All very impressive!"),]
sfm <- sfm[!(sfm$photograph_name==" These had the best captions by far!!"),] %>%
  drop_na() %>% 
  group_by(photograph_name) %>% 
  mutate(overall_score = sum(score), num=n()) %>%    
  ungroup()

# write winners out to files
write.csv(sfm, "data/cleaned_scores.csv")

# pull out the category winner
cat_winner <- merge(aggregate(overall_score ~ category, data=sfm, max), sfm, all.X=T) %>%
  select(-c(originality, technical_quality, artistic_merit, score, judge_id)) %>%
  unique()
write.csv(cat_winner, "data/category_winners.csv")

# pull out the overall winner
overall_winner <- sfm[which.max(sfm$overall_score),] %>%
  select(-c(originality, technical_quality, artistic_merit, score))
write.csv(overall_winner, "data/overall_winner.csv")
