library(lubridate)
library(readxl)
library(stringr)
library(dplyr)

setwd("/Users/samuelho/Desktop/books/Rworking")

keydates <- read_excel("update for forecast.xlsx")

# Enter the start year and end year of forecast here

startyr = 2020
endyr = 2023

compile <- data.frame()

for(i in startyr:endyr){
  
  # Filter out the subject year for which you will be creating 
  
  subjectyr <- filter(keydates, year == i)
  
  # Populate a whole year's worth of dates
  
  qualify <- seq.Date(from = as.Date(paste0(i,"-01-01")), to = as.Date(paste0(i, "-12-31")), by = 1)
  
  # Select dates which are "blocked out" from consideration as show dates
  
  cnyfilter <- seq.Date(from = as.Date(subjectyr$cny) - 7, to = as.Date(subjectyr$cny) + 7, by = 1)
  
  nusmidfilter <- seq.Date(from = as.Date(subjectyr$nusmid) - 7, to = as.Date(subjectyr$nusmid) + 7, by = 1)
  
  nusfinfilter <- seq.Date(from = as.Date(subjectyr$nusfin) - 7, to = as.Date(subjectyr$nusfin) + 14, by = 1)
  
  nusorfilter <- seq.Date(from = as.Date(subjectyr$nusor), to = as.Date(subjectyr$nusor) + 7, by = 1)
  
  hrpfilter <- seq.Date(from = as.Date(subjectyr$hrp) - 7, to = as.Date(subjectyr$hrp) + 7, by = 1)
  
  hrhfilter <- seq.Date(from = as.Date(subjectyr$hrh) - 1, to = as.Date(subjectyr$hrh), by = 1)
  
  natdayfilter <- seq.Date(from = as.Date(paste0(i, "-08-09")), to = as.Date(paste0(i, "-08-09")) + 3, by = 1)
  
  # Apply filter - you can choose to drop or add whatever filters here
  
  blocked <- c(cnyfilter, nusmidfilter, nusfinfilter, nusorfilter, hrpfilter, hrhfilter)
  
  # Find dates which pass the filter
  
  qualify <- data.frame(qualify)
  colnames(qualify) <- "date"
  blocked <- data.frame(blocked)
  colnames(blocked) <- "date"
  
  qualify <- anti_join(qualify, blocked, by = "date")
  
  # Find Fridays, Saturdays and Sundays which qualify - note that lubridate parses Saturdays as wday() = 7, rather than 6
  
  qualify <- qualify %>% 
    mutate(day = wday(date)) %>% 
    filter(day == 6 | day == 7 | day == 1) %>% 
    mutate(day = ifelse(day == 6, "Friday", day)) %>% 
    mutate(day = ifelse(day == 7, "Saturday", day)) %>% 
    mutate(day = ifelse(day == 1, "Sunday", day))
  
  # Take a key from each weekend, using Sundays as the key, to find distinct, full weekends which qualify
  
  qualifysundays <- qualify %>% 
    mutate(fullweekend = ifelse(day == "Sunday" & date == lag(date, 1) + 1 & date == lag(date, 2) + 2, "yes", NA)) %>% 
    filter(fullweekend == "yes")
  
  # Pick the latest qualifying Sunday that is at least 3 weeks before the closing of the financial year
  
  financialyearclosing <- as.Date(paste0(i, "-08-31"))
  
  midyearprod <- qualifysundays %>% 
    filter(date < financialyearclosing - 3*7)
  
  midyearprod <- midyearprod[nrow(midyearprod), ]
  
  # First Company Meeting is 15 working weeks before the show date
  
  midyearprod$firstcoymeet <- filter(qualify, day == "Saturday")[last(which(filter(qualify, day == "Saturday")$date < as.Date(midyearprod$date) - 15*7)), 'date']
  
  # Acceptance deadline is 2 weeks before first company meeting
  
  midyearprod$acceptdeadline <- midyearprod$firstcoymeet - 1*7
  
  # Release of audition results is 1 week before acceptance deadline
  
  midyearprod$auditionresults <- midyearprod$acceptdeadline - 1*7
  
  # Define audition weeks
  
  midyearprod$audweek3 <- midyearprod$auditionresults - 1*7
  midyearprod$audweek2 <- midyearprod$audweek3 - 1*7
  midyearprod$audweek1 <- midyearprod$audweek2 - 1*7
  
  # Sign-up deadline is 1 week before first audition week
  
  midyearprod$signupdeadline <- midyearprod$audweek1 - 1*7
  
  # Sign-ups open 3 weeks before sign up deadline
  
  midyearprod$signupopen <- midyearprod$signupdeadline - 2*7
  
  # Recruitment material deadline is 2 weeks before sign-up opens
  
  midyearprod$recmatdeadline <- midyearprod$signupopen - 2*7
  
  # Check if audition dates clash with CNY/Hari Raya Puasa/Hari Raya Haji, if so, push preceding schedule back by one week and repeat
  
  cnyhariraya <- c( seq.Date(from = as.Date(subjectyr$hrh) -3, to = as.Date(subjectyr$hrh) +3, by = 1), seq.Date(from = as.Date(subjectyr$hrp) -3, to = as.Date(subjectyr$hrp) +3, by = 1), seq.Date(from = as.Date(subjectyr$cny) -3, to = as.Date(subjectyr$cny) +3, by = 1) )
  
  while(midyearprod$audweek3 %in% cnyhariraya){
    midyearprod$audweek3 <- midyearprod$audweek3 -7
    midyearprod$audweek2 <- midyearprod$audweek2 -7
    midyearprod$audweek1 <- midyearprod$audweek1 -7
    midyearprod$signupdeadline <- midyearprod$signupdeadline - 7
    midyearprod$signupopen <- midyearprod$signupopen - 7
    midyearprod$recmatdeadline <- midyearprod$recmatdeadline - 7
  }
  
  while(midyearprod$audweek2 %in% cnyhariraya){
    midyearprod$audweek2 <- midyearprod$audweek2 -7
    midyearprod$audweek1 <- midyearprod$audweek1 -7
    midyearprod$signupdeadline <- midyearprod$signupdeadline - 7
    midyearprod$signupopen <- midyearprod$signupopen - 7
    midyearprod$recmatdeadline <- midyearprod$recmatdeadline - 7
  }
  
  while(midyearprod$audweek1 %in% cnyhariraya){
    midyearprod$audweek1 <- midyearprod$audweek1 -7
    midyearprod$signupdeadline <- midyearprod$signupdeadline - 7
    midyearprod$signupopen <- midyearprod$signupopen - 7
    midyearprod$recmatdeadline <- midyearprod$recmatdeadline - 7
  }
  
  startyearprod <- data.frame(qualifysundays[which(qualifysundays$date <= (midyearprod$signupopen + 1)), 'date'])
  startyearprod <- startyearprod[nrow(startyearprod),]
  startyearprod <- data.frame(startyearprod)
  colnames(startyearprod) <- "date"
  
  startyearprod$firstcoymeet <- startyearprod$date - 15*7 - 1
  
  # Acceptance deadline is 2 weeks before first company meeting
  
  startyearprod$acceptdeadline <- startyearprod$firstcoymeet - 1*7
  
  # Release of audition results is 1 week before acceptance deadline
  
  startyearprod$auditionresults <- startyearprod$acceptdeadline - 1*7
  
  # Define audition weeks
  
  startyearprod$audweek3 <- startyearprod$auditionresults - 1*7
  startyearprod$audweek2 <- startyearprod$audweek3 - 1*7
  startyearprod$audweek1 <- startyearprod$audweek2 - 1*7
  
  # Sign-up deadline is 1 week before first audition week
  
  startyearprod$signupdeadline <- startyearprod$audweek1 - 1*7
  
  # Sign-ups open 3 weeks before sign up deadline
  
  startyearprod$signupopen <- startyearprod$signupdeadline - 2*7
  
  # Recruitment material deadline is 2 weeks before sign-up opens
  
  startyearprod$recmatdeadline <- startyearprod$signupopen - 2*7
  
  startyearprod$prod <- paste0(i, " Start of Year Production")
  midyearprod$prod <- paste0(i, " Mid-Year Production")
  
  midyearprod <- select(midyearprod, date, firstcoymeet, acceptdeadline, auditionresults, audweek3, audweek2, audweek1, signupdeadline, signupopen, recmatdeadline, prod)
  
  commrecruit = data.frame(matrix(ncol =11))
  colnames(commrecruit) = c("date", "firstcoymeet", "acceptdeadline", "auditionresults", "audweek3", "audweek2", "audweek1", "signupdeadline", "signupopen", "recmatdeadline", "prod")
  
  commrecruit$prod = paste0(i, " Committee Handover")
  commrecruit$date = as.Date(paste0(i, "-08-31"))
  commrecruit$firstcoymeet = floor_date(commrecruit$date, unit = "week", week_start = 6) - 5*7
  commrecruit$acceptdeadline = commrecruit$firstcoymeet - 1*7
  commrecruit$auditionresults = commrecruit$acceptdeadline - 1*7
  commrecruit$audweek1 = commrecruit$auditionresults - 1*7
  commrecruit$signupdeadline = commrecruit$audweek1 - 1*7
  commrecruit$signupopen = commrecruit$signupdeadline - 1*7
  commrecruit$recmatdeadline = commrecruit$signupopen - 1*7
  
  compile <- rbind(compile, startyearprod)
  compile = rbind(compile, commrecruit)
  compile <- rbind(compile, midyearprod)
}

# Re-align sign-up openings to show date of previous production, in order to allow more production time and to coincide sign-up publicity with shows. Note that all surrounding dates must be correspondingly shifted, except for show date of course.

handover = compile %>% filter(str_detect(compile$prod, "Committee Handover"))
rownames(handover) = handover$prod
finalcompile = compile
rownames(finalcompile) = finalcompile$prod
compile = compile %>% filter(!str_detect(compile$prod, "Committee Handover"))
rownames(compile) = compile$prod

for(i in 2:nrow(compile)){
  
  if(compile[i, 'signupopen'] > compile[i-1, 'date']){
    shiftdifference = ( compile[i, 'signupopen'] - compile[i-1, 'date'] ) + 1
    
    compile[i,'signupopen'] = compile[i,'signupopen'] - shiftdifference
    compile[i,'recmatdeadline'] = compile[i,'recmatdeadline'] - shiftdifference
    compile[i,'signupdeadline'] = compile[i,'signupdeadline'] - shiftdifference
    compile[i,'audweek1'] = compile[i,'audweek1'] - shiftdifference
    compile[i,'audweek2'] = compile[i,'audweek2'] - shiftdifference
    compile[i,'audweek3'] = compile[i,'audweek3'] - shiftdifference
    compile[i,'auditionresults'] = compile[i,'auditionresults'] - shiftdifference
    compile[i,'acceptdeadline'] = compile[i,'acceptdeadline'] - shiftdifference
    compile[i,'firstcoymeet'] = compile[i,'firstcoymeet'] - shiftdifference
    
  }
}

finalcompile[rownames(compile), ] = compile
finalcompile[rownames(handover), ] = handover

finalcompile <- select(finalcompile, `Production` = prod, `Recruitment material deadline` = recmatdeadline, `Sign-ups open` = signupopen, `Sign-up deadline` = signupdeadline, `Audition Week 1` = audweek1, `Audition Week 2` = audweek2, `Audition Week 3` = audweek3, `Release of audition results` = auditionresults, `Acceptance deadline` = acceptdeadline, `First Company Meeting` = firstcoymeet, `Show date` = date)



write.csv(finalcompile, "2023 Forecast.csv", row.names = FALSE)

# Coerce into GCal import format

gcalimport = data.frame()

for(i in 1:nrow(finalcompile)){

addition = t(finalcompile[i,])
rownames(addition) = paste0(colnames(addition), " ", rownames(addition))
addition = addition[-1,]

gcalimport = rbind(gcalimport, data.frame(addition))

}

colnames(gcalimport) = "Start Date"
gcalimport$`Start Date` = format(as.Date(gcalimport$`Start Date`), '%m/%d/%Y')
gcalimport$`Start Time` = NA
gcalimport$`End Date` = gcalimport$`Start Date`
gcalimport$`All day event` = TRUE
gcalimport$`End Time` = NA
gcalimport$Description = NA
gcalimport$Location = NA
gcalimport$Subject = rownames(gcalimport)

gcalimport = select(gcalimport, Subject, `Start Date`,	`Start Time`,	`End Date`,	`End Time`,	`All day event`,	Description,	Location)

write.csv(gcalimport, "gcalimport.csv", row.names = FALSE)
