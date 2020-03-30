## ----setup, include=FALSE------------------------------------------------
knitr::opts_chunk$set(echo = TRUE)


## ------------------------------------------------------------------------
#install.packages(tidyr)
#install.packages(dplyr)
#install.packages(stringr)
#install.packages("reticulate")
#install.packages("lambda.tools")
#install.packages("xlsx")

library(tidyr)
library(lambda.tools)
library(dplyr)
library(stringr)
library(reticulate)
library(xlsx)


## ------------------------------------------------------------------------
Data <- read.csv(file.choose())


## ------------------------------------------------------------------------
Data_colnames <- Data[1,]
Data_colnames <- lapply(Data_colnames, as.character)


## ------------------------------------------------------------------------
colnames(Data) <- Data_colnames


## ------------------------------------------------------------------------
Data <- Data[-c(1,2),]


## ------------------------------------------------------------------------
Matching_Participants <- dplyr::select(Data, `Full Name`, Gender, Email, Phone, Agency, `Work location (suburb) This information will be used in the matching process`, `Length in Public Sector`, Position, `Please give a brief description of the key aspects of your role. (50 words or less)`, `What is your substantive level? - Selected Choice`, `I am applying to be a`, `Why do you want to be a mentor? (max 3 responses)`, `Which of these Public Sector capabilities are your strengths? (max 2 responses)`, `What are your areas of expertise in the Public Sector? - Selected Choice`, `Are you interested in mentoring more than one mentee?`, `Please tell us anything more about yourself, your application or why you want to be a part of the program.`, `I agree to commit to the program for the duration of 8 months and will attend all required meetings/sessions.`, `Why do you want to be a mentee?(max 2 responses)`, `Which specific Public Sector Capabilities would you like to improve (max 2 responses)`, `What Public Sector Experience are you seeking - Selected Choice`, `What Public Sector Experience are you seeking - Other - Text`)


## ------------------------------------------------------------------------
# Since most columns are character or word type data, therefore start with converting all columns to character is a good start. 
Matching_Participants <-  Matching_Participants %>% mutate_all(as.character)

# The next column to deal with is the substantial level from characters to integers 
Matching_Participants$`What is your substantive level? - Selected Choice`[Matching_Participants$`What is your substantive level? - Selected Choice` == "L1-2"] <- 1.5 ; 
Matching_Participants$`What is your substantive level? - Selected Choice`[Matching_Participants$`What is your substantive level? - Selected Choice` == "L3-4"] <- 3.5; 
Matching_Participants$`What is your substantive level? - Selected Choice`[Matching_Participants$`What is your substantive level? - Selected Choice` == "L5"] <- 5; 
Matching_Participants$`What is your substantive level? - Selected Choice`[Matching_Participants$`What is your substantive level? - Selected Choice` == "L6"] <- 6 ; 
Matching_Participants$`What is your substantive level? - Selected Choice`[Matching_Participants$`What is your substantive level? - Selected Choice` == "L7"] <- 7; 
Matching_Participants$`What is your substantive level? - Selected Choice`[Matching_Participants$`What is your substantive level? - Selected Choice` == "L8"] <- 8 ; 
Matching_Participants$`What is your substantive level? - Selected Choice`[Matching_Participants$`What is your substantive level? - Selected Choice` == "L9 and above"] <- 9

Matching_Participants$`What is your substantive level? - Selected Choice` <- as.numeric(Matching_Participants$`What is your substantive level? - Selected Choice`)

Why_Change <- function(dataset){
  dataset <- gsub("Career advancement", "A", dataset)
  
  dataset <- gsub("Leadership development","B", dataset)
  
  dataset <- gsub("Increase confidence","C", dataset)
  
  dataset <- gsub("Build Networks in the Public Sector", "D", dataset)
  
  dataset <- gsub("Build Networks", "D", dataset)
  
  dataset <- gsub("Learn more about working in the Public Sector","E", dataset)
  
  dataset <- gsub("To share my experience","A", dataset)
  
  dataset <- gsub("To help build Leadership in the Public Sector","B", dataset)
  
  dataset <- gsub("Support others to grow","C", dataset)
  
  dataset <- gsub("Learn from and teach them","E", dataset)
  
  dataset <- gsub("To learn to be an effective mentor","E", dataset)
  
  return(dataset)
}

Matching_Participants$`Why do you want to be a mentor? (max 3 responses)` <-  Why_Change(Matching_Participants$`Why do you want to be a mentor? (max 3 responses)`)

Matching_Participants$`Why do you want to be a mentee?(max 2 responses)` <-  Why_Change(Matching_Participants$`Why do you want to be a mentee?(max 2 responses)`)


## ------------------------------------------------------------------------
# This dataset will contain all the entries that are deemed not able to be used for matching as these entries are missing mandetory fields. 
Flagged <- (Matching_Participants %>% filter(Agency == "" | `Full Name` == "" | `I am applying to be a` == "" | `What is your substantive level? - Selected Choice` == ""))

# This dataset is a continuation of the previous dataset for consistency sake and these entries/users have the first level of mandetory fields filled in. 

Matching_Participants <- (Matching_Participants %>% filter(Agency != "" & `Full Name` != "" & `I am applying to be a` != "" & `What is your substantive level? - Selected Choice` != ""))


## ------------------------------------------------------------------------
# This subset of data contains multiple entries from the same person based on their names, might follow-up with them 
Matching_Participants_Duplicates <- Matching_Participants[duplicated(Matching_Participants$`Full Name`),]

# This is the subset of the data with the duplicates removed and we will continue to use this dataset for consistency
Matching_Participants <- Matching_Participants[!duplicated(Matching_Participants$`Full Name`),]


## ------------------------------------------------------------------------
Matching_Participants <- Matching_Participants[!(str_detect(Matching_Participants$`Full Name`, regex("testing", ignore_case = TRUE))),]


## ------------------------------------------------------------------------
# Starting off, we can how many of the participants are applying to be mentor and how many are applying to be mentee. 
Num_Mentor <- sum(Matching_Participants$`I am applying to be a` == "Mentor")
Num_Mentee <- sum(Matching_Participants$`I am applying to be a` == "Mentee")
Num_Agency <- length(unique(Matching_Participants$Agency))
Year <- 2019 

sink("Mentoring Program Dataset Details.txt")

# Below are just some statements that I have made to get more insight on the participants
sprintf("There are %s mentors and %f mentees that can be matched", Num_Mentor, Num_Mentee)

if(Num_Mentee > Num_Mentor){
  print("There is more mentees than mentors, therefore it is possible that not all mentee applicants will be accepted")
} else if(Num_Mentee < Num_Mentor){
  print("There is more mentor than mentee, therefore it is possible that not every mentor applicants will be accepted")
}

# This is just some extra information 
Mentor_2_Mentees <- sum(Matching_Participants$`Are you interested in mentoring more than one mentee?` == "Yes")
sprintf("There is about %s mentors willing to take more than one mentee, assuming that two is the max, they can take %f", Mentor_2_Mentees, Mentor_2_Mentees*2)

sprintf("So in total as of current, we can take %s mentees into the program this round", Mentor_2_Mentees*2 + (Num_Mentor- Mentor_2_Mentees))
sprintf("So it seems that there will be about %s mentees that can't be assigned", (Num_Mentee - (Mentor_2_Mentees*2 + (Num_Mentor - Mentor_2_Mentees))))

sink("Mentoring Program Dataset Details.txt", append = TRUE)
sink()
closeAllConnections()


## ------------------------------------------------------------------------
# This is separating the skills in terms of their preferences first, second and third, the person don't have to have three preferences, but this just makes my automation part easier 
Matching_Participants <-  Matching_Participants %>% separate(`Why do you want to be a mentor? (max 3 responses)`, sep = ",", c("First preference(Mentor Why)", "Second preference(Mentor Why)", "third preference(Mentor Why)"))


## ------------------------------------------------------------------------
Expertise_change <- function(Dataset){
  Dataset <- gsub("Project Management", "A", Dataset)
  
  Dataset <- gsub("Policy Advice", "B", Dataset)
  
  Dataset <- gsub("Ministerial Communication", "C", Dataset)
  
  Dataset <- gsub("Workforce planning/managing employees", "D", Dataset)
  
  Dataset <- gsub("Other", "E", Dataset)
  
  return(Dataset)
}

Skill_Develope_change <- function(Dataset){
  Dataset <- gsub("Managing Projects", "A", Dataset)
  
  Dataset <- gsub("Policy writing/ advice", "B", Dataset)
  
  Dataset <- gsub("Communicating with Ministers or Senior staff", "C", Dataset)
  
  Dataset <- gsub("Managing others", "D", Dataset)
  
  Dataset <- gsub("Other", "E", Dataset)
  
  return(Dataset)
}

Matching_Participants$`What are your areas of expertise in the Public Sector? - Selected Choice` <-  sapply(Matching_Participants$`What are your areas of expertise in the Public Sector? - Selected Choice`, Expertise_change)

Matching_Participants$`What Public Sector Experience are you seeking - Selected Choice` <- sapply(Matching_Participants$`What Public Sector Experience are you seeking - Selected Choice`, Skill_Develope_change)

Matching_Participants <- Matching_Participants %>% separate(`What Public Sector Experience are you seeking - Selected Choice`, sep = ",", c("Most seeking experience","Second most seeking experience"))


## ------------------------------------------------------------------------
Answer_Change <- function(Dataset){
  Dataset <- gsub("Shape and Manage Strategy - supporting purpose and direction, thinking strategically, harnessing information and opportunities, showing judgement", "A", Dataset)
  
  Dataset <- gsub("Achieve Results - Identifies and uses resources wisely, applies and builds professional expertise, responds positively to change, takes responsibility for managing projects to achieve results.", "B", Dataset)
  
  Dataset <- gsub("Exemplifies Personal Integrity and Self Awareness - Demonstrates public service professionalism and probity, Engages with risk and shows personal courage, Commits to action, Promotes and adopts a positive and balanced approach to work, Demonstrates self awareness and a commitment to personal development.", "C", Dataset)
  
  Dataset <- gsub("Builds Productive Relationships - Nurtures internal and external relationships, Listens to, understands and recognises the needs of others, Values individual differences and diversity, Shares learning and supports others", "D", Dataset)
  
  Dataset <- gsub("Communicate and Influence Effectively - Communicates clearly, Listens, understands and adapts to audience, Negotiates confidently.", "E", Dataset)
  
  return(Dataset)
}

Matching_Participants$`Which of these Public Sector capabilities are your strengths? (max 2 responses)` <-  sapply(Matching_Participants$`Which of these Public Sector capabilities are your strengths? (max 2 responses)`, Answer_Change)

Matching_Participants$`Which specific Public Sector Capabilities would you like to improve (max 2 responses)` <- sapply(Matching_Participants$`Which specific Public Sector Capabilities would you like to improve (max 2 responses)`, Answer_Change)

Matching_Participants <- Matching_Participants %>% separate(`Why do you want to be a mentee?(max 2 responses)`, sep = ",", c("First Reason to be Mentee", "Second Reason to be Mentee"))


## ------------------------------------------------------------------------
# Splitting the dataset even more 
Matching_Participants <- Matching_Participants %>% separate(`Which of these Public Sector capabilities are your strengths? (max 2 responses)`, sep = ",", c("Most capable strength", "Second capable strength"))

Matching_Participants <- Matching_Participants %>% separate(`Which specific Public Sector Capabilities would you like to improve (max 2 responses)`, sep = ",", c("Most want improved skill", "Second most want improved skill"))


## ------------------------------------------------------------------------
Mentor <- Matching_Participants[(Matching_Participants$`I am applying to be a` == "Mentor"),]
Mentor <- Mentor[,1:20]

Mentee <- Matching_Participants[(Matching_Participants$`I am applying to be a` == "Mentee"),]
Mentee <- Mentee[,-12:-18]


## ------------------------------------------------------------------------
Unmatched_Mentee <- nrow(Mentee)
Unmatched_Mentor <- nrow(Mentor)
Level_1_2 <- Mentee %>% tally(Mentee$`What is your substantive level? - Selected Choice` == 1.5)
Level_3_4 <- Mentee %>% tally(Mentee$`What is your substantive level? - Selected Choice` == 3.5)
Level_1_4 <-  Level_1_2 + Level_3_4
Level_5 <- Mentee %>% tally(Mentee$`What is your substantive level? - Selected Choice` == 5)
Level_6 <- Mentee %>% tally(Mentee$`What is your substantive level? - Selected Choice` == 6)
Level_7 <- Mentee %>% tally(Mentee$`What is your substantive level? - Selected Choice` == 7)
Level_8 <- Mentee %>% tally(Mentee$`What is your substantive level? - Selected Choice` == 8)
Level_9 <- Mentee %>% tally(Mentee$`What is your substantive level? - Selected Choice` == 9)
Male <- Matching_Participants %>% tally(Matching_Participants$Gender == "Male")
Female <- Matching_Participants %>% tally(Matching_Participants$Gender == "Female")


## ------------------------------------------------------------------------
# Add a column 
Mentee$Similarity_Score <- 0


## ------------------------------------------------------------------------
Data_Filtering <- function(Mentor, Mentee, People){
  Mentor_Agency <- Mentor[People,5]
  Mentor_Level <- Mentor[People,10]
  Filtered <- (Mentee %>% filter(Agency != Mentor_Agency &`What is your substantive level? - Selected Choice` < Mentor_Level))
  
  return(Filtered)
}


## ------------------------------------------------------------------------
Mentee_Delete <- function(Mentee, Mentee.name){
  Mentee <- Mentee[!(Mentee$`Full Name` == Mentee.name),]
  
  return(Mentee)
}

Mentor_Delete <- function(Mentor, Mentor.name){
  Mentor <- Mentor[!(Mentor$`Full Name` == Mentor.name),]
  
  return(Mentor)
}


## ------------------------------------------------------------------------
# First lets get like the number of mentors vs the number of mentees 
Mentor_vs_Mentee <- cbind(Num_Mentor, Num_Mentee)
write.csv(Mentor_vs_Mentee, "Mentor vs Mentee.csv")
# The next step is to get the different departments counts 
Mentor_Agency <- Mentor %>% group_by(Agency) %>% tally()
write.csv(Mentor_Agency, "Mentor Agency.csv")
Mentee_Agency <- Mentee %>% group_by(Agency) %>% tally()
write.csv(Mentee_Agency, "Mentee Agency.csv")
# Then we will get the data for the levels of mentor and mentees 
Mentor_Level <- Mentor %>% count(`What is your substantive level? - Selected Choice`)
write.csv(Mentor_Level, "Mentor Level.csv")
Mentee_Level <- Mentee %>% count(`What is your substantive level? - Selected Choice`)
write.csv(Mentee_Level, "Mentee Level.csv")


## ------------------------------------------------------------------------
Similarity_Score_Calculation <- function(Loop_Data, Mentor, Mentee, People){
  Mentor_Name <- Mentor[People,1]
  Mentor_Agency <- Mentor[People,5]
  Mentor_Level <- Mentor[People,10]
  Mentor_First_Preference <- Mentor[People,12]
  Mentor_Second_Preference <- Mentor[People,13]
  Mentor_Strength <- Mentor[People,15]
  Mentor_Second_Strength <- Mentor[People,16]
  Mentor_Expertise <- Mentor[People,17]
  
  for(i in 1:nrow(Loop_Data)){
  if(is.na(Loop_Data[i,14]) | is.na(Mentor_First_Preference)){
    next
  } else{
    if(Loop_Data[i,14] == Mentor_First_Preference){
      Loop_Data$Similarity_Score[i] = Loop_Data$Similarity_Score[i] + 1
    }
  }
  }
  
  for(i in 1:nrow(Loop_Data)){
  if(is.na(Loop_Data[i,15])| is.na(Mentor_Second_Preference)){
    next
  } else{
    if(Loop_Data[i,15] == Mentor_Second_Preference){
      Loop_Data$Similarity_Score[i] = Loop_Data$Similarity_Score[i] + 1
    }
  }
  }
  
  for(i in 1:nrow(Loop_Data)){
  if(is.na(Loop_Data[i,16])| is.na(Mentor_Strength)){
    next
  } else{
    if(Loop_Data[i,16] == Mentor_Strength){
      Loop_Data$Similarity_Score[i] = Loop_Data$Similarity_Score[i] + 1
    }
  }
  }
  
  for(i in 1:nrow(Loop_Data)){
  if(is.na(Loop_Data[i,17])| is.na(Mentor_Second_Strength)){
    next
  } else{
    if(Loop_Data[i,17] == Mentor_Second_Strength){
      Loop_Data$Similarity_Score[i] = Loop_Data$Similarity_Score[i] + 1
    }
  }
  }
  
  for(i in 1:nrow(Loop_Data)){
  if(is.na(Loop_Data[i,18]) | is.na(Mentor_Expertise)){
    next
  } else{
    if(Loop_Data[i,18] == Mentor_Expertise){
      Loop_Data$Similarity_Score[i] = Loop_Data$Similarity_Score[i] + 1
    }
  }
  }
  
  Loop_Data[order(Loop_Data$Similarity_Score, decreasing = TRUE),]
  return(Loop_Data)
}


## ------------------------------------------------------------------------
Mentor_Matched <- data.frame(matrix(ncol=0,nrow=0))
Mentee_Matched <- data.frame(matrix(ncol=0,nrow=0))


## ------------------------------------------------------------------------
while(nrow(Mentor) != 0){
for(i in 1:nrow(Mentor)){
  A = Data_Filtering(Mentor, Mentee, People = i)
  if(nrow(A) == 0){
    next
  } else{
    if(nrow(A) != 0){
      B = Similarity_Score_Calculation(A, Mentor, Mentee, i)
      B = B[order(B$Similarity_Score, decreasing = TRUE),]
      Mentee_Name = B$`Full Name`[1]
      MentorName =  Mentor$`Full Name`[i]
      Mentor_Matched <- rbind(Mentor_Matched, Mentor[i,])
      Mentee_Matched <- rbind(Mentee_Matched, B[1,])
      Mentor <- Mentor_Delete(Mentor, Mentor.name = MentorName)
      Mentee <- Mentee_Delete(Mentee, Mentee.name = Mentee_Name)
    }
  }
}
}


## ------------------------------------------------------------------------
# Please run this chunk to get the excel csv of the matched pairs. 
Matched_Mentee_Name <- Mentee_Matched$`Full Name`
Matched_Mentor_Name <- Mentor_Matched$`Full Name`
Matched_Mentee_Level <- Mentee_Matched$`What is your substantive level? - Selected Choice`
Matched_Mentor_Level <- Mentor_Matched$`What is your substantive level? - Selected Choice`
Matched_Mentee_Agency <- Mentee_Matched$Agency
Matched_Mentor_Agency <- Mentor_Matched$Agency
Matched_Mentee_Position <- Mentee_Matched$Position
Matched_Mentor_Position <- Mentor_Matched$Position

Matched_Dataset <- cbind(Matched_Mentee_Name, Matched_Mentor_Name, Matched_Mentee_Level, Matched_Mentor_Level, Matched_Mentee_Agency, Matched_Mentor_Agency, Matched_Mentee_Position, Matched_Mentor_Position )
write.csv(Matched_Dataset, "Matched_Pairs.csv")


## ------------------------------------------------------------------------
Mentee_Mentor_Gender <- table(Matching_Participants$Gender, Matching_Participants$`I am applying to be a`)
write.csv(Mentee_Mentor_Gender, "Gender Ratio.csv")

# Matched vs unmatched mentees 
Mentee_Matches <- nrow(Mentee_Matched)
Mentee_UnMatched <- nrow(Mentee)
Matched_vs_unMatched <- cbind(Mentee_Matches, Mentee_UnMatched)
write.csv(Matched_vs_unMatched, "Matched vs Unmatched.csv")

# Gender Ratio based on agency 
Gender_And_Ratio <- table(Matching_Participants$Gender, Matching_Participants$Agency)
write.csv(Gender_And_Ratio, "Gender and Agency.csv")

# Level and Agency 
Level_And_Agency <- table(Matching_Participants$Agency, Matching_Participants$`What is your substantive level? - Selected Choice`)
write.csv(Level_And_Agency, "Level and Agency.csv")


## ------------------------------------------------------------------------
write.xlsx(Matched_Dataset, file = "Mentor-Mentee.xlsx", sheetName = "Matched Pairs", append = TRUE)

if(nrow(Mentor) == 0){
  write.xlsx(Mentee, file = "Mentor-Mentee.xlsx", sheetName = "Unmatched Mentees", append = TRUE)
} else if(nrow(Mentee) == 0){
  write.xlsx(Mentor, file = "Mentor-Mentee.xlsx", sheetName = "Unmatched Mentors", append = TRUE)
}

Stats <- cbind(Year, Num_Mentor, Num_Mentee, Num_Agency, Unmatched_Mentor, Unmatched_Mentee, Level_1_2, Level_3_4, Level_5, Level_6, Level_7, Level_8, Level_9, Male, Female)
colnames(Stats) <- c("Year","Mentors","Mentees","Agencies","Unmatched Mentors","Unmatched Mentees","Level 1-2","Level 3-4", "Level 5","Level 6","Level 7","Level 8", "Level 9 or above","Male","Female")
write.xlsx(Stats, file = "Mentor-Mentee.xlsx", sheetName = "Stats", append = TRUE)

