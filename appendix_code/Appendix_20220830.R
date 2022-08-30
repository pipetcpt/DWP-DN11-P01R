#memory.limit(size=10000000024)
#rm(list=ls())

if(!require(officer))install.packages('officer'); library(officer) # 요 형태로 이용해 주세요
if(!require(flextable))install.packages('flextable'); library(flextable) # 요 형태로 이용해 주세요
if(!require(tidyverse))install.packages('tidyverse'); library(tidyverse) # 요 형태로 이용해 주세요
if(!require(sas7bdat))install.packages('sas7bdat'); library(sas7bdat) # 요 형태로 이용해 주세요
if(!require(xlsx))install.packages('xlsx'); library(xlsx) # 요 형태로 이용해 주세요
if(!require(maditr))install.packages('maditr'); library(maditr) # 요 형태로 이용해 주세요
if(!require(png))install.packages('png'); library(png) # 요 형태로 이용해 주세요
if(!require(knitr))install.packages('knitr'); library(knitr) # 요 형태로 이용해 주세요
if(!require(R.utils))install.packages('R.utils'); library(R.utils) # 요 형태로 이용해 주세요
if(!require(rvg))install.packages('rvg'); library(rvg) # 요 형태로 이용해 주세요
if(!require(reshape2))install.packages('reshape2'); library(reshape2) # 요 형태로 이용해 주세요
if(!require(chron))install.packages('chron'); library(chron) # 요 형태로 이용해 주세요
if(!require(magrittr))install.packages('magrittr'); library(magrittr) # 요 형태로 이용해 주세요
if(!require(icesTAF))install.packages('icesTAF'); library(icesTAF) # 요 형태로 이용해 주세요
if(!require(moonBook))install.packages('moonBook'); library(moonBook) # 요 형태로 이용해 주세요
if(!require(haven)) install.packages("haven");library(haven) # 요 형태로 이용해 주세요


# Save CSV Files
list.files('sasdata')
#### Error 가 있을 때 진행하지 않는 구조임 - 바꿔야함
subDir  <- "SAS2CSV1"
    sas.files <- list.files('sasdata')[grepl(".sas7bdat", list.files('sasdata'))]      ### sas data folder 따로 만드는게 좋을 듯..
    # dir.create(file.path(subDir))
    for(files_ in sas.files){
        name_ <- toupper(strsplit(files_, ".", fixed = TRUE)[[1]][1])
        outfile <- write.csv(as.data.frame(read_sas(paste0('sasdata/',files_)), NULL), 
                             file = paste0(file.path(subDir), "/", name_, ".csv"))
    }


######### FUNCTIONS #########
source('function/count.mean.sd.r')
source('function/createtable.r')
source('function/flextable.r')
source('function/others.r')

    # VISIT 한국어 코드도 항상 코드화하여 사용하기 
# ======== Starting From Here ===========
VST <- rbind.data.frame(
    c(1,"Scr" ),
    c(2,"RN" ),
    c(3,"P1" ),
    c(4,"P1_Day -1" ),
    c(5,"P1_Day 1" ),
    c(6,"P1_Day 1_4hr" ),
    c(7,"P1_Day 1_8hr" ),
    c(8,"P1_Day 1_12hr" ),
    c(9,"P1_Day 2" ),
    c(10,"P1_Day 3" ),
    c(11,"P1_Day 4" ),
    c(12,"P1_Day 5" ),
    c(13,"P1_Day 6" ),
    c(14,"P1_Day 7" ),
    c(15,"P1_Day 8" ),
    c(16,"P1_Day 9" ),
    c(17,"P1_Day 10" ),
    c(18,"P1_Day 11" ),
    c(19,"P2" ),
    c(20,"P2_Day 28" ),
    c(21,"P_Day 29" ),
    c(22,"P2_Day 29_4hr" ),
    c(23,"P2_Day 29_8hr" ),
    c(24,"P2_Day 29_12hr" ),
    c(25,"P2_Day 30" ),
    c(26,"P2_Day 31" ),
    c(27,"P2_Day 32" ),
    c(28,"P2_Day 33" ),
    c(29,"P2_Day 34" ),
    c(30,"P2_Day 35" ),
    c(31,"P2_Day 36" ),
    c(32,"P2_Day 37" ),
    c(33,"P2_Day 38" ),
    c(34,"P2_Day 39" ),
    c(35,"PSV" ),
    c(36,"AE" ),
    c(37,"CM" ),
    c(38,"VP" ),
    c(39,"CC" ),
    c(40,"UV" ),stringsAsFactors = FALSE)


colnames(VST) <- c("VISIT_CODE", "VISIT_LABEL")
VST$VISIT_CODE <- as.integer(VST$VISIT_CODE)
VST$VISIT_LABEL <- as.character(VST$VISIT_LABEL)

list.files(getwd())

# Load CSV Files
data.path <- paste0(getwd(), "/sasdata")
data.files <- list.files(data.path)
data.files



#SUBJID load #update 22.02.03
subid <- read.xlsx("data/DB/DTC21IP065_AnalysisSet.xlsx",1)
subid$SID <- as.character(subid$SID)
subid$RID <- as.character(subid$RID)
str(subid)
subid <- subid %>% select(SID, RID)

#16.2.1

DS0 <- as.data.frame(read_sas(paste0(data.path, "/", data.files[grepl("\\<ds3\\>", data.files)]), NULL))
DM0 <- as.data.frame(read_sas(paste0(data.path, "/", data.files[grepl("\\<dm\\>", data.files)]), NULL))

DS <- DS0 %>%
    filter(DSTERM_STD == 2) %>%
    mutate(
        Final_visit = DSENDTC_RAW, SUBJID = subjectId,
        Reason_for_discontinuation = case_when(DSDECOD_STD == 1 ~ "??ũ???? Ż??",
                                               DSDECOD_STD == 2 ~ "???????? ??????",
                                               DSDECOD_STD == 3 ~ "?????????ڰ? ?ӻ??????? ?Ǿ?ǰ?? ???? ?ߴ?�� ?䱸?ϰų?, ???????? ???Ǹ? öȸ?ϴ? ????",
                                               DSDECOD_STD == 4 ~ "?̻?????��?? ???Ͽ? ?????ڰ? ????�� ?????? ?? ???ٰ? ?Ǵ??ϴ? ????",
                                               DSDECOD_STD == 5 ~ "?????????ڰ? ?ӻ??????? ?Ǿ?ǰ?? ???????̳? ?ൿ???? Ư??�� ?????ϴµ? ????�� ?? ??��?? ?????Ǵ? ?๰�� ???Ƿ? ?????ϴ? ????",
                                               DSDECOD_STD == 6 ~ "?ӻ????? ?? ??��/��?? ???? ?? ?ߴ??? ??ȹ?? ��?ݻ????? ???Ӱ? ?߰ߵǴ? ??????",
                                               DSDECOD_STD == 7 ~ "??Ÿ ??��?? ???Ͽ? ?????ڰ? ????�� ?????ؾ? ?Ѵٰ? ?Ǵ??? ????")) %>%
    select(SUBJID, Reason_for_discontinuation, Final_visit) %>%
    arrange(SUBJID)

DM <- DM0 %>%
    dplyr::rename(SUBJID = subjectId,
                  Sex = SEX_STD,
                  Age = AGE) %>%
    mutate(Sex = case_when(Sex == 1 ~ '남',
                           Sex == 2 ~ '여')) %>%
    select(SUBJID, Sex, Age)

data <- merge(DS, DM, by="SUBJID")
data <- data[,c("SUBJID", 
                "Sex", 
                "Age", 
                "Final_visit", 
                "Reason_for_discontinuation")]

MyFTable_16.2.1 <- flex.table.fun(data)


# 16.2.4.1
SV0 <- as.data.frame(read_sas(paste0(data.path, "/", data.files[grepl("\\<vs2\\>", data.files)]), NULL))
SU0 <- as.data.frame(read_sas(paste0(data.path, "/", data.files[grepl("\\<su\\>", data.files)]), NULL))

SU <- SU0 %>%
    mutate(SUBJID = subjectId,
        SUYN1_STD = ifelse(SUYN1_STD == 1, "예", "아니오"),
        SUYN2_STD = ifelse(SUYN2_STD == 1, "??", "?ƴϿ?"),
        SUYN3_STD = ifelse(SUYN3_STD == 1, "??", "?ƴϿ?")) %>%
    select(SUBJID, SUDTC_RAW, SUYN1_STD, SUYN2_STD, SUYN3_STD ) %>%
    arrange(SUBJID)

SV <- SV0 %>%
    mutate(SUBJID = subjectId) %>%
    select(SUBJID, HEIGHT, WEIGHT ) %>%
    arrange(SUBJID)



DMSV <- merge(DM, SV, by = c("SUBJID")) %>%
    merge(., SU, by = c("SUBJID"))


DMSV <- DMSV[c("SUBJID", 
               "SUDTC_RAW",
               "Age",
               "SUYN1_STD",
               "SUYN2_STD",
               "SUYN3_STD",
               "HEIGHT",
               "WEIGHT",
               "Sex")]

colnames(DMSV)=c("SUBJID", 
                 "Screening Date",
                 "Age",
                 "Smoking", 
                 "Alcohol",
                 "Caffeine",
                 "Height", 
                 "Weight", 
                 "Sex")

MyFTable_16.2.4.1 <- flex.table.fun(DMSV)


################################################################################################################
###################################  Park maria T part.  #######################################################
################################################################################################################


# 16.2.5.1 # ????

EX0 <- as.data.frame(read_sas(paste0(data.path, "/", data.files[grepl("\\<ex\\>", data.files)]), NULL))
EX02 <- as.data.frame(read_sas(paste0(data.path, "/", data.files[grepl("\\<ex2\\>", data.files)]), NULL))


EX0 <- EX0 %>% mutate(SUBJID = subjectId, newVISIT = paste(FolderName, EXTRT, sep="_"))
EX02 <- EX02 %>% mutate(SUBJID = subjectId, newVISIT = paste(FolderName, EXTRT, sep="_"))

data0 <- dcast(EX0, SUBJID ~ factor(newVISIT, levels = str_sort(unique(EX0$newVISIT), numeric = T)), value.var = "EXSTTIM")
data02 <- dcast(EX02, SUBJID ~ factor(newVISIT, levels = str_sort(unique(EX02$newVISIT), numeric = T)), value.var = "EXSTTIM")
EX <- merge(data0, data02, by = c("SUBJID")) 
head(EX)
MyFTable_16.2.5.1 <- flex.table.fun(EX)


#16.2.6.1  # ?ൿ?? ä??
# PC
PC1 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc1\\>", data.files)], NULL))
PC2 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc2\\>", data.files)], NULL))
PC3 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc3\\>", data.files)], NULL))
PC4 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc4\\>", data.files)], NULL))
PC5 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc5\\>", data.files)], NULL))
PC6 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc6\\>", data.files)], NULL))
PC7 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc7\\>", data.files)], NULL))
PC8 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc8\\>", data.files)], NULL))
PC9 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc9\\>", data.files)], NULL))
PC10 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc10\\>", data.files)], NULL))
PC11 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc11\\>", data.files)], NULL))

PC01_1 <- PC1 %>%
    filter(FolderName == "Period 1 Day 1") %>%
    mutate(PCTPT = str_replace( PCTPT, "_-60", "")) %>%
    mutate(SUBJID = subjectId,
         newVISIT = paste(FolderName, PCTPT_STD, sep="_")) %>%
        select(SUBJID, newVISIT, PCTIM)
DPC1_1 <- dcast(PC01_1, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC01_1$newVISIT), numeric = T)), value.var = "PCTIM")
    
PC01_2 <- PC1 %>%
    filter(FolderName == "Period 2 Day 29") %>%
    mutate(PCTPT = str_replace( PCTPT, "_-60", "")) %>%
    mutate(SUBJID = subjectId,
           newVISIT = paste(FolderName, PCTPT_STD, sep="_")) %>%
    select(SUBJID, newVISIT, PCTIM)
DPC1_2 <- dcast(PC01_2, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC01_2$newVISIT), numeric = T)), value.var = "PCTIM")

    
PC02 <- PC2 %>%
        mutate(SUBJID = subjectId,
               newVISIT = paste(FolderName, PCTPT, sep="_")) %>%
        select(SUBJID, newVISIT, PCTIM)
DPC2 <- dcast(PC02, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC02$newVISIT), numeric = T)), value.var = "PCTIM")

PC03 <- PC3 %>%
    mutate(SUBJID = subjectId,
           newVISIT = paste(FolderName, PCTPT, sep="_")) %>%
    select(SUBJID, newVISIT, PCTIM)
DPC3 <- dcast(PC03, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC03$newVISIT), numeric = T)), value.var = "PCTIM")

PC23 <- merge(DPC2, DPC3, by = c("SUBJID")) 


PC04 <- PC4 %>%
    mutate(SUBJID = subjectId,
           newVISIT = paste(FolderName, PCTPT, sep="_")) %>%
    select(SUBJID, newVISIT, PCTIM)
DPC4 <- dcast(PC04, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC04$newVISIT), numeric = T)), value.var = "PCTIM")

PC05 <- PC5 %>%
    mutate(SUBJID = subjectId,
           newVISIT = paste(FolderName, PCTPT, sep="_")) %>%
    select(SUBJID, newVISIT, PCTIM)
DPC5 <- dcast(PC05, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC05$newVISIT), numeric = T)), value.var = "PCTIM")

PC45 <- merge(DPC4, DPC5, by = c("SUBJID")) 


PC06 <- PC6 %>%
    mutate(SUBJID = subjectId,
           newVISIT = paste(FolderName, PCTPT, sep="_")) %>%
    select(SUBJID, newVISIT, PCTIM)
DPC6 <- dcast(PC06, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC06$newVISIT), numeric = T)), value.var = "PCTIM")

PC07 <- PC7 %>%
    mutate(SUBJID = subjectId,
           newVISIT = paste(FolderName, PCTPT, sep="_")) %>%
    select(SUBJID, newVISIT, PCTIM)
DPC7 <- dcast(PC07, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC07$newVISIT), numeric = T)), value.var = "PCTIM")

PC67 <- merge(DPC6, DPC7, by = c("SUBJID")) 

PC08 <- PC8 %>%
    mutate(SUBJID = subjectId,
           newVISIT = paste(FolderName, PCTPT, sep="_")) %>%
    select(SUBJID, newVISIT, PCTIM)
DPC8 <- dcast(PC08, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC08$newVISIT), numeric = T)), value.var = "PCTIM")

PC09 <- PC9 %>%
    mutate(SUBJID = subjectId,
           newVISIT = paste(FolderName, PCTPT, sep="_")) %>%
    select(SUBJID, newVISIT, PCTIM)
DPC9 <- dcast(PC09, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC09$newVISIT), numeric = T)), value.var = "PCTIM")

PC89 <- merge(DPC8, DPC9, by = c("SUBJID")) 

PC010 <- PC10 %>%
    mutate(SUBJID = subjectId,
           newVISIT = paste(FolderName, PCTPT, sep="_")) %>%
    select(SUBJID, newVISIT, PCTIM)
DPC10 <- dcast(PC010, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC010$newVISIT), numeric = T)), value.var = "PCTIM")

PC011 <- PC11 %>%
    mutate(SUBJID = subjectId,
           newVISIT = paste(FolderName, PCTPT, sep="_")) %>%
    select(SUBJID, newVISIT, PCTIM)
DPC11 <- dcast(PC011, SUBJID ~ factor(newVISIT, levels = str_sort(unique(PC011$newVISIT), numeric = T)), value.var = "PCTIM")

PC1011 <- merge(DPC10, DPC11, by = c("SUBJID")) 


MyFTable_16.2.6.1.1 <- flex.table.fun(DPC1_1)
MyFTable_16.2.6.1.2 <- flex.table.fun(DPC1_2)
MyFTable_16.2.6.1.3 <- flex.table.fun(PC23)
MyFTable_16.2.6.1.4 <- flex.table.fun(PC45)
MyFTable_16.2.6.1.5 <- flex.table.fun(PC67)
MyFTable_16.2.6.1.6 <- flex.table.fun(PC89)
MyFTable_16.2.6.1.7 <- flex.table.fun(PC1011)



################################################################################################################
################################################################################################################
################################################################################################################



# 16.2.7
AE0 <- as.data.frame(read_sas(data_file = data.files[grepl("\\<ae\\>", data.files)], NULL)) 

AE <- AE0 %>%
    filter(AELDNM_STD != 3) %>%
    mutate(SUBJID = subjectId, RNNUM = Subject,
           AESEV = case_when(AESEV_STD == 1 ~ "????(Mild)",
                             AESEV_STD == 2 ~ "?ߵ???(Moderate)",
                             AESEV_STD == 3 ~ "????(Severe)"),
           AESER = case_when(AESER_STD == 1 ~ "?ߴ????? ??��(Non-Serious)",
                             AESER_STD == 2 ~ "?ߴ???(Serious)"),
           AEOUT = case_when(AEOUT_STD == 1 ~ "ȸ????/?ذ???(Recovered/Resolved)",
                             AEOUT_STD == 2 ~ "ȸ??????/?ذ?????(Recovering/Resolving)",
                             AEOUT_STD == 3 ~ "ȸ?????? ??��/?ذ????? ??��(Not recovered/Not resolved)",
                             AEOUT_STD == 4 ~ "ȸ???ذ??Ǿ?��?? ??��???? ??��(Recovered/Resolved with sequelae)",
                             AEOUT_STD == 5 ~ "????(Fatal)",
                             AEOUT_STD == 6 ~ "?? ?? ??��(Unknown)"),
           AEREL = case_when(AEREL_STD == 1 ~ "???ü??? ??��(Related)",
                             AEREL_STD == 2 ~ "???ü??? ??��(Not related)",
                             AEREL_STD == 3 ~ "?򰡺Ұ???(Unassessable, Unclassifiable)"),
           AEACN = case_when(AEACN_STD == 1 ~ "????????(Drug Interrupted)",
                             AEACN_STD == 2 ~ "????(Dose Reduced)",
                             AEACN_STD == 3 ~ "????(Dose Increased)",
                             AEACN_STD == 4 ~ "?뷮??ȭ ??��(Dose Maintained)",
                             AEACN_STD == 5 ~ "?? ?? ??��(Unknown)",
                             AEACN_STD == 6 ~ "?ش????? ??��(Not Applicable)"),
           AEACNOTH = case_when(AEACNOTH_STD == 1 ~ "??��(Not Done)",
                                AEACNOTH_STD == 2 ~ "?๰????(Medication)",
                                AEACNOTH_STD == 3 ~ "??Ÿ(Others)")) %>%
    drop_na(AESEV) %>%
    select(SUBJID,
           RNNUM,
           AETERM,
           AESTDTC_RAW,
           AESTTIM,
           AEENDTC_RAW,
           AEENDTIM,
           AESEV,
           AESER,
           AEOUT,
           AEREL,
           AEACN,
           AEACNOTH) %>%
    arrange(SUBJID) 


colnames(AE) <- c("Subject ID",
                  "Enrollment No.",
                  "?̻????��?",
                  "??????",
                  "???۽ð?",
                  "��????",
                  "��???ð?",
                  "??????",
                  "?ߴ뼺",
                  "????",
                  "?Ǿ?ǰ?????ΰ?????",
                  "?ӻ????????Ǿ?ǰ?? ???? ��ġ",
                  "?̻????��? ???? ��ġ")


MyFTable_16.2.7 <- flex.table.fun(AE)


# 16.4.1 ??��


#16.4.2.1 ~ 16.4.2.7
# Main
# LAB data processing
HM0 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<lab\\>", data.files)], NULL)) 
HM <- HM0 %>% mutate(SUBJID = subjectId, VISIT=FolderName) %>%
    dplyr::select(SUBJID, VISIT, AnalyteName, AnalyteValue) %>% dplyr::rename(TEST = AnalyteName, 
                                                                              RES = AnalyteValue)

#Chemistry lab data 
HM0 <- read.csv('SAS2CSV1/LBCH.csv', stringsAsFactors = F) 
head(CH0)
CH <- CH0 %>% 
    mutate(SUBJID = subjectId) %>%
    dcast(CH0, SUBJID+)
    dplyr::select(SUBJID, VISIT, CHTEST, CHORRES) %>% dplyr::rename(TEST = CHTEST, 
                                                                              RES = CHORRES)
HMCH <- rbind(HM, CH) %>% mutate(TEST = case_when(endsWith(TEST, "-GT") ~ "GGT",
                                                  TRUE ~ as.character(TEST)))


# DB Code  ;  Appendix Code
name_map <- data.frame(rbind(
                             # ?????а˻?
                             c("WBC", "WBC"), 
                             c("RBC", "RBC"),
                             c("HGB", "Hemoglobin"),
                             c("HCT", "Hematocrit"),
                             c("PLAT", "Platelets"), 
                             c("NEUT", "Seg.neutrophils"), 
                             c("LYM", "Lympho."),
                             c("MONO", "Mono."),
                             c("EOS", "Eosino."),
                             c("BASO", "Baso."),
                             c("MCV", "MCV"),
                             c("MCH", "MCH"),
                             c("MCHC", "MCHC"),
                             
                             # ????ȭ?а˻?
                             c("GLUC_C", "Glucose"),
                             c("BUN", "BUN"),
                             c("URATE", "Uric acid"),
                             c("PROT_C", "Total protein"),
                             c("ALB", "Albumin"),
                             c("BILI_C", "Total bilirubin"),
                             c("ALP", "Alk.phosphatase"),
                             c("AST", "AST"),
                             c("ALT", "ALT"),
                             c("GGT", "GGT"),
                             c("LDH", "LDH"),
                             c("CREAT", "Creatinine"),
                             c("SODIUM", "Sodium"),
                             c("K", "Potassium"),
                             c("CL", "Chloride"),
                             c("CA", "Calcium"),
                             c("PHOS", "Phosphorus"),
                             c("CHOL", "Total cholesterol"),
                             c("EGFR", "MDRD-eGFR"),
                             c("CPK", "CPK"),
                             c("TG", "Triglyceride")))


colnames(name_map) <- c("AnalyteName", "Fullname")

MyFTable_16.4.2.1To16.4.2.7 <- list()
filter_var <- "TEST"
value_var <- "RES"


for(i in 1:length(name_map$AnalyteName)){
    MyFTable_16.4.2.1To16.4.2.7[[i]] <- flex.table.fun(
        data = create.table(csvfile    = HM,
                            code_id    = name_map$AnalyteName[i], 
                            fullname   = name_map$Fullname[i],
                            filter_var = filter_var,
                            value_var  = value_var,
                            period_    = NULL,
                            type_      = "numeric"),
        calc_ = TRUE)
    
    print(paste("Creating Table for", name_map$Fullname[i], "on subject code"))
}



BC <- HM %>% 
    filter(VISIT == "Screening") %>%
    filter(TEST == "PT" |
               TEST == "APTT") %>%
    mutate(RES = as.numeric(RES)) %>%
    drop_na(RES) %>%
    select(SUBJID, TEST, RES) %>%
    arrange(SUBJID, TEST)

BC.pivot <- dcast(BC, SUBJID ~ TEST, value.var = "RES")

MyFTable_16.4.2.13 <- flex.table.fun(data = BC.pivot, calc_ = TRUE)


#16.4.2.8 ~ 16.4.2.12
# Main 
name_map <- data.frame(rbind(
                             # ???˻?
                             c("GLUC_U", "Glucose"), 
                             c("COLOR", "Color"),
                             c("LEUKASE", "Leukocyte"),
                             c("BILI_U", "Bilirubin"),
                             c("KETONEBD", "Ketone"),
                             c("SPGRAV", "Specific gravity"),
                             c("OCCBLD", "Occult blood"),
                             c("PH", "PH"),
                             c("PROT_U", "Protein"),
                             c("UROBIL", "Urobilinogen"),
                             c("NITRITE", "Nitrite"),
                             
                             # ???˻?2
                             c("WBC count(UF)", "Microscopy WBC"),
                             c("RBC count(UF)", "Microscopy RBC"))
)
colnames(name_map) <- c("AnalyteName", "Fullname")

MyFTable_16.4.2.8To16.4.2.12 <- list()
filter_var <- "TEST"
value_var <- "RES"

for(i in 1:length(name_map$AnalyteName)){
    MyFTable_16.4.2.8To16.4.2.12[[i]] <- flex.table.fun(
        data = create.table(csvfile    = HM, 
                            code_id    = name_map$AnalyteName[i], 
                            fullname   = name_map$Fullname[i],
                            filter_var = filter_var,
                            value_var  = value_var,
                            period_    = NULL,
                            type_      = FALSE),
        calc_ = FALSE)
    
    print(paste("Table created for", name_map$Fullname[i]))
}




#16.4.2.13

BC <- HM %>% 
    filter(VISIT == "Screening") %>%
    filter(TEST == "PT" |
               TEST == "APTT") %>%
    mutate(RES = as.numeric(RES)) %>%
    drop_na(RES) %>%
    select(SUBJID, TEST, RES) %>%
    arrange(SUBJID, TEST)

BC.pivot <- dcast(BC, SUBJID ~ TEST, value.var = "RES")

MyFTable_16.4.2.13 <- flex.table.fun(data = BC.pivot, calc_ = TRUE)


# 16.4.2.14

SR.LAB <- HM %>%
    filter(TEST == "HBSAG" | 
               TEST == "HCVCLD" | 
               TEST == "HIVAB" | 
               TEST == "Syphilis reagin test") %>%
    select(SUBJID, TEST, RES) %>%
    arrange(SUBJID)

#DB spec check! 
#chaged value 1 -> Negative #update 22.02.03
SR.LAB[SR.LAB==1] <- "Negative"

AnalyteName.pivot <- dcast(SR.LAB, SUBJID ~ TEST, value.var = "RES")

MyFTable_16.4.2.14 <- flex.table.fun(AnalyteName.pivot)


# EXTRA

EX.LAB <- HM %>%
    filter(TEST == "BARB" | 
               TEST == "BNZDZPN" | 
               TEST == "OPIATE" | 
               TEST == "AMPHET" |
               TEST == "COCAINE") %>%
    select(SUBJID, TEST, RES) %>%
    arrange(SUBJID)

AnalyteName.EXLAB <- dcast(EX.LAB, SUBJID ~ TEST, value.var = "RES")

MyFTable_16.4.2.15 <- flex.table.fun(AnalyteName.EXLAB)




#16.4.3
MH0 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<mh\\>", data.files)], NULL))
MH <- MH0 %>%
    mutate(SUBJID = subjectId, PT = MHTERM_PT, SOC = MHTERM_SOC, MHSTDTC = MHDTC_RAW) %>%
    select(SUBJID, MHSTDTC, PT, SOC) %>%
    filter(PT != '')

MyFTable_16.4.3 <- flex.table.fun(MH)



# 16.4.4
IE0 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<ie\\>", data.files)], NULL))

IE <- IE0 %>% 
filter(IETEST_STD != "SCIE") %>%
    mutate(SUBJID = subjectId) %>%
    select(SUBJID, IETEST_STD, IEORRES_STD) %>%
    arrange(SUBJID, IETEST_STD, IEORRES_STD)
IE.pivot <- dcast(IE, SUBJID ~ IETEST_STD, value.var = "IEORRES_STD")



IE.pivot[IE.pivot==1] <- "Yes"
IE.pivot[IE.pivot==2] <- "No"

MyFTable_16.4.4 <- flex.table.fun(IE.pivot) 


# 16.4.5 ??ü ?˻?
PE0 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pe\\>", data.files)], NULL))
PE0 <- PE0 %>% mutate(SUBJID = subjectId)

data <- dcast(PE0, SUBJID ~ factor(FolderName, levels = str_sort(unique(PE0$FolderName), numeric = T)), value.var = 'PEYN_STD') %>% select(starts_with("SUBJID"), starts_with("Screening"), everything())
data[data==1] <- "Normal"
data[data==2] <- "Abnormal"
data[is.na(data)] <- "NA"
MyFTable_16.4.5 <- flex.table.fun(data)


# 16.4.6.1 ~ 16.4.6.5
# Main

VS0 <- as.data.frame(read_sas(data_file = data.files[grepl("\\<vs\\>", data.files)], NULL))

VS0 <- VS0 %>%
    mutate(FolderName = str_replace( FolderName, "(???? ??_4hr)", "_4hr")) %>%
    mutate(FolderName = str_replace( FolderName, "(???? ??_8hr)", "_8hr")) %>%
    mutate(FolderName = str_replace( FolderName, "(???? ??_12hr)", "_12hr"))



# BPSYS
VS.SY <- VS0 %>%
    filter(VSTEST_STD == "SYSBP") %>%
    mutate(SUBJID = subjectId)
VS0.BPSYS <- flex.table.fun(
    dcast(VS.SY, SUBJID ~ factor(FolderName,
                               levels = str_sort(
                                   unique(VS.SY$FolderName), numeric = T
                               )),
          value.var = "VSORRES") %>%
        dplyr::select(SUBJID, starts_with("Screening"), everything()),
    calc_ = TRUE
)



# BPDIA
VS.DI <- VS0 %>%
    filter(VSTEST_STD == "DIABP") %>%
    mutate(SUBJID = subjectId)
VS0.BPSYS <- flex.table.fun(
    dcast(VS.DI, SUBJID ~ factor(FolderName,
                                 levels = str_sort(
                                     unique(VS.DI$FolderName), numeric = T
                                 )),
          value.var = "VSORRES") %>%
        dplyr::select(SUBJID, starts_with("Screening"), everything()),
    calc_ = TRUE
)






# PULSE
VS.PL <- VS0 %>%
    filter(VSTEST_STD == "PULSE") %>%
    mutate(SUBJID = subjectId)
VS0.BPSYS <- flex.table.fun(
    dcast(VS.PL, SUBJID ~ factor(FolderName,
                                 levels = str_sort(
                                     unique(VS.PL$FolderName), numeric = T
                                 )),
          value.var = "VSORRES") %>%
        dplyr::select(SUBJID, starts_with("Screening"), everything()),
    calc_ = TRUE
)



# TEMP
VS.TE <- VS0 %>%
    filter(VSTEST_STD == "TEMP") %>%
    mutate(SUBJID = subjectId)
VS0.BPSYS <- flex.table.fun(
    dcast(VS.TE, SUBJID ~ factor(FolderName,
                                 levels = str_sort(
                                     unique(VS.TE$FolderName), numeric = T
                                 )),
          value.var = "VSORRES") %>%
        dplyr::select(SUBJID, starts_with("Screening"), everything()),
    calc_ = TRUE
)



# 16.4.7 
# Main
EG0 <- as.data.frame(read_sas(data_file=data.files[grepl("\\<eg\\>", data.files)], NULL))
EG <- visit.match.fun(infile = EG0, visitData = VST) %>% filter(EGND == 1)

EG.VR <- EG0 %>%
    filter(EGTEST_STD == "VR") %>%
    mutate(SUBJID = subjectId)
EGVR  <- flex.table.fun(dcast(EG.VR, SUBJID~factor(FolderName, levels=str_sort(unique(EG.VR$FolderName), numeric=T)), value.var = "EGORRES_RAW")%>%
                            select(SUBJID, starts_with("Screening"), everything()), calc_ = TRUE)

EG.PR <- EG0 %>%
    filter(EGTEST_STD == "PR") %>%
    mutate(SUBJID = subjectId)
EGPR  <- flex.table.fun(dcast(EG.PR, SUBJID~factor(FolderName, levels=str_sort(unique(EG.PR$FolderName), numeric=T)), value.var = "EGORRES_RAW")%>%
                            select(SUBJID, starts_with("Screening"), everything()), calc_ = TRUE)

EG.QRS <- EG0 %>%
    filter(EGTEST_STD == "QRS") %>%
    mutate(SUBJID = subjectId)
EGQRS <- flex.table.fun(dcast(EG.QRS, SUBJID~factor(FolderName, levels=str_sort(unique(EG.QRS$FolderName), numeric=T)), value.var = "EGORRES_RAW")%>%
                            select(SUBJID, starts_with("Screening"), everything()), calc_ = TRUE)

EG.QT <- EG0 %>%
    filter(EGTEST_STD == "QT") %>%
    mutate(SUBJID = subjectId)
EGQT  <- flex.table.fun(dcast(EG.QT, SUBJID~factor(FolderName, levels=str_sort(unique(EG.QT$FolderName), numeric=T)), value.var = "EGORRES_RAW")%>%
                            select(SUBJID, starts_with("Screening"), everything()), calc_ = TRUE)

EG.QTC <- EG0 %>%
    filter(EGTEST_STD == "QTC") %>%
    mutate(SUBJID = subjectId)
EGQTC <- flex.table.fun(dcast(EG.QTC, SUBJID~factor(FolderName, levels=str_sort(unique(EG.QTC$FolderName), numeric=T)), value.var = "EGORRES_RAW")%>%
                            select(SUBJID, starts_with("Screening"), everything()), calc_ = TRUE)

ECGNORM <- EG0 %>%
    mutate(SUBJID = subjectId, VISIT = FolderName,EGNOR = EGCLSIG_STD) %>%
    select(SUBJID, VISIT, EGNOR) %>%
    distinct(SUBJID, VISIT, EGNOR) %>%
    mutate(EGNOR = case_when(EGNOR == 1 ~ "��??",
                             EGNOR == 2 ~ "??��??/NCS",
                             EGNOR == 3 ~ "??��??/CS",
                             TRUE ~ as.character(EGNOR)))
    

ECGNORM1 <- dcast(ECGNORM, SUBJID~VISIT, value.var = "EGNOR") %>% select(SUBJID, starts_with("Screening"), everything())
                    
ECGNORM1[is.na(ECGNORM1)] <- "NA"
ECGNORM <- ECGNORM1 %>%
    select(SUBJID, starts_with("Screening"), everything())
ECGNORM <- flex.table.fun(ECGNORM, calc_ = FALSE)





##############################
###       Doc & Table      ###
##############################

doc <- read_docx(path = "C:/Users/cmc/Documents/003_appendix/Template/Template.docx")
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #1
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #2
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #3
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #4
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #5
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #6
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #7
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #8
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #9
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #10
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #11
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #12
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #13
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #14
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #15
doc <- body_add_par(doc, "Format cell text values", style = "heading 1") #16
doc <- body_add_par(doc, "Appendix", style = "heading 1") #16
doc <- body_add_par(doc, "????��??", style = "heading 2") #16.1
doc <- body_add_par(doc, "?ӻ????? ??ȹ??", style = "heading 3") #16.1.1
doc <- body_add_par(doc, "Refer to ISF", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "???ʱ??ϼ? ????", style = "heading 3") #16.1.2
doc <- body_add_par(doc, "Refer to ISF", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "?ӻ????? ?ɻ?��??ȸ?? ???? ?? ?ɻ?????, ???鵿?Ǽ??? ???Ǹ? ��?? ???��� ????", style = "heading 3") #16.1.3
doc <- body_add_par(doc, "Refer to ISF", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "????å???? ?? ??????, ???????????? ???? ?? ???????࿡ ?????? ????", style = "heading 3") #16.1.4
doc <- body_add_par(doc, "Refer to ISF", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "????å???? ?? ?????????? ?Ǵ? ???? ?Ƿ????? ????", style = "heading 3") #16.1.5
doc <- body_add_par(doc, "Refer to section 1.2", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "2???? ?̻??? batch???? ����?? ?ӻ????????Ǿ?ǰ�� ???? ?޾?�� ???? ?? ����??ȣ ???? ?? ?? batch?? ?????? ???????? ????", style = "heading 3") #16.1.6
doc <- body_add_par(doc, "Not Applicable", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "????�� ??�� ???? ?? ??��ǥ", style = "heading 3") #16.1.7
doc <- body_add_par(doc, "Refer to ISF", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "��??Ȯ?μ?", style = "heading 3") #16.1.8
doc <- body_add_par(doc, "Refer to Audit Certificates", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "?????? ?????? ???? ????", style = "heading 3") #16.1.9
doc <- body_add_par(doc, "Refer to statistical analysis plan", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "?? ?ǽñ????? ?????ǰ? ǥ??ȭ ????, ??Ÿ ?ڷ??? ???? ????�� ��?? ?????? ????", style = "heading 3") #16.1.10
doc <- body_add_par(doc, "Refer to ISF", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "?ӻ??????????? ?????Ͽ?�� ???? ???ǹ?", style = "heading 3") #16.1.11
doc <- body_add_par(doc, "Not Applicable", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "?ӻ??????????? ?򰡿? ?????? ????�� ??ģ ???���??", style = "heading 3") #16.1.12
doc <- body_add_par(doc, "Not Applicable", style = "rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "?????????? ?ڷ? ????", style = "heading 2") #16.2

doc <- body_add_par(doc, "?ߵ?Ż????", style = "heading 3")	#16.2.1
doc <- body_add_flextable(doc, MyFTable_16.2.1)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "??????ȹ?? ??Ż", style = "heading 3") #16.2.2
doc <- body_add_par(doc, "Refer to section 10.2", style = "rTableLegend")
doc <- body_add_break(doc)

doc <- body_add_par(doc, "?ൿ?? ?򰡿??? ��?ܵ? ??????????", style = "heading 3") #16.2.3
doc <- body_add_par(doc, "Not Applicable", style = "rTableLegend")
doc <- body_add_break(doc)

doc <- body_add_par(doc, "?????????? Ư??ǥ", style = "heading 3") #16.2.4
doc <- body_add_flextable(doc, MyFTable_16.2.4.1)
doc <- body_add_break(doc)


doc <- body_add_par(doc, "???��? ?? ???߳??? ?ڷ?", style = "heading 3") #16.2.5
doc <- body_add_par(doc, "16.2.5.1 ?????????ں? ???? ???? ?ð?", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.2.5.1)
doc <- body_add_break(doc)


doc <- body_add_par(doc, "16.2.5.2 ?????????ں? ???? ?? ???? (Period 1 Tamsulosin)", style="rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.2.5.3 ?????????ں? ???? ?? ???? (Period 2 Mirabegron)", style="rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.2.5.4 ?????????ں? ???? ?? ???? (Period 3 Tamsulosin)", style="rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.2.5.5 ?????????ں? ???? ?? ???? (Period 3 Mirabegron)", style="rTableLegend")
doc <- body_add_break(doc)

doc <- body_add_par(doc, "?????????ں? ?ൿ?? ????", style = "heading 3") #16.2.6
doc <- body_add_par(doc, "16.2.6.1 ?????????ں? ????ȹ ?м? ???? (Tamsulosin)", style="rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.2.6.2 ?????????ں? ????ȹ ?м? ???? (Mirabegron)", style="rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.2.6.3 ?????????ں? ???߳??? ?��????? (Tamsulosin)", style="rTableLegend")
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.2.6.4 ?????????ں? ???߳??? ?��????? (Mirabegron)", style="rTableLegend")
doc <- body_add_break(doc)

doc <- body_end_section_continuous(doc) # ???? ????
doc <- body_add_par(doc, "16.2.6.5 ?????????ں? ä???????ð?", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.2.6.1.1)
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.2.6.1.2)
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.2.6.1.3)
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.2.6.1.4)
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.2.6.1.5)
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.2.6.1.6)
doc <- body_add_break(doc)


doc <- body_add_par(doc, "?????????ں? ?̻????? ????", style = "heading 3") #16.2.7
doc <- body_add_flextable(doc, MyFTable_16.2.7)
doc <- body_add_break(doc)
doc <- body_end_section_landscape(doc) #??
doc <- body_add_par(doc, "?????ں? ?ӻ??˻? ??��??ġ ?ڷ?", style = "heading 3") #16.2.8
doc <- body_add_break(doc)


doc <- body_add_par(doc, "???ʱ??ϼ?", style = "heading 2") #16.3
doc <- body_add_par(doc, "????, ?Ǵ? ?ٸ? ?ߴ??? ?̻?????, ?ٸ? ��?Ǽ? ?ִ? ?̻????��? ???? ???ʱ??ϼ?", style = "heading 3") #16.3.1
doc <- body_add_par(doc, "Not Applicable", style = "rTableLegend")
doc <- body_add_par(doc, "?̿??? ???ʱ??ϼ?", style = "heading 3") #16.3.2
doc <- body_add_par(doc, "Not Applicable", style = "rTableLegend")
doc <- body_add_par(doc, "?????????ں? ?ڷ? ????", style = "heading 2") # 16.4
doc <- body_add_par(doc, "???? ?๰ ?Ǵ? ???? ?๰", style = "heading 3") # 16.4.1




doc <- body_add_par(doc, "???????? ?˻?", style = "heading 3") 	#16.4.2
doc <- body_add_par(doc, "16.4.2.1 ?????? ?˻?", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[1]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[2]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[3]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[4]])
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.2.2 ?????? ?˻? (????)", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[5]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[6]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[7]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[8]])
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.2.3 ?????? ?˻? (????)", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[9]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[10]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[11]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[12]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[13]])
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.2.4 ????ȭ?а˻?", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[14]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[15]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[16]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[17]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[18]])
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.2.5 ????ȭ?а˻? (????)", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[19]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[20]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[21]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[22]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[23]])
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.2.6 ????ȭ?а˻? (????)", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[24]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[25]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[26]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[27]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[28]])
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.2.7 ????ȭ?а˻? (????)", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[29]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[30]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[31]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[32]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[33]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.1To16.4.2.7[[34]])
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.2.8 ?Һ??˻?",style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[1]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[2]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[3]])
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.4.2.9 ?Һ??˻? (????)", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[4]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[5]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[6]])
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.4.2.10 ?Һ??˻? (????)", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[7]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[8]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[9]])
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.4.2.11 ?Һ??˻? (????)", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[10]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[11]])
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.4.2.12 ?Һ??˻? (????)", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[12]])
doc <- body_add_break(doc)
doc <- body_add_flextable(doc, MyFTable_16.4.2.8To16.4.2.12[[13]])
doc <- body_add_break(doc)


doc <- body_add_par(doc, "16.4.2.13 ?????��??˻?", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.13)
doc <- body_add_break(doc)
doc <- body_add_par(doc, "16.4.2.14 ??û?˻?", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.2.14)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "????", style = "heading 3") 	#16.4.3
doc <- body_add_flextable(doc, MyFTable_16.4.3)
doc <- body_add_break(doc)


doc <- body_add_par(doc, "???????????? ??��", style = "heading 3") 	#16.4.4
doc <- body_add_flextable(doc, MyFTable_16.4.4)
doc <- body_add_break(doc)


doc <- body_add_par(doc, "??ü ?˻?", style = "heading 3") 	#16.4.5
doc <- body_add_par(doc, "16.4.5.1 ??ü ?˻?; ", style="rTableLegend")
doc <- body_add_flextable(doc, MyFTable_16.4.5)
doc <- body_add_break(doc)


doc <- body_add_par(doc, "Ȱ?? ¡??", style = "heading 3") # 16.4.6
doc <- body_add_par(doc, "16.4.6.1 Ȱ??¡??[SBP]", style="rTableLegend")
doc <- body_add_flextable(doc, VS0.BPSYS)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.6.2 Ȱ??¡??[DBP]", style="rTableLegend")
doc <- body_add_flextable(doc, VS0.BPDIA)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.6.5 Ȱ??¡??[PULSE]", style="rTableLegend")
doc <- body_add_flextable(doc, VS0.HRATE)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.6.7 Ȱ??¡??[BODY TEMPERATURE]", style="rTableLegend")
doc <- body_add_flextable(doc, VS0.TEMP)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "12-Lead ??????", style = "heading 3") # 16.4.7
doc <- body_add_par(doc, "16.4.7.1 12-LEAD ??????[Ventricular Rate]", style="rTableLegend")
doc <- body_add_flextable(doc, EGVR)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.7.2 12-LEAD ??????[PR Interval]", style="rTableLegend")
doc <- body_add_flextable(doc, EGPR)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.7.3 12-LEAD ??????[QRS]", style="rTableLegend")
doc <- body_add_flextable(doc, EGQRS)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.7.4 12-LEAD ??????[QT]", style="rTableLegend")
doc <- body_add_flextable(doc, EGQT)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.7.5 12-LEAD ??????[QTC]", style="rTableLegend")
doc <- body_add_flextable(doc, EGQTC)
doc <- body_add_break(doc)

doc <- body_add_par(doc, "16.4.7.6 12-LEAD ??????[Overalll]", style="rTableLegend")
doc <- body_add_flextable(doc, ECGNORM )
doc <- body_add_break(doc)

appendix.name <- strsplit(getwd(), "/")[[1]][4]
print(doc, target = paste0(dirname(getwd()), "/", appendix.name, "_APPENDIX.docx"))


##################################################################################
###############################   STAT ###########################################
##################################################################################

# Table 1 Subject Disposition
info <- read.csv("./../infoList.csv", header = T, sep = ",")
EL <- as.data.frame(read_sas(data_file=data.files[grepl("\\<el\\>", data.files)], NULL))
enrolledID <- as.data.frame(read_sas(data_file=data.files[grepl("\\<rn\\>", data.files)], NULL)) %>% dplyr::select(SUBJID, RNNO)
enrolledEL <- EL %>% filter(SUBJID %in% enrolledID$SUBJID)


# Table 1 Subject Disposition

## Screening summary
# Number of volunteers screened
nrow(DS0)
# Number of subjects enrolled
nrow(DS0%>% filter(FINSTAT %in% c(1,2)))
# Number of volunteers NOT enrolled
nrow(DS0) - nrow(DS0%>% filter(FINSTAT %in% c(1,2)))

## Reasons for screening failure
# Eligibility criteria
nrow(DS0 %>% filter(SFSTAT == 1))
# Consent withdrawal
nrow(DS0 %>% filter(SFSTAT == 2))
# Others
nrow(DS0 %>% filter(SFSTAT == 3))

## Subject allocation
# Enrolled 
nrow(DS0%>% filter(FINSTAT %in% c(1,2)))
# Treated 

# Non-treated

# Completed
nrow(DS0 %>% filter(FINSTAT == 1))
# Dropped-out
nrow(DS0 %>% filter(FINSTAT == 2))

## Reason for drop-out
table(DS0$WDRSTAT)


# Table 2. Subject Demographics & baseline characteristics summary
enrolledID <- as.data.frame(read_sas(data_file=data.files[grepl("\\<rn\\>", data.files)], NULL)) %>% filter(RNND==1) %>% dplyr::select(SUBJID)
# Demographics
table2Demo <- DMSV %>% filter(SUBJID %in% enrolledID$SUBJID)

# AGE
mytable(~ Age, data = table2Demo)            
min(table2Demo$Age) ; max(table2Demo$Age)

# Height
mytable(~ Height, data = table2Demo, digit = 2)
min(table2Demo$Height) ; max(table2Demo$Height)

# Weight
mytable(~ Weight, data = table2Demo, digit = 2)
min(table2Demo$Weight) ; max(table2Demo$Weight)

# Vital Sign
table2Vital <- VS0 %>% filter(SUBJID %in% enrolledID$SUBJID)

# Systolic blood pressure
mytable(~SBP, data = table2Vital, digit = 2)
min(table2Vital$SBP) ; max(table2Vital$SBP)

# Diastolic blood pressure
mytable(~DBP, data = table2Vital, digit = 2)
min(table2Vital$DBP) ; max(table2Vital$DBP)

# Pulse Rate
mytable(~HR, data = table2Vital, digit = 2)
min(table2Vital$HR) ; max(table2Vital$HR)

# Body Temp
mytable(~BT, data = table2Vital, digit = 2)
min(table2Vital$BT) ; max(table2Vital$BT)

# ECG
table2ECG <- EG0 %>% filter(SUBJID %in% enrolledID$SUBJID)

# VR 
mytable(~EGVR, data = table2ECG, digit = 2)
min(table2ECG$EGVR, na.rm = TRUE) ; max(table2ECG$EGVR, na.rm = T)

# PR interval
mytable(~EGPR, data = table2ECG, digit = 2)
min(table2ECG$EGPR, na.rm = TRUE) ; max(table2ECG$EGPR, na.rm = T)

# QRSD
mytable(~EGQRS, data = table2ECG, digit = 2)
min(table2ECG$EGQRS, na.rm = TRUE) ; max(table2ECG$EGQRS, na.rm = T)

# QT
mytable(~EGQT, data = table2ECG, digit = 2)
min(table2ECG$EGQT, na.rm = TRUE) ; max(table2ECG$EGQT, na.rm = T)

# QTc
mytable(~EGQTC, data = table2ECG, digit = 2)
min(table2ECG$EGQTC, na.rm = TRUE) ; max(table2ECG$EGQTC, na.rm = T)


# Table 3 Subject Demographcis cont.
table3MH <- MH0 %>% filter(SUBJID %in% enrolledID$SUBJID)
mytable(~SOC, data= table3MH)
mytable(~Smoking, data= table2Demo) ; mytable(~Alcohol, data= table2Demo) ; mytable(~Caffeine, data= table2Demo)


# lifeStyle 
table(table2Demo$Smoking)
table(table2Demo$Alcohol)
table(table2Demo$Caffeine)

# table3 Adverse event and adverse drug reaction
EX <- EX0 %>% 
    mutate(period = case_when(EXTRT1 == 1 & is.na(EXTRT2) ~ 'period1',
                              is.na(EXTRT1) & EXTRT2 == 1 ~ 'period2',
                              EXTRT1 == 1 & EXTRT2 == 1 ~ 'period3')) %>%
    distinct(SUBJID, period, .keep_all = T) %>%
    drop_na(period)

AE <- read.csv("D:/Appendix/appendix/KDF1901_DDI/AE.csv")

table3AE <- AE %>% 
    filter(SUBJID %in% enrolledID$SUBJID) %>% 
    dplyr::select(SUBJID,period, AESTDTC, PT, SOC, AESER, AESEV, AEREL) %>%
    filter(PT != '')

write.csv(table3AE, "D:/Appendix/appendix/KDF1901_DDI/table3ae.csv", row.names = F)



AEsummary.fun <- function(data) {
    # Calculate the number of subject and number of events 
    # for each AE and ADR for each category
    #
    # Args:
    #	data: Adverse Events table
    #
    # Returns:
    #	Return the list of summarized table for Adverse Events 
    
    if(identical(data, logical(0))) stop('No data file found')
    tryCatch(
        expr = {
            data$SOC <- as.character(data$SOC)
            data$PT <- as.character(data$PT)
            rowname_ <- c(unique(data$SOC), unique(data$PT))
            res_ <- try(as.data.frame(cbind(rowname_)))
            if(class(res_) != 'data.frame') return(-1)
            
            ## count the number of AE 
            # number of total events
            AEevent_ <- nrow(data) 
            # number of total subjects
            AEsubj_ <- length(unique(data$SUBJID)) 
            AE <- paste0(AEsubj_, '[', AEevent_, ']')
            
            ## count the number of ADR
            # number of total events
            ADRevent_ <- nrow(subset(data, AEREL == 1))
            # number of total subjects
            ADRsubj_ <- length(unique(subset(data, AEREL == 1)$SUBJID))
            ADR <- paste0(ADRsubj_, '[', ADRevent_, ']')
            
            # Append to the res_ data frame
            res_[1, 2] <- AE
            res_[1, 3] <- ADR
            
            pt_ <- unique(data$PT)
            
            for (i in 1:length(pt_)) {
                test__ <- try(data %>% filter(PT == pt_[[i]]))
                if(class(data) != 'data.frame') return(-1)
                
                ## count the number of AE 
                # number of total events
                AEevent_ <- nrow(test__)
                # number of total subjects
                AEsubj_ <- length(unique(test__$SUBJID))
                AE <- paste0(AEsubj_, '[', AEevent_, ']')
                
                ## count the number of ADR
                # number of total events
                ADRevent_ <- nrow(subset(test__, AEREL == 1))
                # number of total subjects
                ADRsubj_ <- length(unique(subset(test__, AEREL == 1)$SUBJID))
                ADR <- paste0(ADRsubj_, '[', ADRevent_, ']')
                
                # Append to the next row of res_ data frame
                res_[i + 1, 2] <- AE
                res_[i + 1, 3] <- ADR
                colnames(res_) <- c('SOC/PT', 'AEs', 'ADRs')
            }
            return(res_)
        },
        error = function(e){
            print(e)
        },
        warning = function(w){
            message('Caught an warning!')
            print(w)
        },
        finally = NULL
    )
}


multiple_ <- function(data, group_, totalNum_){
    # Perform AEsummary function multiple times 
    #
    # Args:
    #	data: list of Adverse Events table
    #	group_: group of cohorts used in the clinical trial
    #
    # Returns:
    #	The list of list containing the summarized table of AE
    
    idx <- 1
    AEList <- list()
    groupNum <- totalNum_[-(length(totalNum_))]
    totalNum <- tail(totalNum_, n=1)
    
    tryCatch(
        expr = {
            for(i in 1:length(group_)){
                testA <- data %>% filter(period == group_[[i]])
                socA <- unique(testA$SOC)
                for(j in 1:length(socA)){
                    testAsoc <- testA %>% filter(SOC == socA[[j]])
                    beforeFinal <- AEsummary.fun(testAsoc)
                    beforeFinal$group = as.character(group_[[i]])
                    beforeFinal$totalNum = as.numeric(groupNum[[i]])
                    AEList[[idx]] <- beforeFinal
                    idx <- idx+1
                }
                print(paste0('working on ', group_[i]))
            }
            
            # calculate total
            total_ <- data
            socA <- unique(total_$SOC)
            for(k in 1:length(socA)){
                testAsoc <- total_ %>% filter(SOC == socA[[k]])
                beforeFinal <- AEsummary.fun(testAsoc)
                beforeFinal$group <- 'total'
                beforeFinal$totalNum = as.numeric(totalNum)
                AEList[[idx]] <- beforeFinal
                idx <- idx+1
            }
            print('working on total')
            return(AEList)
        },
        error = function(e){
            print(e)
        },
        finally = NULL
    )
}


# Main
main <- function() {
    # Main function for AE analysis
    group_ <- c(1, 2, 3)
    totalNum <- c(40, 37, 31, 41)
    a <-
        multiple_(data = table3AE,
                  group_ =  group_,
                  totalNum_ = totalNum)
    alist <- do.call(rbind, a)
    
    for (i in 1:nrow(alist)) {
        subj_ <- as.numeric(strsplit(alist$AEs[[i]], '\\[|\\]')[[1]][1])
        event_ <- as.numeric(strsplit(alist$AEs[[i]], '\\[|\\]')[[1]][2])
        total_ <- alist$totalNum[[i]]

        alist$AEsF[[i]] = paste0(
            subj_, '(', round(subj_ / total_ * 100, 1), ')[',event_,']'
        )
        
        subj_ <- as.numeric(strsplit(alist$ADRs[[i]], '\\[|\\]')[[1]][1])
        event_ <- as.numeric(strsplit(alist$ADRs[[i]], '\\[|\\]')[[1]][2])
        alist$ADRsF[[i]] =paste0(
            subj_, '(', round(subj_ / total_ * 100, 1), ')[',event_,']'
        )
    }
    print(alist)
    #return(alist)
}


if (getOption('run.main', default = TRUE)) {
    main()
    #write.csv(main(), "D:/Appendix/appendix/KDF1901_DDI/AElist.csv", row.names = F)
}

