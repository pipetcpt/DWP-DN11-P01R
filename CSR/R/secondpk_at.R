##### table 만들기

library(tidyverse)
library(readxl)
library(cowplot)
library(NonCompart)

data <- read_excel("2. 시료 분석 결과 정리_DTPK-CS-22-014_Rebamipide_draft4_20220715 (읽기전용).xlsx",
                   sheet="Sample Conc. (통계용 전달)")
data <- data[,c(1,2,3,5)]
colnames(data)
colnames(data) <- c("Subject", "Day", "Time", "Conc")

#distinct(data[1])

dropout <- c("R010", "R040")

#distinct(data[3])[9,1]

data1 <- data %>% 
  mutate(Time = ifelse(Time == "F_0", -8, Time)) %>% 
  mutate(Time = parse_number(Time)) %>%
  mutate(Conc = as.numeric(Conc)) %>%  
  mutate(Conc = ifelse(is.na(Conc), 0, Conc)) %>%
  filter(!(Subject %in% dropout)) %>% 
  print()


table1 <- data1 %>% 
  group_by(Day, Time) %>% 
  summarise(Mean = mean(Conc), SD = sd(Conc), Median=median(Conc), Min=min(Conc), Max=max(Conc)) %>% 
  ungroup() %>% 
  mutate("CV(%)" = SD/Mean*100) %>% 
  mutate_at(-(1:2), round, 2) %>% 
  mutate_at(-(1:2), ~sprintf("%0.2f", .x)) %>% 
  unite("Mean ± SD", Mean, SD, sep = " ± ")  %>% 
  select(1:3, 7, everything()) %>% 
  print()


write_csv(table1,"table1.csv")






##### 전체 figure 그리기

title <- ggdraw() + 
  draw_label(
    "Mean Concentration-Time Curves",
    fontface = 'bold',
    x = 0,
    hjust = 0
  ) +
  theme(
    # add margin on the left of the drawing canvas,
    # so title is aligned with left edge of first plot
    plot.margin = margin(0, 0, 0, 7)
  )


mean_curve <- data1 %>% filter(!Time == -8) %>% 
  group_by(Day, Time) %>% 
  mutate(mean = mean(Conc), sd = sd(Conc)) %>% 
  ggplot(aes(x = Time, y = mean, col = Day)) +
  geom_point(size = 2, alpha = 0.5) +
  geom_line(size = 0.5, alpha = 0.5) +
  geom_errorbar(aes(ymin = mean, ymax = mean + sd)) +
  facet_wrap(~ Day) + 
  labs(x = "Time (hr)",
       y = "Concentration (pg/mL)") +
  theme_bw() +
  theme(text=element_text(size = 12)) +
  theme(strip.background = element_rect(fill = NA, colour = NA), 
        legend.position="bottom") +
  ggsci::scale_color_aaas()

plot_row <- plot_grid(mean_curve , mean_curve + scale_y_log10(),
                      labels = c("A (Linear)", "B (Semilog)"), label_size = 12, ncol=1) 

plot_grid(title, plot_row, ncol = 1, rel_heights = c(0.1, 1))

ggsave("mean_curve.png")

#mean_curve + geom_point(aes(x=0, y=300))


##### 개별 figure 그리기

indi_curve <- data1 %>% filter(!Time == -8) %>% 
  ggplot(aes(x = Time, y = Conc, col = Day)) +
  geom_point(size = 1, alpha = 0.5) +  
  geom_line(size = 0.5, alpha = 0.5) +
  facet_wrap(~ Subject + Day, scale = "free_x") +
  theme_bw() +
  theme(strip.background = element_rect(fill = NA, colour = NA), 
        legend.position = "bottom") +
  ggsci::scale_color_aaas()

indi_curve + labs(x = "Time (h)", y = "Concentration (pg/mL)",
                  title = "Individual Concentration-Time Curves (Linear)") 

ggsave("indi_curve_linear.png", width = 15, height = 12)

indi_curve + labs(x = "Time (h)", y = "Concentration (pg/mL)",
                  title = "Individual Concentration-Time Curves (Semilog)") +
  scale_y_log10() 


ggsave("indi_curve_semilog.png", width = 15, height = 12)





##### 실제 채혈 시간 반영하기 - data2

sidrid <- read_excel("DWPDN11_P01R_DataCenter_DataSet_List_20220817084019.xlsx", sheet = "RN") %>% 
  select(SID = SUBJID, Subject = RNNO)
actual.T0 <- read_excel("DWPDN11_P01R_DataCenter_DataSet_List_20220817084019.xlsx", sheet = "PB")
actual.T <- actual.T0 %>% 
  select(SID = SUBJID, Day = VISIT, Time = PBNT, Dev = PBDETM) %>%
  mutate(Day = ifelse(Day == 3, "1d", "5d")) %>% 
  mutate(Time = ifelse(Day == "5d" & Time == "첫 투여 직전 (0 h)", -8, parse_number(Time))) %>% 
  mutate(Dev = as.numeric(Dev)) %>% 
  mutate(aTime = Time + Dev/60) %>% 
  left_join(sidrid) %>%
  print()

data2 <- data1 %>%  
  left_join(actual.T) %>% 
  mutate(aTime = ifelse(Time == 0, 0, aTime)) %>% 
  select(SID, Subject, Day, Time, Dev, aTime, Conc) %>% 
  print()



##### NCA 

data3 <- data2 %>% 
  rbind(data2 %>% filter(Day == "5d" & Time < 6) %>% mutate(Day = "5d.1"))

nca3 <- tblNCA(as.data.frame(data3 %>% filter(!Time == -8)), 
               key = c("SID", "Subject", "Day"),
               colTime = "aTime",
               colConc = "Conc",
               dose = 1.5,
               doseUnit = "mg",
               concUnit = "pg/mL")



table3 <- nca3 %>%
  select(Day, Subject, AUCLST, CMAX, AUCIFO, TMAX, LAMZHL, CLFO, VZFO) %>%
  gather(param, value, AUCLST:VZFO) %>%
  na.omit() %>% 
  group_by(Day, param) %>%
  summarise(N=n(), Mean = mean(value), SD = sd(value), Median = median(value), Min = min(value), Max = max(value)) %>% 
  ungroup() %>% 
  mutate("CV(%)" = SD/Mean*100) %>% 
  mutate_at(-(1:3), round, 2) %>% 
  mutate_at(-(1:3), ~sprintf("%0.2f", .x)) %>% 
  unite("Mean ± SD", Mean, SD, sep = " ± ") %>% 
  mutate(param = ifelse(Day == "5d.1" & param == "AUCLST", "AUCt", param)) %>% 
  mutate(Day = ifelse(param == "AUCt", "5d", Day)) %>% 
  filter(!Day == "5d.1") %>% 
  filter(!(Day == "1d" & param == "AUCIFO")) %>% 
  mutate(param = factor(param, level = c("AUCt", "AUCLST", "AUCIFO", "CMAX", "CLFO", "TMAX", "LAMZHL", "VZFO"))) %>% 
  select(1:4, 8, everything()) %>% 
  arrange(Day, param) 
  

save(data3, nca3, table3, indiNCA, indiCon, file = "data_at.Rdata")
load("data_at.Rdata")
write_csv(table3, 'table3.csv')

### appendix 

indiNCA <- nca3 %>%
  select(Day, Subject, AUCLST, CMAX, AUCIFO, TMAX, LAMZHL, CLFO, VZFO) %>%
  gather(param, value, AUCLST:VZFO) %>%
  spread(Subject, value) %>% 
  mutate(param = ifelse(Day == "5d.1" & param == "AUCLST", "AUCt", param)) %>% 
  mutate(Day = ifelse(param == "AUCt", "5d", Day)) %>% 
  filter(!Day == "5d.1") %>% 
  filter(!(Day == "1d" & param == "AUCIFO")) %>% 
  mutate(param = factor(param, level = c("AUCt", "AUCLST", "AUCIFO", "CMAX", "CLFO", "TMAX", "LAMZHL", "VZFO"))) %>% 
  arrange(Day, param) %>% 
  mutate_at(-(1:2), round, 2)  

write_csv(indiNCA, 'indiNCA.csv')

indiCon <- data %>% 
  mutate(Time = ifelse(Time == "F_0", -8, Time)) %>% 
  mutate(Time = parse_number(Time)) %>%
  filter(!(Subject %in% dropout)) %>% 
  spread(Subject, Conc) 

write_csv(indiCon, 'indiCon.csv')
