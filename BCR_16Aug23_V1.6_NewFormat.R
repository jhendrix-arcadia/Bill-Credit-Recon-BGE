#Packages to be installed only in the First Run
#(Remove # before the First Run)
#install.packages("data.table")
#install.packages("dplyr")
#install.packages("bizdays")
#install.packages("plyr")
#install.packages("tidyverse")
#install.packages("flextable")
#install.packages('readxl')
#install.packages("xlsx")
#install.packages("writexl")
#install.packages("openxlsx")
#install.packages("berryFunctions")
#install.packages("anytime")
#install.packages("rJava")
#install.packages("formattable")
#install.packages("scales")




#Library Files (Mandatory)
library("data.table")
library("dplyr")
library("bizdays")
library("plyr")
library(tidyverse)
library(officer)
library(flextable)
library(tidyquant)
library(readxl)
library(writexl)
library(openxlsx)
library ("berryFunctions")
library(rJava)
library(xlsx)
library("anytime")  
library(formattable)
library(scales)



#Set Working directory
setwd("D:/Bill Credit Recon/BGE White Subscriber")
list.files(pattern = ".csv")
list.files(pattern = ".xlsx")


#Read input files
Allocation_List<-read_excel("White Subscriber Spreadsheet 20230601.xlsx",sheet="CSEGS Subscriber List")
Host_File<-read_excel("4372372567-48314982023-06-30 - Subcribed detail.xlsx")
Host_File_Final<-read_excel("4372372567-48314982023-06-30 - Subcribed detail.xlsx")
Total_KWh<-read_excel("4372372567-48314982023-06-30 - Subcribed detail.xlsx",range = "H3:I3", col_names = FALSE)
HEX_Report<-read_csv("export_2023-08-18T1307.csv")
Attrition_Report<-read_csv("export_2023-08-18T1308.csv")
Previous_month_report <- read_excel("White Credit Reconciliation Report - April 2023.xlsx",sheet = "April Report")
Accounts_List <- read_csv("VNMSolarProjectExport.csv")
Formula<- read_excel("Foumala.xlsx")


#Renaming Columns in Total_kWh data
colnames(Total_KWh)[1] <- "CSEG Export"
colnames(Total_KWh)[2] <- "kWh"


#Removing the summary part in the previous report and host file
Previous_month_report<- Previous_month_report[-(1:17),] #enter the row no before the header row
Host_File<- Host_File[-(1:5),] #enter the row no before the header row

#Renaming header
names(Previous_month_report) <- Previous_month_report[1,]
names(Host_File) <- Host_File[1,]

#Removing header items from row 1
Previous_month_report <- Previous_month_report[-1,]
Host_File <- Host_File[-1,]
Host_File <- Host_File[-96,]

#-#Changing data type to numeric 
Host_File$`Actual kWh Allocated` <- as.numeric(Host_File$`Actual kWh Allocated`)
Host_File$`Adjustment kWh` <- as.numeric(Host_File$`Adjustment kWh`)
Host_File$`Community Solar Adjustment` <- as.numeric(Host_File$`Community Solar Adjustment`)
Host_File$`Initial Bank Balance kWh` <- as.numeric(Host_File$`Initial Bank Balance kWh`)

#-# Remove hyphen from numbers in the column
Host_File$`Actual kWh Allocated` <- abs(Host_File$`Actual kWh Allocated`)
Host_File$`Adjustment kWh` <- abs(Host_File$`Adjustment kWh`)
Host_File$`Community Solar Adjustment` <- abs(Host_File$`Community Solar Adjustment`)
Host_File$`Initial Bank Balance kWh` <- abs(Host_File$`Initial Bank Balance kWh`)


# ... Part - 1 Columns from Host Bill and Allocation List...

#Renaming Columns in Allocation List
colnames(Allocation_List)[1] <- "choice_ID_AL"
colnames(Allocation_List)[2] <- "Account_No"
colnames(Allocation_List)[3] <- "Percentage_Allocation_List"

#Filtering Required Columns from Allocation List
Allocation_List_Filtered <- Allocation_List[,c('choice_ID_AL','Account_No','Percentage_Allocation_List')]

#Removing Duplicates from the filtered Allocation List w.r.t Choice ID
Allocation_List_Filtered <- Allocation_List_Filtered[!duplicated(Allocation_List_Filtered$`choice_ID_AL`),]

#Filtering Required Columns from Host File
Host_File_Filtered <- Host_File[,c('Subscriber Choice ID','Actual kWh Allocated','Adjustment kWh','Community Solar Adjustment','Allocation Percentage','Initial Bank Balance kWh')]

#Removing Duplicates from the filtered Host File w.r.t Choice ID
Host_File_Filtered <- Host_File_Filtered[!duplicated(Host_File_Filtered$`Subscriber Choice ID`),]

#Merging Host and Allocation List by Choice ID and saving it as Filtered List 1
Filtered_List_1 <- merge(Host_File_Filtered, Allocation_List_Filtered, by.x = c("Subscriber Choice ID"),by.y = c("choice_ID_AL"),all  = TRUE)


#-#Changing data type to numeric 
Filtered_List_1$`Allocation Percentage` <- as.numeric(Filtered_List_1$`Allocation Percentage`)

#Replacing NA with 0 in numeric columns
Filtered_List_1<- Filtered_List_1 %>% mutate_if(is.numeric,~replace_na(.,0))

#Calculating Percentage Difference (Host_File% - Allocation_List%)
Filtered_List_1$PerDiff <- round((Filtered_List_1$`Allocation Percentage`-Filtered_List_1$Percentage_Allocation_List),digits =2)


#Removing Duplicates from the filtered Host File w.r.t Choice ID
Filtered_List_1 <- Filtered_List_1[!duplicated(Filtered_List_1$Account_No),]

#Renaming Column Names in Filtered List 1
colnames(Filtered_List_1)[2] <- "Transferred_kWh"
colnames(Filtered_List_1)[3] <- "Applied_kWh"
colnames(Filtered_List_1)[4] <- "Expected_Credit_from_Customer_Level"
colnames(Filtered_List_1)[5] <- "Percentage_Host_File"
colnames(Filtered_List_1)[7] <- "Account No_Allocation List"

#Calculating Expected KWh
Filtered_List_1$Expected_kWh <- round((Filtered_List_1$`Percentage_Allocation_List` * Total_KWh$kWh))
#-# Remove hyphen from numbers in the column
Filtered_List_1$Expected_kWh <- abs(Filtered_List_1$Expected_kWh)

#Calculating Combo Column in  Filtered_List_1
Filtered_List_1$Combo <-paste0(Filtered_List_1$`Account No_Allocation List`, Filtered_List_1$Expected_Credit_from_Customer_Level)

#... Part - 2 Columns From HEX Report ....

#Calculating Combo in  HEX Report
HEX_Report$Combo <-paste0(HEX_Report$UTILITY_ACCOUNT_NUMBER, HEX_Report$CS_CREDITS_USD)

#Calculating Service end date and payment date in  HEX Report
HEX_Report$SERVICE_END_DATE<- format(as.Date(HEX_Report$SERVICE_END_DATE),'%m/%d/%Y')
HEX_Report$Payment_Date<- format(as.Date(HEX_Report$REMITTED_AT),'%m/%d/%Y')


#Filtering required columns in HEX Report
HEX_Report_Filtered <- HEX_Report[,c('UTILITY_ACCOUNT_NUMBER','SERVICE_END_DATE','CS_CREDITS_USD','SOLAR_DEVELOPER_REMITTANCE_TOTAL_USD','STRIPE_DESTINATION_PAYMENT_ID','Payment_Date','Combo','EP_UTILITY_ACCOUNT_ID')]

#Removing Duplicates w.r.t Combo 
Filtered_List_1 <- Filtered_List_1[!duplicated(Filtered_List_1$Combo),]
HEX_Report_Filtered <- HEX_Report_Filtered[!duplicated(HEX_Report_Filtered$Combo),]

# Remove hyphen from numbers in the column
Filtered_List_1$Combo <- gsub("-", "", Filtered_List_1$Combo)

#Combining Filtered_List_1 and HEX Report and saving it as Filtered List 2
Filtered_List_2 <- merge(Filtered_List_1,HEX_Report_Filtered, by = c("Combo"), all.x = TRUE)

#Changing data type to character
Filtered_List_2$EP_UTILITY_ACCOUNT_ID<- as.character(Filtered_List_2$EP_UTILITY_ACCOUNT_ID)

#Replacing NA with 0 in numeric columns
Filtered_List_2<- Filtered_List_2 %>% mutate_if(is.numeric,~replace_na(.,0))


#Renaming Column Names in Filtered List 2
colnames(Filtered_List_2)[14] <- "Total_Bill_Credits"
colnames(Filtered_List_2)[15] <- "Stripe_Transfer_Amount"
colnames(Filtered_List_2)[16] <- "Payment_ID"
colnames(Filtered_List_2)[18] <- "Utility_Account_Id"

#Calculating Subscription Fee
#-#Changing data type to numeric 
Filtered_List_2$Expected_Credit_from_Customer_Level <- as.numeric(Filtered_List_2$Expected_Credit_from_Customer_Level)

Filtered_List_2$Subscription_Fee_Calc <- round((Filtered_List_2$Expected_Credit_from_Customer_Level * 0.95),digits =2)
Filtered_List_2$Subscription_Fee<- ifelse(Filtered_List_2$Total_Bill_Credits !=0, Filtered_List_2$Subscription_Fee<-Filtered_List_2$Subscription_Fee_Calc,0)

#Rounding off values to 2 decimal places
Filtered_List_2$Stripe_Transfer_Amount<-round(Filtered_List_2$Stripe_Transfer_Amount, digits = 2)
Filtered_List_2$Subscription_Fee<-round(Filtered_List_2$Subscription_Fee, digits = 2)
#-#Remove hyphen from numbers in the column
Filtered_List_2$Subscription_Fee <- abs(Filtered_List_2$Subscription_Fee)
#-#Changing data type to numeric 
#Filtered_List_2$Subscription_Fee <- as.numeric(Filtered_List_2$Subscription_Fee)

#Calculating Stripe Match
Filtered_List_2$Stripe_Match_Check1 <- ifelse(Filtered_List_2$Stripe_Transfer_Amount>0,'T','F')
Filtered_List_2$Stripe_Match_Check2 <- ifelse(Filtered_List_2$Stripe_Transfer_Amount - Filtered_List_2$Subscription_Fee <=0.05 ,'TRUE','FALSE')
Filtered_List_2$Stripe_Match <- ifelse(Filtered_List_2$Stripe_Match_Check1 =='T', Filtered_List_2$Stripe_Match <- Filtered_List_2$Stripe_Match_Check2,'--')

#Calculating Rate
#-#Changing data type to numeric 
Filtered_List_2$Applied_kWh <- as.numeric(Filtered_List_2$Applied_kWh)


Filtered_List_2$Rate <- ifelse(Filtered_List_2$Stripe_Transfer_Amount > 0,Filtered_List_2$Rate<-(Filtered_List_2$Total_Bill_Credits / Filtered_List_2$Applied_kWh),'NA')

#Changing data type to character
Filtered_List_2$Rate<- as.character(Filtered_List_2$Rate)

#Replacing NA with --
Filtered_List_2["Rate"][Filtered_List_2["Rate"] == 'NA'] <- '--'

#...Part - 3 Columns From Attrition Report and Previous Month Report...

#Calculating Utility Rollover Applied
#-#Changing data type to numeric 
Filtered_List_2$Applied_kWh <- as.numeric(Filtered_List_2$Applied_kWh)
Filtered_List_2$Transferred_kWh <- as.numeric(Filtered_List_2$Transferred_kWh)

Filtered_List_2$Utility_Rollover_Applied <- ifelse(Filtered_List_2$Applied_kWh>Filtered_List_2$Transferred_kWh,(Filtered_List_2$Applied_kWh-Filtered_List_2$Transferred_kWh),0)

#Calculating Termination Reason in Attrition Report
Attrition_Report$Termination_Reason <-paste0(Attrition_Report$ATTRITION_DATE, Attrition_Report$ATTRITION_REASON)

#Removing Duplicates w.r.t account no
Attrition_Report <- Attrition_Report[!duplicated(Attrition_Report$UTILITY_ACCOUNT_NUMBER),]
Filtered_List_2 <- Filtered_List_2[!duplicated(Filtered_List_2$`Account No_Allocation List`),]

#Filtering required columns from attrition report
Attrition_Report_Filtered <- Attrition_Report[,c("UTILITY_ACCOUNT_NUMBER","ATTRITION_REASON","ATTRITION_DATE")]
Attrition_Report_Filtered$ATTRITION_DATE<- format(as.Date(Attrition_Report_Filtered$ATTRITION_DATE),'%m/%d/%Y')

#Merging Filtered list 2 and Attrition report filtered w.r.t account no
Filtered_List_2<- merge(Filtered_List_2,Attrition_Report_Filtered, by.x = c("Account No_Allocation List"), by.y = c("UTILITY_ACCOUNT_NUMBER"), all.x= TRUE)

#Calculating Date of Termination
Filtered_List_2$Date_of_Termination <- ifelse(Filtered_List_2$Stripe_Transfer_Amount >0,'--',Filtered_List_2$Date_of_Termination<-Filtered_List_2$ATTRITION_DATE)
Filtered_List_2$Notes<- ifelse(Filtered_List_2$Stripe_Transfer_Amount >0,'--',Filtered_List_2$Notes<-Filtered_List_2$ATTRITION_REASON)

#Adding reasons for Stripe Transfer amount with value 0
Filtered_List_2$Date_of_Termination <- ifelse(Filtered_List_2$Stripe_Transfer_Amount ==0.00,Filtered_List_2$Date_of_Termination<-Filtered_List_2$ATTRITION_DATE,'--')
Filtered_List_2$Notes<- ifelse(Filtered_List_2$Stripe_Transfer_Amount ==0.00,Filtered_List_2$Notes<-Filtered_List_2$ATTRITION_REASON,' ')

#Replacing NA with --
Filtered_List_2 <- Filtered_List_2 %>% replace_na(list(Date_of_Termination = '--',Notes = ' '))

#Filtering Expected to Rollover Next Month from Previous Month Report
df1 <- Previous_month_report[,c("Choice ID","Utility Final Bank")] #Check column name if error occurs

#Changing Column Name 
colnames(df1)[2] <- "Utility Final Bank"

#Removing Duplicates w.r.t choice id
df1 <- df1[!duplicated(df1$`Choice ID`),]
Filtered_List_2 <- Filtered_List_2[!duplicated(Filtered_List_2$`Subscriber Choice ID`),]

#Changing data type to numeric 
df1$`Utility Final Bank` <- as.numeric(df1$`Utility Final Bank`)


#Combining Filtered List 2 and df1 (Filtered list from Previous month report)
Filtered_List_2 <- merge(Filtered_List_2,df1,by.x = c("Subscriber Choice ID"),by.y = c("Choice ID"),all.x = TRUE)

#Replacing NA with 0
Filtered_List_2$`Utility Final Bank` <- Filtered_List_2$`Utility Final Bank` %>% replace_na(0)

#Calculating Utility Initial Bank
#changing data type to numeric
Filtered_List_2$`Initial Bank Balance kWh`<- as.numeric(Filtered_List_2$`Initial Bank Balance kWh`)

Filtered_List_2$`Initial Bank Balance kWh` <- abs(Filtered_List_2$`Initial Bank Balance kWh`)
Filtered_List_2$`Utility Final Bank` <- abs(Filtered_List_2$`Utility Final Bank`)

Filtered_List_2$UIB <- (Filtered_List_2$`Initial Bank Balance kWh`+Filtered_List_2$`Utility Final Bank`)

#Changing Column name
colnames(Filtered_List_2)[31] <- "Utility_Initial_Bank"

#Calculating Utility Banked This Month
Filtered_List_2$UBM_Check <- ifelse(Filtered_List_2$Date_of_Termination == '--',1,0)
Filtered_List_2$Utility_Banked_This_Month <- ifelse(Filtered_List_2$Applied_kWh < Filtered_List_2$Transferred_kWh,(Filtered_List_2$Transferred_kWh - Filtered_List_2$Applied_kWh),0)

#Calculating Utility Final Bank
#changing data type to numeric
Filtered_List_2$Utility_Initial_Bank<- as.numeric(Filtered_List_2$Utility_Initial_Bank)
Filtered_List_2$Utility_Final_Bank_FR <- ((Filtered_List_2$Utility_Initial_Bank - Filtered_List_2$Utility_Rollover_Applied) + Filtered_List_2$Utility_Banked_This_Month)

#Calculating Arcadia Initial Bank
PreviousMonth_List_AIB <- Previous_month_report[,c('Utility Account Number','Arcadia Final Bank')]

#Changing Column Name 
colnames(PreviousMonth_List_AIB)[2] <- "Arcadia Initial Bank"

#Remove duplicates w.r.t Utility Account Number in Previous Month Report
PreviousMonth_List_AIB <- PreviousMonth_List_AIB[!duplicated(PreviousMonth_List_AIB$`Utility Account Number`),]

#Remove duplicates w.r.t Account No_Allocation List in Filtered List 2
Filtered_List_2 <- Filtered_List_2[!duplicated(Filtered_List_2$`Account No_Allocation List`),]

#changing data type to numeric
PreviousMonth_List_AIB$`Arcadia Initial Bank`<- as.numeric(PreviousMonth_List_AIB$`Arcadia Initial Bank`)

#Combining Previous Month Report and Filtered List 2 w.r.t Account No and Saving it as Filtered List 3
Filtered_List_3 <- merge(Filtered_List_2,PreviousMonth_List_AIB, by.x = c("Account No_Allocation List"), by.y = c("Utility Account Number"),all.x= TRUE)

#Replacing NA with 0
Filtered_List_3$`Arcadia Initial Bank` <- Filtered_List_3$`Arcadia Initial Bank` %>% replace_na(0)

#Calculating Arcadia_Rollover_Applied
Filtered_List_3$Arcadia_Rollover_Applied <- round(ifelse((Filtered_List_3$Stripe_Transfer_Amount > Filtered_List_3$Subscription_Fee),(Filtered_List_3$Stripe_Transfer_Amount - Filtered_List_3$Subscription_Fee),0),digits = 2)

#Calculating Arcadia Banked this month
Filtered_List_3$Arcadia_Banked_thismonth <- round(ifelse((Filtered_List_3$Subscription_Fee>Filtered_List_3$Stripe_Transfer_Amount),(Filtered_List_3$Subscription_Fee - Filtered_List_3$Stripe_Transfer_Amount),0),digits = 2)


#Calculating Arcadia Banked Final
Filtered_List_3$Arcadia_Final_Banked <-((Filtered_List_3$`Arcadia Initial Bank` - Filtered_List_3$Arcadia_Rollover_Applied) + Filtered_List_3$Arcadia_Banked_thismonth)


#Calculating Host Bank Credits
Filtered_List_3$Host_Bank_Credits <- ifelse(Filtered_List_3$Transferred_kWh == 0, (Filtered_List_3$Expected_kWh * Total_KWh$kWh * 0.95),0)

#Calculating Missed Revenue
Filtered_List_3$MR_Check <- ifelse((Filtered_List_3$Date_of_Termination == '--' & Filtered_List_3$Stripe_Transfer_Amount == 0),1,0)
Filtered_List_3$Missed_Revenue <- ifelse((Filtered_List_3$MR_Check == 1 & Filtered_List_3$Transferred_kWh > 0), ((Filtered_List_3$Transferred_kWh + Filtered_List_3$Utility_Rollover_Applied) * Total_KWh$kWh * 0.95),0)

#Removing Special Characters
Accounts_List$`utility account number` <- gsub('=','',Accounts_List$`utility account number`)
Accounts_List$`utility account number` <- gsub('[^[:alnum:] ]','',Accounts_List$`utility account number`)

#Filtering required columns from accounts list
Accounts_List_Filtered <- Accounts_List[,c('utility_account_id','utility account number','user_id')]

#Removing Duplicates from the filtered Host File w.r.t utility account number
Accounts_List_Filtered <- Accounts_List_Filtered[!duplicated(Accounts_List_Filtered$`utility account number`),]
Filtered_List_3 <- merge(Filtered_List_3,Accounts_List_Filtered, by.x = c("Account No_Allocation List"), by.y = c("utility account number"),all.x= TRUE)

#Saving Final Report Draft
#write.csv(Filtered_List_3,"FinalReport_Draft_23Aug.csv",row.names = FALSE)

#Creating a copy of final Report Draft
Final_report <- Filtered_List_3

#Removing Columns that are not required in Final Report
Final_report_draft<- Final_report %>% select(-c('UTILITY_ACCOUNT_NUMBER','Combo','Percentage_Host_File','Initial Bank Balance kWh','PerDiff','UBM_Check','MR_Check'))

#Rearranging columns in Final Report
Column_Names_InOrder <- Final_report_draft[,c("utility_account_id","Account No_Allocation List","Subscriber Choice ID","Percentage_Allocation_List","Expected_kWh","Transferred_kWh","Applied_kWh","Expected_Credit_from_Customer_Level","Total_Bill_Credits","Subscription_Fee","Stripe_Transfer_Amount","Stripe_Match","Rate","SERVICE_END_DATE","Payment_ID","Payment_Date","Utility_Initial_Bank","Utility_Rollover_Applied","Utility_Banked_This_Month","Utility_Final_Bank_FR","Arcadia Initial Bank","Arcadia_Rollover_Applied","Arcadia_Banked_thismonth","Arcadia_Final_Banked","Host_Bank_Credits","Missed_Revenue","Date_of_Termination","Notes")]
colnames(Column_Names_InOrder)


#Changing data type to character
Column_Names_InOrder$utility_account_id<- as.character(Column_Names_InOrder$utility_account_id)

#Replacing NA with -- 
Column_Names_InOrder <- Column_Names_InOrder %>% replace_na(list(Stripe_Match = '--',SERVICE_END_DATE  = '--',Payment_ID = '--',Payment_Date = '--',utility_account_id = '--',`Account No_Allocation List` = '--',choice_ID = '--'))




#Dropping the Null row
Column_Names_InOrder <- Column_Names_InOrder %>% filter(row_number() <= n()-1)

#Copying to another data frame
df2<-Column_Names_InOrder


#Converting numeric to currency format
df2$Expected_Credit_from_Customer_Level <- currency(df2$Expected_Credit_from_Customer_Level, digits = 2L)
df2$Total_Bill_Credits <- currency(df2$Total_Bill_Credits, digits = 2L)
df2$Subscription_Fee <- currency(df2$Subscription_Fee, digits = 2L)
df2$Stripe_Transfer_Amount <- currency(df2$Stripe_Transfer_Amount, digits = 2L)
df2$`Arcadia Initial Bank` <- currency(df2$`Arcadia Initial Bank`, digits = 2L)
df2$Arcadia_Rollover_Applied <- currency(df2$Arcadia_Rollover_Applied, digits = 2L)
df2$Arcadia_Banked_thismonth <- currency(df2$Arcadia_Banked_thismonth, digits = 2L)
df2$Arcadia_Final_Banked <- currency(df2$Arcadia_Final_Banked, digits = 2L)
df2$Host_Bank_Credits <- currency(df2$Host_Bank_Credits, digits = 2L)
df2$Missed_Revenue <- currency(df2$Missed_Revenue, digits = 2L)

#Saving Final Report
xlsx::write.xlsx(as.data.frame(df2),"Final_Report_24Aug.xlsx",
                 sheetName="Final_Report",
                 col.names=TRUE,append=TRUE,row.names = FALSE,showNA = FALSE)

#Host_File_Final <- Host_File_Final[!duplicated(Host_File$`Subscriber Choice ID`),]
xlsx::write.xlsx(as.data.frame(Host_File_Final),"Final_Report_24Aug.xlsx",
                 sheetName="Host_File",
                 col.names=TRUE,append=TRUE,row.names = FALSE,showNA = FALSE)

#xlsx::write.xlsx(as.data.frame(Total_KWh),"Final_Report_18Aug.xlsx",
                 #sheetName="Total kWh",
                 #col.names=TRUE,append=TRUE,row.names = FALSE,showNA = FALSE)

xlsx::write.xlsx(as.data.frame(Allocation_List_Filtered),"Final_Report_24Aug.xlsx",
                 sheetName="Allocation_List",
                 col.names=TRUE,append=TRUE,row.names = FALSE,showNA = FALSE)

xlsx::write.xlsx(as.data.frame(Formula),"Final_Report_24Aug.xlsx",
                 sheetName="Formula",
                 col.names=TRUE,append=TRUE,row.names = FALSE,showNA = FALSE)

