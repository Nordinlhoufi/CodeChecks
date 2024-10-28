### Retrieving used packages ###  Packages can only be retrieved when once installed already. Install by typing " install.packages("xlsx") " in Console
library("xts")
library("sandwich")
library("lmtest")
library("tseries")

### Redirects to the correct working directory ### 
### setwd("P:/Valuations/Market Adjustment Factor/R-Script/2024 Q1")
setwd("C:/Users/jzoghlami/Documents/MAF")

### Reads in current data ###
UPCMVexclFinancials<-read.csv("Input Files/MV.csv",check.names=FALSE,stringsAsFactors = FALSE,fileEncoding="UTF-8-BOM")
UPCInvested<-read.csv("Input Files/Invested.csv",check.names=FALSE,stringsAsFactors = FALSE,fileEncoding="UTF-8-BOM")
UPCProceeds<-read.csv("Input Files/Proceeds.csv",check.names=FALSE,stringsAsFactors = FALSE,fileEncoding="UTF-8-BOM")
UPCMapping<-read.csv("Input Files/Holding GUIDs.csv",check.names=FALSE,stringsAsFactors = FALSE,fileEncoding="UTF-8-BOM")

FIGI<-read.csv("Input Files/Bloomberg FIGI Files.csv",check.names=FALSE,stringsAsFactors = FALSE,fileEncoding="UTF-8-BOM")
Indices<-read.csv("Input Files/Final Index File Bloomberg (hardcopy).csv",check.names=FALSE,stringsAsFactors = FALSE,fileEncoding="UTF-8-BOM")
### Important to keep Indices input files in same format. Identifiers (example "US BO 10") plus dates in uneven columns and Index name plus values in even columns. ###

UPCMVexclFinancials$`Closing Date`<-as.Date(UPCMVexclFinancials$`Closing Date`,format = "%m/%d/%Y")
UPCMVexclFinancials$`Exit Date`<-as.Date(UPCMVexclFinancials$`Exit Date`,format = "%m/%d/%Y")
colnames(UPCMVexclFinancials)[14:ncol(UPCMVexclFinancials)]<-paste(substr(colnames(UPCMVexclFinancials)[14:ncol(UPCMVexclFinancials)],1,4),substr(colnames(UPCMVexclFinancials)[14:ncol(UPCMVexclFinancials)],5,6))         #First 13 columns are static data and set column header in correct format
UPCMV<-UPCMVexclFinancials 
UPCInfo<-UPCMV[,1:13]      ### Create data frame containing static data points per UPC

### Defines the superset of Quarters and Market Indices ###
QuartersVector<-as.data.frame(paste((1:100-1)%/%4+2000,c("Q4","Q1","Q2","Q3")[1:100%%4+1]))
names(QuartersVector) <- "Qtr"
MarketIndices<-colnames(Indices)[seq(2,ncol(Indices),2)]

### Function to retrieve position of last day of the Quarter, needed as trading days differ per index ###
GetLastQuarter <- function(x){
  myXts <-as.xts(Indices[!(is.na(Indices[,x*2-1]) | Indices[,x*2-1] == "") ,(x*2-1):(x*2)],order.by = as.Date(Indices[!(is.na(Indices[,x*2-1]) | Indices[,x*2-1] == ""),x*2-1],format = "%m/%d/%Y")) 
  Y <- myXts[xts:::endof(myXts,"quarters")]
  y <- index(Y)
  
  ix <- which(as.Date(Indices[,x*2-1],format = "%m/%d/%Y") %in% y )
  
  QQ <- cbind.data.frame(as.character(as.yearqtr(y,format = "%m/%d/%Y")),ix)
  names(QQ)[1] <- "Qtr"
  YY <- merge(QuartersVector,QQ, by = "Qtr",all.x = TRUE)
  
  zz <- YY[,2]
  return(zz)
}
LastDayofQuarter <- sapply(1:length(MarketIndices),GetLastQuarter)
rownames(LastDayofQuarter) <- QuartersVector$Qtr
colnames(LastDayofQuarter)<-MarketIndices

### Calculate the % change in the relevant market index mean price over the relevant DaysBefore duration, between current quarter and prior quarter ###
MarketChangeFormula<-function(LastDayofQuarter,SelectedMarketIndices,Quarters,DaysBeforeCQ,DaysBeforePQ,DaysAfterPQ,LastDay=1)
  # LastDayOfQuarter - matrix of integer index values for positions of last tradings of each quarter, per market index (column headings)
  # SelectedMarketIndices - vector of string values containing selected market index identifiers
  # Quarters - vector of string values containing selected quarters over which the market change needs to be determined
  # DaysBeforeCQ - integer variable containing number of days before current quarter end over which to calculate mean
  # DaysBeforePQ - integer variable containing number of days before prior quarter end over which to calculate mean
  # DaysAfterPQ - days with which to extend period over which prior quarter mean is calculated
{
  MarketIndicesSubset<-match(SelectedMarketIndices, colnames(LastDayofQuarter))
  AverageIndexValueCQ<-sapply(MarketIndicesSubset,function(j) sapply(Quarters, function(i) if(!is.na(LastDayofQuarter[i,j])){if(LastDayofQuarter[i,j]>(DaysBeforeCQ-1)){mean(as.numeric(Indices[c((LastDayofQuarter[i,j]-DaysBeforeCQ+1):LastDayofQuarter[i,j],rep(LastDayofQuarter[i,j],(LastDay-1))),j*2]))}else{NA}}else{NA}))
  AverageIndexValuePQ<-sapply(MarketIndicesSubset,function(j) sapply(Quarters, function(i) if(!is.na(LastDayofQuarter[i,j])){if(LastDayofQuarter[i,j]>(DaysBeforePQ-1)){mean(as.numeric(Indices[(LastDayofQuarter[i,j]-DaysBeforePQ+1):(LastDayofQuarter[i,j]+DaysAfterPQ),j*2]))}else{NA}}else{NA}))
  MarketChange<-AverageIndexValueCQ[-1,]/AverageIndexValuePQ[-nrow(AverageIndexValuePQ),]-1
  rownames(MarketChange)<-Quarters[-1]
  colnames(MarketChange)<-SelectedMarketIndices
  return(MarketChange)
}

### Set 0 (or close to zero) invested to NA as invested cannot be 0 ###
UPCInvested[UPCInvested<=100]<-NA

### Fill in the gaps in Invested values ###
UPCInvested[,15:(ncol(UPCInvested)-4)][is.na(UPCInvested[,15:(ncol(UPCInvested)-4)]) & !is.na(UPCInvested[,14:(ncol(UPCInvested)-5)]) & !is.na(UPCInvested[,19:ncol(UPCInvested)]) & UPCInvested[,14:(ncol(UPCInvested)-5)]==UPCInvested[,19:ncol(UPCInvested)]]<-UPCInvested[,14:(ncol(UPCInvested)-5)][is.na(UPCInvested[,15:(ncol(UPCInvested)-4)]) & !is.na(UPCInvested[,14:(ncol(UPCInvested)-5)]) & !is.na(UPCInvested[,19:ncol(UPCInvested)]) & UPCInvested[,14:(ncol(UPCInvested)-5)]==UPCInvested[,19:ncol(UPCInvested)]]
UPCInvested[,15:(ncol(UPCInvested)-3)][is.na(UPCInvested[,15:(ncol(UPCInvested)-3)]) & !is.na(UPCInvested[,14:(ncol(UPCInvested)-4)]) & !is.na(UPCInvested[,18:ncol(UPCInvested)]) & UPCInvested[,14:(ncol(UPCInvested)-4)]==UPCInvested[,18:ncol(UPCInvested)]]<-UPCInvested[,14:(ncol(UPCInvested)-4)][is.na(UPCInvested[,15:(ncol(UPCInvested)-3)]) & !is.na(UPCInvested[,14:(ncol(UPCInvested)-4)]) & !is.na(UPCInvested[,18:ncol(UPCInvested)]) & UPCInvested[,14:(ncol(UPCInvested)-4)]==UPCInvested[,18:ncol(UPCInvested)]]
UPCInvested[,15:(ncol(UPCInvested)-2)][is.na(UPCInvested[,15:(ncol(UPCInvested)-2)]) & !is.na(UPCInvested[,14:(ncol(UPCInvested)-3)]) & !is.na(UPCInvested[,17:ncol(UPCInvested)]) & UPCInvested[,14:(ncol(UPCInvested)-3)]==UPCInvested[,17:ncol(UPCInvested)]]<-UPCInvested[,14:(ncol(UPCInvested)-3)][is.na(UPCInvested[,15:(ncol(UPCInvested)-2)]) & !is.na(UPCInvested[,14:(ncol(UPCInvested)-3)]) & !is.na(UPCInvested[,17:ncol(UPCInvested)]) & UPCInvested[,14:(ncol(UPCInvested)-3)]==UPCInvested[,17:ncol(UPCInvested)]]
UPCInvested[,15:(ncol(UPCInvested)-1)][is.na(UPCInvested[,15:(ncol(UPCInvested)-1)]) & !is.na(UPCInvested[,14:(ncol(UPCInvested)-2)]) & !is.na(UPCInvested[,16:ncol(UPCInvested)]) & UPCInvested[,14:(ncol(UPCInvested)-2)]==UPCInvested[,16:ncol(UPCInvested)]]<-UPCInvested[,14:(ncol(UPCInvested)-2)][is.na(UPCInvested[,15:(ncol(UPCInvested)-1)]) & !is.na(UPCInvested[,14:(ncol(UPCInvested)-2)]) & !is.na(UPCInvested[,16:ncol(UPCInvested)]) & UPCInvested[,14:(ncol(UPCInvested)-2)]==UPCInvested[,16:ncol(UPCInvested)]]

### Set 0 LTD proceeds to NA when non zero in prior quarters as not possible ###
UPCProceedsTemp<-UPCProceeds[,-c(1:13)]
UPCProceedsTemp[is.na(UPCProceedsTemp)]<-0
UPCProceedsCumulative<-cbind(UPCInfo,t(apply(UPCProceedsTemp,1,cumsum)))
UPCProceeds[UPCProceeds==0 & UPCProceedsCumulative>0]<-NA

### Fill in the gaps in Proceeds values ###
UPCProceeds[,15:(ncol(UPCProceeds)-4)][is.na(UPCProceeds[,15:(ncol(UPCProceeds)-4)]) & !is.na(UPCProceeds[,14:(ncol(UPCProceeds)-5)]) & !is.na(UPCProceeds[,19:ncol(UPCProceeds)]) & UPCProceeds[,14:(ncol(UPCProceeds)-5)]==UPCProceeds[,19:ncol(UPCProceeds)]]<-UPCProceeds[,14:(ncol(UPCProceeds)-5)][is.na(UPCProceeds[,15:(ncol(UPCProceeds)-4)]) & !is.na(UPCProceeds[,14:(ncol(UPCProceeds)-5)]) & !is.na(UPCProceeds[,19:ncol(UPCProceeds)]) & UPCProceeds[,14:(ncol(UPCProceeds)-5)]==UPCProceeds[,19:ncol(UPCProceeds)]]
UPCProceeds[,15:(ncol(UPCProceeds)-3)][is.na(UPCProceeds[,15:(ncol(UPCProceeds)-3)]) & !is.na(UPCProceeds[,14:(ncol(UPCProceeds)-4)]) & !is.na(UPCProceeds[,18:ncol(UPCProceeds)]) & UPCProceeds[,14:(ncol(UPCProceeds)-4)]==UPCProceeds[,18:ncol(UPCProceeds)]]<-UPCProceeds[,14:(ncol(UPCProceeds)-4)][is.na(UPCProceeds[,15:(ncol(UPCProceeds)-3)]) & !is.na(UPCProceeds[,14:(ncol(UPCProceeds)-4)]) & !is.na(UPCProceeds[,18:ncol(UPCProceeds)]) & UPCProceeds[,14:(ncol(UPCProceeds)-4)]==UPCProceeds[,18:ncol(UPCProceeds)]]
UPCProceeds[,15:(ncol(UPCProceeds)-2)][is.na(UPCProceeds[,15:(ncol(UPCProceeds)-2)]) & !is.na(UPCProceeds[,14:(ncol(UPCProceeds)-3)]) & !is.na(UPCProceeds[,17:ncol(UPCProceeds)]) & UPCProceeds[,14:(ncol(UPCProceeds)-3)]==UPCProceeds[,17:ncol(UPCProceeds)]]<-UPCProceeds[,14:(ncol(UPCProceeds)-3)][is.na(UPCProceeds[,15:(ncol(UPCProceeds)-2)]) & !is.na(UPCProceeds[,14:(ncol(UPCProceeds)-3)]) & !is.na(UPCProceeds[,17:ncol(UPCProceeds)]) & UPCProceeds[,14:(ncol(UPCProceeds)-3)]==UPCProceeds[,17:ncol(UPCProceeds)]]
UPCProceeds[,15:(ncol(UPCProceeds)-1)][is.na(UPCProceeds[,15:(ncol(UPCProceeds)-1)]) & !is.na(UPCProceeds[,14:(ncol(UPCProceeds)-2)]) & !is.na(UPCProceeds[,16:ncol(UPCProceeds)]) & UPCProceeds[,14:(ncol(UPCProceeds)-2)]==UPCProceeds[,16:ncol(UPCProceeds)]]<-UPCProceeds[,14:(ncol(UPCProceeds)-2)][is.na(UPCProceeds[,15:(ncol(UPCProceeds)-1)]) & !is.na(UPCProceeds[,14:(ncol(UPCProceeds)-2)]) & !is.na(UPCProceeds[,16:ncol(UPCProceeds)]) & UPCProceeds[,14:(ncol(UPCProceeds)-2)]==UPCProceeds[,16:ncol(UPCProceeds)]]

### Exclude the first X days after entry because UPC is often valued at cost. Cut-off chosen at 365 days, see file "First Valuation support for 365 days cut off" ###
DaysAfterClosing<-365
for(i in 1:nrow(UPCMV))
{
  if(!is.na(UPCMV$`Closing Date`[i]))
  {
    if((ncol(UPCMV)-max(0,which(colnames(UPCMV)==as.character(as.yearqtr(UPCMV$`Closing Date`[i]))))+1)<=4)
    {
      UPCMV[i,which(colnames(UPCMV)==as.character(as.yearqtr(UPCMV$`Closing Date`[i]))):ncol(UPCMV)]<-NA
    }
    UPCMV[i,c(max(14,which(colnames(UPCMV)==as.character(as.yearqtr(UPCMV$`Closing Date`[i]+DaysAfterClosing)))-4):max(14,which(colnames(UPCMV)==as.character(as.yearqtr(UPCMV$`Closing Date`[i]+DaysAfterClosing)))-1))]<-NA
  }
}

### Exclude the tail end valuations after minimal X years ###
YearsAfterInception<-10
for(i in 1:nrow(UPCMV))
{
  if(UPCMV$VY[i]!=1111){UPCMV[i,which(UPCMV$VY[i]+YearsAfterInception<as.numeric(substr(colnames(UPCMV),1,4)))]<-NA}
}

### Calculate the Value Creation ###
UPCFMV<-UPCMV[,-c(1:13,ncol(UPCMV))]+(UPCInvested[,-c(1:14)]-UPCInvested[,-c(1:13,ncol(UPCInvested))])-(UPCProceeds[,-c(1:14)]-UPCProceeds[,-c(1:13,ncol(UPCProceeds))])
colnames(UPCFMV)<-colnames(UPCMV)[-c(1:14)]
UPCFMV[UPCFMV<0]<-NA

### Remove exits. Defined as quarters with proceeds and MV remains 0 onwards ###
### EDITS 1 : suppress all the below from original MAF script code (below commented)
#UPCMVTemp<-UPCMV[,-c(1:13)]
#UPCMVTemp[is.na(UPCMVTemp)]<-0
#UPCMVTemp<-t(apply(UPCMVTemp[,ncol(UPCMVTemp):1],1,cumsum))
#UPCMVTemp<-UPCMVTemp[,ncol(UPCMVTemp):1]
#UPCFMV[(UPCProceeds[,-c(1:14)]-UPCProceeds[,-c(1:13,ncol(UPCProceeds))])>0 & UPCMVTemp[,-1]==0]<-NA
#UPCValueCreation<-cbind(UPCInfo,UPCMV[,-c(1:14)]-UPCFMV)

###EDIT 2 : add the following lines of code for the treatment of post-exit quarter removal that has been approved by Kroll
UPCMVTemp <-UPCMV[,-c(1:13)] %>% replace(is.na(.), 0) #all value columns 
UPCMVTemp<-t(apply(UPCMVTemp[,ncol(UPCMVTemp):1],1,cumsum)) #transpose (no cumulative sum is calculated)
UPCMVTemp<- data.frame(UPCMVTemp[,ncol(UPCMVTemp):1])
names(UPCMVTemp) <- names(UPCMV[,-c(1:13)])
  
#old method corrected 
#Higher proceeds than previous quarter and zero MV in next quarter with new additional action that all subsequent quarters are set to zero
UPCFMV[(UPCProceeds[,-c(1:14)]-UPCProceeds[,-c(1:13,ncol(UPCProceeds))])>0 & UPCMVTemp[,-1]==0] <- "Exit quarter inferred old method"
UPCFMV = data.frame(t(data.frame(t(UPCFMV)) %>% mutate(across(starts_with("X"), ~ {position <- which(. == "Exit quarter inferred old method")[1]  #Find first identified exit quarter per UPC
                            if (!is.na(position)) {
                              replace(., seq_along(.) > position, "Post-Exit quarter inferred") #replace all subsequent occurrences
                            } else {.}})))) %>% `rownames<-`(seq_len(nrow(UPCFMV)))  %>% #output here if you want to create a separate df for further inspection in Excel etc.
    mutate(across(where(is.character), ~ replace(., . %in% c("Exit quarter inferred old method", "Post-Exit quarter inferred"), NA))) %>% mutate_if(is.character,as.numeric)


UPCValueCreation<-cbind(UPCInfo,UPCMV[,-c(1:14)]-UPCFMV)

###END EDITS


### Remove Market Value when there is no Value Creation in the subsequent Period to ensure the value change is calculated on an accurate basis ###
UPCMV[,-c(1:13,ncol(UPCMV))][is.na(UPCValueCreation[,-c(1:13)]>UPCMV[,-c(1:13,ncol(UPCMV))])]<-NA

### Remove the duplicated Partnership Group Line items due to different Secondary deals ####
UPCMV<-aggregate(UPCMV[,14:ncol(UPCMV)], by=list(UPCMV$`Deal Name - Partnership Group`,UPCMV$`UPC Holding GUID`),mean,na.rm = TRUE)
UPCValueCreation<-aggregate(UPCValueCreation[,14:ncol(UPCValueCreation)], by=list(UPCValueCreation$`Deal Name - Partnership Group`,UPCValueCreation$`UPC Holding GUID`),mean,na.rm = TRUE)

### Remove Partnership Group so that each UPC is included once. To distinguish between NA values and actual 0 values the mean is first calculated which evaluates to NaN in case of NA values ####
UPCMVNaN<-aggregate(UPCMV[,3:ncol(UPCMV)], by=list(UPCMV$Group.2),mean,na.rm = TRUE)
UPCMV<-aggregate(UPCMV[,3:ncol(UPCMV)], by=list(UPCMV$Group.2),sum,na.rm = TRUE)
UPCMV[is.na(UPCMVNaN)]<-NA
UPCValueCreationNaN<-aggregate(UPCValueCreation[,3:ncol(UPCValueCreation)], by=list(UPCValueCreation$Group.2),mean,na.rm = TRUE)
UPCValueCreation<-aggregate(UPCValueCreation[,3:ncol(UPCValueCreation)], by=list(UPCValueCreation$Group.2),sum,na.rm = TRUE)
UPCValueCreation[is.na(UPCValueCreationNaN)]<-NA
rownames(UPCMV)<-UPCMV[,1]
rownames(UPCValueCreation)<-UPCValueCreation[,1]
UPCValueCreation<-UPCValueCreation[,-1]

### Calculate UPC Value Change. We estimated optimal cut-off by looking at minimal RMSE for various cut-off percentages ###
MaxValueChange<-0.24
UPCValueChange<-UPCValueCreation/UPCMV[,-c(1,ncol(UPCMV))]
UPCValueChange[abs(UPCValueChange)>MaxValueChange]<-NA
UPCMapping<-UPCMapping[match(rownames(UPCValueChange),UPCMapping[,1]),]
UPCValueChange<-cbind(UPCMapping[,-1],UPCValueChange)
rownames(UPCValueChange)<-UPCMapping[,1]

### Value Change Overall ###
ValueChangeOverallUnweighted<-t(t(colMeans(UPCValueChange[,-c(1:3)],na.rm = TRUE)))
colnames(ValueChangeOverallUnweighted)<-"Overall"
ValueChangeOverallCapitalWeighted<-t(t(colSums(UPCValueCreation,na.rm=TRUE)/colSums(UPCMV[,-c(1,ncol(UPCMV))],na.rm=TRUE)))
colnames(ValueChangeOverallCapitalWeighted)<-colnames(ValueChangeOverallUnweighted)
ValueChangeOverallCapitalWeighted[abs(ValueChangeOverallCapitalWeighted)>MaxValueChange]<-NA

### Value Change per Region ###
ValueChangeRegionUnweighted<-aggregate(UPCValueChange[,-c(1:3)], by=list(UPCValueChange$Region),mean,na.rm = TRUE)
rownames(ValueChangeRegionUnweighted)<-ValueChangeRegionUnweighted[,1]
ValueChangeRegionUnweighted<-t(ValueChangeRegionUnweighted[,-1])
ValueChangeRegionCapitalWeighted<-t(aggregate(UPCValueCreation, by=list(UPCMapping$Region),sum,na.rm = TRUE)[,-1]/aggregate(UPCMV[,-c(1,ncol(UPCMV))], by=list(UPCMapping$Region),sum,na.rm = TRUE)[,-1])
colnames(ValueChangeRegionCapitalWeighted)<-colnames(ValueChangeRegionUnweighted)
ValueChangeRegionCapitalWeighted[abs(ValueChangeRegionCapitalWeighted)>MaxValueChange]<-NA

### Value Change per Stage ###
ValueChangeStageUnweighted<-aggregate(UPCValueChange[,-c(1:3)], by=list(UPCValueChange$Stage),mean,na.rm = TRUE)
rownames(ValueChangeStageUnweighted)<-ValueChangeStageUnweighted[,1]
ValueChangeStageUnweighted<-t(ValueChangeStageUnweighted[,-1])
ValueChangeStageCapitalWeighted<-t(aggregate(UPCValueCreation, by=list(UPCMapping$Stage),sum,na.rm = TRUE)[,-1]/aggregate(UPCMV[,-c(1,ncol(UPCMV))], by=list(UPCMapping$Stage),sum,na.rm = TRUE)[,-1])
colnames(ValueChangeStageCapitalWeighted)<-colnames(ValueChangeStageUnweighted)
ValueChangeStageCapitalWeighted[abs(ValueChangeStageCapitalWeighted)>MaxValueChange]<-NA

### Value Change per Sector ###
ValueChangeSectorUnweighted<-aggregate(UPCValueChange[,-c(1:3)], by=list(UPCValueChange$Sector),mean,na.rm = TRUE)
rownames(ValueChangeSectorUnweighted)<-ValueChangeSectorUnweighted[,1]
ValueChangeSectorUnweighted<-t(ValueChangeSectorUnweighted[,-1])
ValueChangeSectorCapitalWeighted<-t(aggregate(UPCValueCreation, by=list(UPCMapping$Sector),sum,na.rm = TRUE)[,-1]/aggregate(UPCMV[,-c(1,ncol(UPCMV))], by=list(UPCMapping$Sector),sum,na.rm = TRUE)[,-1])
colnames(ValueChangeSectorCapitalWeighted)<-colnames(ValueChangeSectorUnweighted)
ValueChangeSectorCapitalWeighted[abs(ValueChangeSectorCapitalWeighted)>MaxValueChange]<-NA

### Value Change per Region & Stage ###
ValueChangeRegionStageUnweighted<-aggregate(UPCValueChange[,-c(1:3)], by=list(UPCValueChange$Region,UPCValueChange$Stage),mean,na.rm = TRUE)
rownames(ValueChangeRegionStageUnweighted)<-paste(ValueChangeRegionStageUnweighted[,1],ValueChangeRegionStageUnweighted[,2])
ValueChangeRegionStageUnweighted<-t(ValueChangeRegionStageUnweighted[,-c(1:2)])
ValueChangeRegionStageCapitalWeighted<-t(aggregate(UPCValueCreation, by=list(UPCMapping$Region,UPCMapping$Stage),sum,na.rm = TRUE)[,-c(1,2)]/aggregate(UPCMV[,-c(1,ncol(UPCMV))], by=list(UPCMapping$Region,UPCMapping$Stage),sum,na.rm = TRUE)[,-c(1,2)])
colnames(ValueChangeRegionStageCapitalWeighted)<-colnames(ValueChangeRegionStageUnweighted)
ValueChangeRegionStageCapitalWeighted[abs(ValueChangeRegionStageCapitalWeighted)>MaxValueChange]<-NA

### Value Change per Region & Sector ###
ValueChangeRegionSectorUnweighted<-aggregate(UPCValueChange[,-c(1:3)], by=list(UPCValueChange$Region,UPCValueChange$Sector),mean,na.rm = TRUE)
rownames(ValueChangeRegionSectorUnweighted)<-paste(ValueChangeRegionSectorUnweighted[,1],ValueChangeRegionSectorUnweighted[,2])
ValueChangeRegionSectorUnweighted<-t(ValueChangeRegionSectorUnweighted[,-c(1:2)])
ValueChangeRegionSectorCapitalWeighted<-t(aggregate(UPCValueCreation, by=list(UPCMapping$Region,UPCMapping$Sector),sum,na.rm = TRUE)[,-c(1,2)]/aggregate(UPCMV[,-c(1,ncol(UPCMV))], by=list(UPCMapping$Region,UPCMapping$Sector),sum,na.rm = TRUE)[,-c(1,2)])
colnames(ValueChangeRegionSectorCapitalWeighted)<-colnames(ValueChangeRegionSectorUnweighted)
ValueChangeRegionSectorCapitalWeighted[abs(ValueChangeRegionSectorCapitalWeighted)>MaxValueChange]<-NA

### Value Change per Stage & Sector ###
ValueChangeStageSectorUnweighted<-aggregate(UPCValueChange[,-c(1:3)], by=list(UPCValueChange$Stage,UPCValueChange$Sector),mean,na.rm = TRUE)
rownames(ValueChangeStageSectorUnweighted)<-paste(ValueChangeStageSectorUnweighted[,1],ValueChangeStageSectorUnweighted[,2])
ValueChangeStageSectorUnweighted<-t(ValueChangeStageSectorUnweighted[,-c(1:2)])
ValueChangeStageSectorCapitalWeighted<-t(aggregate(UPCValueCreation, by=list(UPCMapping$Stage,UPCMapping$Sector),sum,na.rm = TRUE)[,-c(1,2)]/aggregate(UPCMV[,-c(1,ncol(UPCMV))], by=list(UPCMapping$Stage,UPCMapping$Sector),sum,na.rm = TRUE)[,-c(1,2)])
colnames(ValueChangeStageSectorCapitalWeighted)<-colnames(ValueChangeStageSectorUnweighted)
ValueChangeStageSectorCapitalWeighted[abs(ValueChangeStageSectorCapitalWeighted)>MaxValueChange]<-NA

### Value Change per Bucket ###
ValueChangeBucketUnweighted<-aggregate(UPCValueChange[,-c(1:3)], by=list(UPCValueChange$Region,UPCValueChange$Stage,UPCValueChange$Sector),mean,na.rm = TRUE)
rownames(ValueChangeBucketUnweighted)<-paste(ValueChangeBucketUnweighted[,1],ValueChangeBucketUnweighted[,2],ValueChangeBucketUnweighted[,3])
ValueChangeBucketUnweighted<-t(ValueChangeBucketUnweighted[,-c(1:3)])
ValueChangeBucketCapitalWeighted<-t(aggregate(UPCValueCreation, by=list(UPCMapping$Region,UPCMapping$Stage,UPCMapping$Sector),sum,na.rm = TRUE)[,-c(1,2,3)]/aggregate(UPCMV[,-c(1,ncol(UPCMV))], by=list(UPCMapping$Region,UPCMapping$Stage,UPCMapping$Sector),sum,na.rm = TRUE)[,-c(1,2,3)])
colnames(ValueChangeBucketCapitalWeighted)<-colnames(ValueChangeBucketUnweighted)
ValueChangeBucketCapitalWeighted[abs(ValueChangeBucketCapitalWeighted)>MaxValueChange]<-NA


#### Select Weighting Method ###
CapitalWeighted<-TRUE
ValueChangeOverall<-if(CapitalWeighted){ValueChangeOverallCapitalWeighted}else{ValueChangeOverallUnweighted}
ValueChangeRegion<-if(CapitalWeighted){ValueChangeRegionCapitalWeighted}else{ValueChangeRegionUnweighted}
ValueChangeStage<-if(CapitalWeighted){ValueChangeStageCapitalWeighted}else{ValueChangeStageUnweighted}
ValueChangeSector<-if(CapitalWeighted){ValueChangeSectorCapitalWeighted}else{ValueChangeSectorUnweighted}
ValueChangeRegionStage<-if(CapitalWeighted){ValueChangeRegionStageCapitalWeighted}else{ValueChangeRegionStageUnweighted}
ValueChangeRegionSector<-if(CapitalWeighted){ValueChangeRegionSectorCapitalWeighted}else{ValueChangeRegionSectorUnweighted}
ValueChangeStageSector<-if(CapitalWeighted){ValueChangeStageSectorCapitalWeighted}else{ValueChangeStageSectorUnweighted}
ValueChangeBucket<-if(CapitalWeighted){ValueChangeBucketCapitalWeighted}else{ValueChangeBucketUnweighted}

### Calculate Predicted Values ####
DaysBeforeCQ<-25
DaysBeforePQ<-25
DaysAfterPQ<-0
FirstQuarter<-"2013 Q1"
LastQuarter<-colnames(UPCMV)[ncol(UPCMV)]     #Can be manually set to any quarter
EstimationWindow<-0                           #In Quarters.Set equal to 0 if expanding window is preferred.
ExpandingWindowFirstQuarter<-"2008 Q1"

QuartersNeeded <- as.vector(QuartersVector[(min(which(QuartersVector==FirstQuarter)-EstimationWindow-1,which(QuartersVector==ExpandingWindowFirstQuarter)-1)):which(QuartersVector==LastQuarter),])
TestPeriod <- as.vector(QuartersVector[which(QuartersVector==FirstQuarter):which(QuartersVector==LastQuarter),])
MarketChange<-MarketChangeFormula(LastDayofQuarter,MarketIndices,QuartersNeeded,DaysBeforeCQ,DaysBeforePQ,DaysAfterPQ)
SelectedIndicesRegion<-SelectedIndicesStage<-SelectedIndicesSector<-SelectedIndicesRegionStage<-SelectedIndicesRegionSector<-SelectedIndicesStageSector<-SelectedIndicesBucket<-c()
PredictedOverall<-array(NA,c(length(TestPeriod),ncol(ValueChangeOverall),2),list(TestPeriod,colnames(ValueChangeOverall),c("Forecasted","Fitted")))
PredictedRegion<-array(NA,c(length(TestPeriod),ncol(ValueChangeRegion),2),list(TestPeriod,colnames(ValueChangeRegion),c("Forecasted","Fitted")))
PredictedStage<-array(NA,c(length(TestPeriod),ncol(ValueChangeStage),2),list(TestPeriod,colnames(ValueChangeStage),c("Forecasted","Fitted")))
PredictedSector<-array(NA,c(length(TestPeriod),ncol(ValueChangeSector),2),list(TestPeriod,colnames(ValueChangeSector),c("Forecasted","Fitted")))
PredictedRegionStage<-array(NA,c(length(TestPeriod),ncol(ValueChangeRegionStage),2),list(TestPeriod,colnames(ValueChangeRegionStage),c("Forecasted","Fitted")))
PredictedRegionSector<-array(NA,c(length(TestPeriod),ncol(ValueChangeRegionSector),2),list(TestPeriod,colnames(ValueChangeRegionSector),c("Forecasted","Fitted")))
PredictedStageSector<-array(NA,c(length(TestPeriod),ncol(ValueChangeStageSector),2),list(TestPeriod,colnames(ValueChangeStageSector),c("Forecasted","Fitted")))
PredictedBucket<-array(NA,c(length(TestPeriod),ncol(ValueChangeBucket),2),list(TestPeriod,colnames(ValueChangeBucket),c("Forecasted","Fitted")))

### Calculate Forecasted Values ###
for(Q in TestPeriod)
{
  if(EstimationWindow!=0){Quarters <- as.vector(QuartersVector[(which(QuartersVector==Q)-EstimationWindow):which(QuartersVector==Q),])}
  if(EstimationWindow==0){Quarters <- as.vector(QuartersVector[(which(QuartersVector==ExpandingWindowFirstQuarter)):which(QuartersVector==Q),])} 
  if(Q==FirstQuarter){SelectedIndexOverall<-names(which.max(as.data.frame(cor(ValueChangeOverall[QuartersNeeded[-1],],MarketChange,use="pairwise.complete.obs"))))}
  RegressionData<-as.data.frame(cbind(ValueChangeOverall[Quarters,,drop=FALSE],MarketChange[Quarters,SelectedIndexOverall,drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  Regression<-lm(ValueChange ~ Index, RegressionData[-nrow(RegressionData),])
  PredictedOverall[Q,,"Forecasted"]<-Regression$coefficients[1]+Regression$coefficients[2]*RegressionData[nrow(RegressionData),2]
  for(i in 1:ncol(ValueChangeRegion))
  {
    if(Q==FirstQuarter)
    {
      PotentialIndices<-names(Indices)[which(substr(names(Indices),1,2) == colnames(ValueChangeRegion)[i] & rep(c(TRUE,FALSE),length(MarketIndices)))+1]
      SelectedIndicesRegion<-c(SelectedIndicesRegion,names(which.max(as.data.frame(cor(ValueChangeRegion[QuartersNeeded[-1],colnames(ValueChangeRegion)[i]],MarketChange[,PotentialIndices],use="pairwise.complete.obs")))))
    }
    if(!is.na(SelectedIndicesRegion[i]))
    {
      RegressionData<-as.data.frame(cbind(ValueChangeRegion[Quarters,colnames(ValueChangeRegion)[i],drop=FALSE],MarketChange[Quarters,SelectedIndicesRegion[i],drop=FALSE]))
      colnames(RegressionData)<-c("ValueChange","Index")
      Regression<-lm(ValueChange ~ Index, RegressionData[-nrow(RegressionData),])
      PredictedRegion[Q,colnames(ValueChangeRegion)[i],"Forecasted"]<-Regression$coefficients[1]+Regression$coefficients[2]*RegressionData[nrow(RegressionData),2]
    }
  }
  for(i in 1:ncol(ValueChangeStage))
  {
    if(Q==FirstQuarter)
    {
      PotentialIndices<-names(Indices)[which(substr(names(Indices),4,5) == colnames(ValueChangeStage)[i] & rep(c(TRUE,FALSE),length(MarketIndices)))+1]
      SelectedIndicesStage<-c(SelectedIndicesStage,names(which.max(as.data.frame(cor(ValueChangeStage[QuartersNeeded[-1],colnames(ValueChangeStage)[i]],MarketChange[,PotentialIndices],use="pairwise.complete.obs")))))
    }
    if(!is.na(SelectedIndicesStage[i]))
    {
      RegressionData<-as.data.frame(cbind(ValueChangeStage[Quarters,colnames(ValueChangeStage)[i],drop=FALSE],MarketChange[Quarters,SelectedIndicesStage[i],drop=FALSE]))
      colnames(RegressionData)<-c("ValueChange","Index")
      Regression<-lm(ValueChange ~ Index, RegressionData[-nrow(RegressionData),])
      PredictedStage[Q,colnames(ValueChangeStage)[i],"Forecasted"]<-Regression$coefficients[1]+Regression$coefficients[2]*RegressionData[nrow(RegressionData),2]
    }
  }
  for(i in 1:ncol(ValueChangeSector))
  {
    if(Q==FirstQuarter)
    {
      PotentialIndices<-names(Indices)[which(substr(names(Indices),7,8) == colnames(ValueChangeSector)[i] & rep(c(TRUE,FALSE),length(MarketIndices)))+1]
      SelectedIndicesSector<-c(SelectedIndicesSector,names(which.max(as.data.frame(cor(ValueChangeSector[QuartersNeeded[-1],colnames(ValueChangeSector)[i],drop=FALSE],MarketChange[,PotentialIndices],use="pairwise.complete.obs")))))
    }
    if(!is.na(SelectedIndicesSector[i]))
    {
      RegressionData<-as.data.frame(cbind(ValueChangeSector[Quarters,colnames(ValueChangeSector)[i],drop=FALSE],MarketChange[Quarters,SelectedIndicesSector[i],drop=FALSE]))
      colnames(RegressionData)<-c("ValueChange","Index")
      Regression<-lm(ValueChange ~ Index, RegressionData[-nrow(RegressionData),])
      PredictedSector[Q,colnames(ValueChangeSector)[i],"Forecasted"]<-Regression$coefficients[1]+Regression$coefficients[2]*RegressionData[nrow(RegressionData),2]
    }
  }
  for(i in 1:ncol(ValueChangeRegionStage))
  {
    if(Q==FirstQuarter)
    {
      PotentialIndices<-names(Indices)[which(substr(names(Indices),1,2) == substr(colnames(ValueChangeRegionStage)[i],1,2) & substr(names(Indices),4,5) == substr(colnames(ValueChangeRegionStage)[i],4,5) & rep(c(TRUE,FALSE),length(MarketIndices)))+1]
      SelectedIndicesRegionStage<-c(SelectedIndicesRegionStage,names(which.max(as.data.frame(cor(ValueChangeRegionStage[QuartersNeeded[-1],colnames(ValueChangeRegionStage)[i]],MarketChange[,PotentialIndices],use="pairwise.complete.obs")))))
    }
    if(!is.na(SelectedIndicesRegionStage[i]))
    {
      RegressionData<-as.data.frame(cbind(ValueChangeRegionStage[Quarters,colnames(ValueChangeRegionStage)[i],drop=FALSE],MarketChange[Quarters,SelectedIndicesRegionStage[i],drop=FALSE]))
      colnames(RegressionData)<-c("ValueChange","Index")
      Regression<-lm(ValueChange ~ Index, RegressionData[-nrow(RegressionData),])
      PredictedRegionStage[Q,colnames(ValueChangeRegionStage)[i],"Forecasted"]<-Regression$coefficients[1]+Regression$coefficients[2]*RegressionData[nrow(RegressionData),2]
    }
  }
  for(i in 1:ncol(ValueChangeRegionSector))
  {
    if(Q==FirstQuarter)
    {
      PotentialIndices<-names(Indices)[which(substr(names(Indices),1,2) == substr(colnames(ValueChangeRegionSector)[i],1,2) & substr(names(Indices),7,8) == substr(colnames(ValueChangeRegionSector)[i],4,5) & rep(c(TRUE,FALSE),length(MarketIndices)))+1]
      SelectedIndicesRegionSector<-c(SelectedIndicesRegionSector,names(which.max(as.data.frame(cor(ValueChangeRegionSector[QuartersNeeded[-1],colnames(ValueChangeRegionSector)[i]],MarketChange[,PotentialIndices],use="pairwise.complete.obs")))))
    }
    if(!is.na(SelectedIndicesRegionSector[i]))
    {
      RegressionData<-as.data.frame(cbind(ValueChangeRegionSector[Quarters,colnames(ValueChangeRegionSector)[i],drop=FALSE],MarketChange[Quarters,SelectedIndicesRegionSector[i],drop=FALSE]))
      colnames(RegressionData)<-c("ValueChange","Index")
      if(sum(!is.na(RegressionData[-nrow(RegressionData),1]))>2)
      {
        Regression<-lm(ValueChange ~ Index, RegressionData[-nrow(RegressionData),])
        PredictedRegionSector[Q,colnames(ValueChangeRegionSector)[i],"Forecasted"]<-Regression$coefficients[1]+Regression$coefficients[2]*RegressionData[nrow(RegressionData),2]
      }
    }
  }
  for(i in 1:ncol(ValueChangeStageSector))
  {
    if(Q==FirstQuarter)
    {
      PotentialIndices<-names(Indices)[which(substr(names(Indices),4,5) == substr(colnames(ValueChangeStageSector)[i],1,2) & substr(names(Indices),7,8) == substr(colnames(ValueChangeStageSector)[i],4,5) & rep(c(TRUE,FALSE),length(MarketIndices)))+1]
      SelectedIndicesStageSector<-c(SelectedIndicesStageSector,names(which.max(as.data.frame(cor(ValueChangeStageSector[QuartersNeeded[-1],colnames(ValueChangeStageSector)[i]],MarketChange[,PotentialIndices],use="pairwise.complete.obs")))))
    }
    if(!is.na(SelectedIndicesStageSector[i]))
    {
      RegressionData<-as.data.frame(cbind(ValueChangeStageSector[Quarters,colnames(ValueChangeStageSector)[i],drop=FALSE],MarketChange[Quarters,SelectedIndicesStageSector[i],drop=FALSE]))
      colnames(RegressionData)<-c("ValueChange","Index")
      if(sum(!is.na(RegressionData[-nrow(RegressionData),1]))>2)
      {
        Regression<-lm(ValueChange ~ Index, RegressionData[-nrow(RegressionData),])
        PredictedStageSector[Q,colnames(ValueChangeStageSector)[i],"Forecasted"]<-Regression$coefficients[1]+Regression$coefficients[2]*RegressionData[nrow(RegressionData),2]
      }
    }
  }
  for(i in 1:ncol(ValueChangeBucket))
  {
    if(Q==FirstQuarter)
    {
      if(is.na(var(ValueChangeBucket[QuartersNeeded[-1],colnames(ValueChangeBucket)[i]],na.rm = TRUE)) | var(ValueChangeBucket[QuartersNeeded[-1],colnames(ValueChangeBucket)[i]],na.rm = TRUE)==0)
      {
        SelectedIndicesBucket<-c(SelectedIndicesBucket,NA)
      } else
      {
        PotentialIndices<-names(Indices)[which(substr(names(Indices),1,2) == substr(colnames(ValueChangeBucket)[i],1,2) & substr(names(Indices),4,5) == substr(colnames(ValueChangeBucket)[i],4,5) & substr(names(Indices),7,8) == substr(colnames(ValueChangeBucket)[i],7,8) & rep(c(TRUE,FALSE),length(MarketIndices)))+1]
        SelectedIndicesBucket<-c(SelectedIndicesBucket,names(which.max(as.data.frame(cor(ValueChangeBucket[QuartersNeeded[-1],colnames(ValueChangeBucket)[i]],MarketChange[,PotentialIndices],use="pairwise.complete.obs")))))
      }
    }
    if(!is.na(SelectedIndicesBucket[i]))
    {
      RegressionData<-as.data.frame(cbind(ValueChangeBucket[Quarters,colnames(ValueChangeBucket)[i],drop=FALSE],MarketChange[Quarters,SelectedIndicesBucket[i],drop=FALSE]))
      colnames(RegressionData)<-c("ValueChange","Index")
      if(sum(!is.na(RegressionData[-nrow(RegressionData),1]))>2)
      {
        Regression<-lm(ValueChange ~ Index, RegressionData[-nrow(RegressionData),])
        PredictedBucket[Q,colnames(ValueChangeBucket)[i],"Forecasted"]<-Regression$coefficients[1]+Regression$coefficients[2]*RegressionData[nrow(RegressionData),2]
      }
    }
  }
}

### Calculate Fitted Values ###
RegressionData<-as.data.frame(cbind(ValueChangeOverall[TestPeriod ,,drop=FALSE],MarketChange[TestPeriod,SelectedIndexOverall,drop=FALSE]))
colnames(RegressionData)<-c("ValueChange","Index")
Regression<-lm(ValueChange ~ Index, RegressionData)
PredictedOverall[,,"Fitted"]<-predict(Regression,RegressionData)
for(i in 1:ncol(ValueChangeRegion))
{
  RegressionData<-as.data.frame(cbind(ValueChangeRegion[TestPeriod,colnames(ValueChangeRegion)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesRegion[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  Regression<-lm(ValueChange ~ Index, RegressionData)
  PredictedRegion[,colnames(ValueChangeRegion)[i],"Fitted"]<-predict(Regression,RegressionData)
}
for(i in 1:ncol(ValueChangeStage))
{
  RegressionData<-as.data.frame(cbind(ValueChangeStage[TestPeriod,colnames(ValueChangeStage)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesStage[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  Regression<-lm(ValueChange ~ Index, RegressionData)
  PredictedStage[,colnames(ValueChangeStage)[i],"Fitted"]<-predict(Regression,RegressionData)
}
for(i in 1:ncol(ValueChangeSector))
{
  RegressionData<-as.data.frame(cbind(ValueChangeSector[TestPeriod,colnames(ValueChangeSector)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesSector[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  Regression<-lm(ValueChange ~ Index, RegressionData)
  PredictedSector[,colnames(ValueChangeSector)[i],"Fitted"]<-predict(Regression,RegressionData)
}
for(i in 1:ncol(ValueChangeRegionStage))
{
  RegressionData<-as.data.frame(cbind(ValueChangeRegionStage[TestPeriod,colnames(ValueChangeRegionStage)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesRegionStage[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  Regression<-lm(ValueChange ~ Index, RegressionData)
  PredictedRegionStage[,colnames(ValueChangeRegionStage)[i],"Fitted"]<-predict(Regression,RegressionData)
}
for(i in 1:ncol(ValueChangeRegionSector))
{
  RegressionData<-as.data.frame(cbind(ValueChangeRegionSector[TestPeriod,colnames(ValueChangeRegionSector)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesRegionSector[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  if(sum(!is.na(RegressionData[,1]))>2)
  {
    Regression<-lm(ValueChange ~ Index, RegressionData)
    PredictedRegionSector[,colnames(ValueChangeRegionSector)[i],"Fitted"]<-predict(Regression,RegressionData)
  }
}
for(i in 1:ncol(ValueChangeStageSector))
{
  RegressionData<-as.data.frame(cbind(ValueChangeStageSector[TestPeriod,colnames(ValueChangeStageSector)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesStageSector[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  if(sum(!is.na(RegressionData[-nrow(RegressionData),1]))>2)
  {
    Regression<-lm(ValueChange ~ Index, RegressionData)
    PredictedStageSector[,colnames(ValueChangeStageSector)[i],"Fitted"]<-predict(Regression,RegressionData)
  }
}
for(i in 1:ncol(ValueChangeBucket))
{
  if(!is.na(SelectedIndicesBucket[i]))
  {
    RegressionData<-as.data.frame(cbind(ValueChangeBucket[TestPeriod,colnames(ValueChangeBucket)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesBucket[i],drop=FALSE]))
    colnames(RegressionData)<-c("ValueChange","Index")
    if(sum(!is.na(RegressionData[,1]))>2)
    {
      Regression<-lm(ValueChange ~ Index, RegressionData)
      PredictedBucket[,colnames(ValueChangeBucket)[i],"Fitted"]<-predict(Regression,RegressionData)
    }
  }
}

### Calculate Prediction Errors per UPC ###    
TypeofPrediction<-c("Fitted","Forecasted")[2]
ErrorType<-2                                    ### 1 is absolute errors, 2 is squared errors. We use squared errors as most commonly used. 
PredictedErrors<-UPCValueChange[,1:11]
PredictedErrors[,-c(1:3)]<-NA
colnames(PredictedErrors)[-c(1:3)]<-c("Overall OLS","Region OLS","Stage OLS","Sector OLS","RegionStage OLS","RegionSector OLS","StageSector OLS","Bucket OLS")
PredictedErrorsCount<-PredictedErrors

### Max value is cut-off with which we minimize the RMSE of forecasted values so we believe we capture systemic value change. ####
MaxValueChange<-0.24
UPCValueChange<-UPCValueCreation/UPCMV[,-c(1,ncol(UPCMV))]
UPCValueChange[abs(UPCValueChange)>MaxValueChange]<-NA
UPCValueChange<-cbind(UPCMapping[,-1],UPCValueChange)

PredictedErrors[,"Overall OLS"]<-rowSums(abs(t(replicate(nrow(UPCValueChange),PredictedOverall[,,TypeofPrediction]))-UPCValueChange[,TestPeriod])^ErrorType,na.rm = TRUE)
PredictedErrorsCount[,"Overall OLS"]<-rowSums(!is.na(t(replicate(nrow(UPCValueChange),PredictedOverall[,,TypeofPrediction]))-UPCValueChange[,TestPeriod]))
PredictedErrors[,"Region OLS"]<-rowSums(abs(t(PredictedRegion[,match(UPCValueChange$Region,colnames(PredictedRegion)),TypeofPrediction])-UPCValueChange[,TestPeriod])^ErrorType,na.rm=TRUE)
PredictedErrorsCount[,"Region OLS"]<-rowSums(!is.na(t(PredictedRegion[,match(UPCValueChange$Region,colnames(PredictedRegion)),TypeofPrediction])-UPCValueChange[,TestPeriod]))
PredictedErrors[,"Stage OLS"]<-rowSums(abs(t(PredictedStage[,match(UPCValueChange$Stage,colnames(PredictedStage)),TypeofPrediction])-UPCValueChange[,TestPeriod])^ErrorType,na.rm=TRUE)
PredictedErrorsCount[,"Stage OLS"]<-rowSums(!is.na(t(PredictedStage[,match(UPCValueChange$Stage,colnames(PredictedStage)),TypeofPrediction])-UPCValueChange[,TestPeriod]))
PredictedErrors[,"Sector OLS"]<-rowSums(abs(t(PredictedSector[,match(UPCValueChange$Sector,colnames(PredictedSector)),TypeofPrediction])-UPCValueChange[,TestPeriod])^ErrorType,na.rm=TRUE)
PredictedErrorsCount[,"Sector OLS"]<-rowSums(!is.na(t(PredictedSector[,match(UPCValueChange$Sector,colnames(PredictedSector)),TypeofPrediction])-UPCValueChange[,TestPeriod]))
PredictedErrors[,"RegionStage OLS"]<-rowSums(abs(t(PredictedRegionStage[,match(paste(UPCValueChange$Region,UPCValueChange$Stage),colnames(PredictedRegionStage)),TypeofPrediction])-UPCValueChange[,TestPeriod])^ErrorType,na.rm=TRUE)
PredictedErrorsCount[,"RegionStage OLS"]<-rowSums(!is.na(t(PredictedRegionStage[,match(paste(UPCValueChange$Region,UPCValueChange$Stage),colnames(PredictedRegionStage)),TypeofPrediction])-UPCValueChange[,TestPeriod]))
PredictedErrors[,"RegionSector OLS"]<-rowSums(abs(t(PredictedRegionSector[,match(paste(UPCValueChange$Region,UPCValueChange$Sector),colnames(PredictedRegionSector)),TypeofPrediction])-UPCValueChange[,TestPeriod])^ErrorType,na.rm=TRUE)
PredictedErrorsCount[,"RegionSector OLS"]<-rowSums(!is.na(t(PredictedRegionSector[,match(paste(UPCValueChange$Region,UPCValueChange$Sector),colnames(PredictedRegionSector)),TypeofPrediction])-UPCValueChange[,TestPeriod]))
PredictedErrors[,"StageSector OLS"]<-rowSums(abs(t(PredictedStageSector[,match(paste(UPCValueChange$Stage,UPCValueChange$Sector),colnames(PredictedStageSector)),TypeofPrediction])-UPCValueChange[,TestPeriod])^ErrorType,na.rm=TRUE)
PredictedErrorsCount[,"StageSector OLS"]<-rowSums(!is.na(t(PredictedStageSector[,match(paste(UPCValueChange$Stage,UPCValueChange$Sector),colnames(PredictedStageSector)),TypeofPrediction])-UPCValueChange[,TestPeriod]))
PredictedErrors[,"Bucket OLS"]<-rowSums(abs(t(PredictedBucket[,match(paste(UPCValueChange$Region,UPCValueChange$Stage,UPCValueChange$Sector),colnames(PredictedBucket)),TypeofPrediction])-UPCValueChange[,TestPeriod])^ErrorType,na.rm=TRUE)
PredictedErrorsCount[,"Bucket OLS"]<-rowSums(!is.na(t(PredictedBucket[,match(paste(UPCValueChange$Region,UPCValueChange$Stage,UPCValueChange$Sector),colnames(PredictedBucket)),TypeofPrediction])-UPCValueChange[,TestPeriod]))

### Calculate Prediction Errors per bucket ###
SumPredictedErrorsBucket<-aggregate(PredictedErrors[,-c(1:3)], by=list(PredictedErrors$Region,PredictedErrors$Stage,PredictedErrors$Sector),sum,na.rm = TRUE)
PredictedErrorsCountBucket<-aggregate(PredictedErrorsCount[,-c(1:3)], by=list(PredictedErrorsCount$Region,PredictedErrorsCount$Stage,PredictedErrorsCount$Sector),sum)
SumPredictedErrorsBucket[SumPredictedErrorsBucket==0]<-NA
rownames(SumPredictedErrorsBucket)<-paste(SumPredictedErrorsBucket[,1],SumPredictedErrorsBucket[,2],SumPredictedErrorsBucket[,3])
ErrorsMetricBucket<-(SumPredictedErrorsBucket[,-c(1:3)]/PredictedErrorsCountBucket[,-c(1:3)])^(1/ErrorType)           # The power expression is needed to transform to root mean squared error.
ErrorsMetricBucket[is.na(apply(ErrorsMetricBucket,1,FUN=which.min)==0),1]<-0                                          # If all predictions method are NA set overall to 0 such that overall becomes the selected estimation method.
#write.csv(ErrorsMetricBucket,"RMSEBuckettest.csv")

### Determine Final Predictions per bucket ###
FinalPredictionGroup<-matrix("Overall",nrow(ErrorsMetricBucket),1,dimnames=list(rownames(ErrorsMetricBucket),"Final Group"))
FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==2,]<-substr(rownames(FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==2,,drop=FALSE]),1,2)
FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==3,]<-substr(rownames(FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==3,,drop=FALSE]),4,5)
FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==4,]<-substr(rownames(FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==4,,drop=FALSE]),7,8)
FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==5,]<-substr(rownames(FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==5,,drop=FALSE]),1,5)
FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==6,]<-paste(substr(rownames(FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==6,,drop=FALSE]),1,2),substr(rownames(FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==6,,drop=FALSE]),7,8))
FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==7,]<-substr(rownames(FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==7,,drop=FALSE]),4,8)
FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==8,]<-substr(rownames(FinalPredictionGroup[as.matrix(apply(ErrorsMetricBucket,1,FUN=which.min))==8,,drop=FALSE]),1,8)

FinalPredictionGroupAll<-matrix("Overall",88,1,dimnames=list(paste(expand.grid(colnames(PredictedRegion),colnames(PredictedStage),colnames(PredictedSector))[,1],expand.grid(colnames(PredictedRegion),colnames(PredictedStage),colnames(PredictedSector))[,2],expand.grid(colnames(PredictedRegion),colnames(PredictedStage),colnames(PredictedSector))[,3]),"Final Group"))
FinalPredictionGroupAll[rownames(FinalPredictionGroupAll) %in% rownames(FinalPredictionGroup),]<-FinalPredictionGroup

PredictedAll<-t(cbind(PredictedOverall[,,TypeofPrediction],PredictedRegion[,,TypeofPrediction],PredictedStage[,,TypeofPrediction],PredictedSector[,,TypeofPrediction],PredictedRegionStage[,,TypeofPrediction],PredictedRegionSector[,,TypeofPrediction],PredictedStageSector[,,TypeofPrediction],PredictedBucket[,,TypeofPrediction]))
rownames(PredictedAll)[1]<-"Overall"
FinalPredictions<-PredictedAll[match(FinalPredictionGroup,rownames(PredictedAll)),]
rownames(FinalPredictions)<-rownames(FinalPredictionGroup)

### Calculate predicted Value Creation vs Actual per UPC by applying estimated value change on prior quarter MV ###
PredictedValueChangeperUPC<-FinalPredictions[match(paste(UPCMapping[,2],UPCMapping[,3],UPCMapping[,4]),rownames(FinalPredictions)),]
PredictedValueCreationperUPC<-PredictedValueChangeperUPC*UPCMV[,QuartersVector[(match(TestPeriod,as.vector(QuartersVector[,]))-1),]]
colnames(PredictedValueCreationperUPC)<-TestPeriod

### Plot value creation actual vs predicted ###
#par(mar=c(1,1,1,1))
#par(mar=c(5.1,4.1,4.1,2.1))
#plot(1:length(TestPeriod),colSums(UPCValueCreation[,TestPeriod],na.rm=TRUE),xaxt="n",xlab="Test Period",ylab="Value Creation",main="Value Creation Actual vs Predicted")
#axis(1,at=1:length(TestPeriod),labels=TestPeriod)
#lines(colSums(PredictedValueCreationperUPC,na.rm=TRUE),lwd=2)
#lines(rep(0,length(TestPeriod)))
#legend('bottomleft',legend = c("MAF","No MAF"),lty=1,lwd=c(2,1),cex=1,col=c(1,1),xjust=1)


############################################################################################################################################################################
### Generate Output files ###

MarketChangeLongerPeriod<-MarketChangeFormula(LastDayofQuarter,MarketIndices,as.vector(QuartersVector[1:87,]),DaysBeforeCQ,DaysBeforePQ,DaysAfterPQ)

RegressionOverall<-cbind.data.frame(matrix(NA,1,14,dimnames = list(c("Overall"),c("Constant","Std. Error Constant","p-value Constant","Parameter","Std. Error Parameter","p-value Parameter","R2","DW Statistic","DW p-value","BP Statistic","BP p-value","DF p-value ValueChange","DF p-value Index","DF p-value Index extended"))),matrix(SelectedIndexOverall,length(SelectedIndexOverall),1,dimnames = list(c("Overall"),c("Index"))))
RegressionRegion<-cbind.data.frame(matrix(NA,4,14,dimnames = list(colnames(PredictedRegion),c("Constant","Std. Error Constant","p-value Constant","Parameter","Std. Error Parameter","p-value Parameter","R2","DW Statistic","DW p-value","BP Statistic","BP p-value","DF p-value ValueChange","DF p-value Index","DF p-value Index extended"))),matrix(SelectedIndicesRegion,length(SelectedIndicesRegion),1,dimnames = list(colnames(PredictedRegion),c("Index"))))
RegressionStage<-cbind.data.frame(matrix(NA,2,14,dimnames = list(colnames(PredictedStage),c("Constant","Std. Error Constant","p-value Constant","Parameter","Std. Error Parameter","p-value Parameter","R2","DW Statistic","DW p-value","BP Statistic","BP p-value","DF p-value ValueChange","DF p-value Index","DF p-value Index extended"))),matrix(SelectedIndicesStage,length(SelectedIndicesStage),1,dimnames = list(colnames(PredictedStage),c("Index"))))
RegressionSector<-cbind.data.frame(matrix(NA,11,14,dimnames = list(colnames(PredictedSector),c("Constant","Std. Error Constant","p-value Constant","Parameter","Std. Error Parameter","p-value Parameter","R2","DW Statistic","DW p-value","BP Statistic","BP p-value","DF p-value ValueChange","DF p-value Index","DF p-value Index extended"))),matrix(SelectedIndicesSector,length(SelectedIndicesSector),1,dimnames = list(colnames(PredictedSector),c("Index"))))
RegressionRegionStage<-cbind.data.frame(matrix(NA,8,14,dimnames = list(colnames(PredictedRegionStage),c("Constant","Std. Error Constant","p-value Constant","Parameter","Std. Error Parameter","p-value Parameter","R2","DW Statistic","DW p-value","BP Statistic","BP p-value","DF p-value ValueChange","DF p-value Index","DF p-value Index extended"))),matrix(SelectedIndicesRegionStage,length(SelectedIndicesRegionStage),1,dimnames = list(colnames(PredictedRegionStage),c("Index"))))
RegressionRegionSector<-cbind.data.frame(matrix(NA,44,14,dimnames = list(colnames(PredictedRegionSector),c("Constant","Std. Error Constant","p-value Constant","Parameter","Std. Error Parameter","p-value Parameter","R2","DW Statistic","DW p-value","BP Statistic","BP p-value","DF p-value ValueChange","DF p-value Index","DF p-value Index extended"))),matrix(SelectedIndicesRegionSector,length(SelectedIndicesRegionSector),1,dimnames = list(colnames(PredictedRegionSector),c("Index"))))
RegressionStageSector<-cbind.data.frame(matrix(NA,22,14,dimnames = list(colnames(PredictedStageSector),c("Constant","Std. Error Constant","p-value Constant","Parameter","Std. Error Parameter","p-value Parameter","R2","DW Statistic","DW p-value","BP Statistic","BP p-value","DF p-value ValueChange","DF p-value Index","DF p-value Index extended"))),matrix(SelectedIndicesStageSector,length(SelectedIndicesStageSector),1,dimnames = list(colnames(PredictedStageSector),c("Index"))))
RegressionBucket<-cbind.data.frame(matrix(NA,ncol(PredictedBucket),14,dimnames = list(colnames(PredictedBucket),c("Constant","Std. Error Constant","p-value Constant","Parameter","Std. Error Parameter","p-value Parameter","R2","DW Statistic","DW p-value","BP Statistic","BP p-value","DF p-value ValueChange","DF p-value Index","DF p-value Index extended"))),matrix(SelectedIndicesBucket,length(SelectedIndicesBucket),1,dimnames = list(colnames(PredictedBucket),c("Index"))))

RegressionData<-as.data.frame(cbind(ValueChangeOverall[TestPeriod ,,drop=FALSE],MarketChange[TestPeriod,SelectedIndexOverall,drop=FALSE]))
colnames(RegressionData)<-c("ValueChange","Index")
Regression<-lm(ValueChange ~ Index, RegressionData)
RegressionOverall[1,c(1:11,13)]<-c(summary(Regression)$coefficients[1,-3],summary(Regression)$coefficients[2,-3],summary(Regression)$r.squared,dwtest(Regression)$statistic,dwtest(Regression)$p.value,bptest(Regression)$statistic,bptest(Regression)$p.value,adf.test(RegressionData[,2])$p.value)
if(sum(is.na(RegressionData[,1]))==0){RegressionOverall[1,12]<-adf.test(RegressionData[,1])$p.value}
if(sum(is.na(MarketChangeLongerPeriod[,SelectedIndexOverall,drop=FALSE]))==0){RegressionOverall[1,14]<-adf.test(MarketChangeLongerPeriod[,SelectedIndexOverall,drop=FALSE])$p.value}
for(i in 1:nrow(RegressionRegion))
{
  RegressionData<-as.data.frame(cbind(ValueChangeRegion[TestPeriod,colnames(ValueChangeRegion)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesRegion[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  Regression<-lm(ValueChange ~ Index, RegressionData)
  RegressionRegion[i,c(1:11,13)]<-c(summary(Regression)$coefficients[1,-3],summary(Regression)$coefficients[2,-3],summary(Regression)$r.squared,dwtest(Regression)$statistic,dwtest(Regression)$p.value,bptest(Regression)$statistic,bptest(Regression)$p.value,adf.test(RegressionData[,2])$p.value)
  if(sum(is.na(RegressionData[,1]))==0){RegressionRegion[i,12]<-adf.test(RegressionData[,1])$p.value}
  if(sum(is.na(MarketChangeLongerPeriod[,SelectedIndicesRegion[i],drop=FALSE]))==0){RegressionRegion[i,14]<-adf.test(MarketChangeLongerPeriod[,SelectedIndicesRegion[i],drop=FALSE])$p.value}
}
for(i in 1:nrow(RegressionStage))
{
  RegressionData<-as.data.frame(cbind(ValueChangeStage[TestPeriod,colnames(ValueChangeStage)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesStage[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  Regression<-lm(ValueChange ~ Index, RegressionData)
  RegressionStage[i,c(1:11,13)]<-c(summary(Regression)$coefficients[1,-3],summary(Regression)$coefficients[2,-3],summary(Regression)$r.squared,dwtest(Regression)$statistic,dwtest(Regression)$p.value,bptest(Regression)$statistic,bptest(Regression)$p.value,adf.test(RegressionData[,2])$p.value)
  if(sum(is.na(RegressionData[,1]))==0){RegressionStage[i,12]<-adf.test(RegressionData[,1])$p.value}
  if(sum(is.na(MarketChangeLongerPeriod[,SelectedIndicesStage[i],drop=FALSE]))==0){RegressionStage[i,14]<-adf.test(MarketChangeLongerPeriod[,SelectedIndicesStage[i],drop=FALSE])$p.value}
}
for(i in 1:nrow(RegressionSector))
{
  RegressionData<-as.data.frame(cbind(ValueChangeSector[TestPeriod,colnames(ValueChangeSector)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesSector[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  Regression<-lm(ValueChange ~ Index, RegressionData)
  RegressionSector[i,c(1:11,13)]<-c(summary(Regression)$coefficients[1,-3],summary(Regression)$coefficients[2,-3],summary(Regression)$r.squared,dwtest(Regression)$statistic,dwtest(Regression)$p.value,bptest(Regression)$statistic,bptest(Regression)$p.value,adf.test(RegressionData[,2])$p.value)
  if(sum(is.na(RegressionData[,1]))==0){RegressionSector[i,12]<-adf.test(RegressionData[,1])$p.value}
  if(sum(is.na(MarketChangeLongerPeriod[,SelectedIndicesSector[i],drop=FALSE]))==0){RegressionSector[i,14]<-adf.test(MarketChangeLongerPeriod[,SelectedIndicesSector[i],drop=FALSE])$p.value}
}
for(i in 1:nrow(RegressionRegionStage))
{
  RegressionData<-as.data.frame(cbind(ValueChangeRegionStage[TestPeriod,colnames(ValueChangeRegionStage)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesRegionStage[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  Regression<-lm(ValueChange ~ Index, RegressionData)
  RegressionRegionStage[i,c(1:11,13)]<-c(summary(Regression)$coefficients[1,-3],summary(Regression)$coefficients[2,-3],summary(Regression)$r.squared,dwtest(Regression)$statistic,dwtest(Regression)$p.value,bptest(Regression)$statistic,bptest(Regression)$p.value,adf.test(RegressionData[,2])$p.value)
  if(sum(is.na(RegressionData[,1]))==0){RegressionRegionStage[i,12]<-adf.test(RegressionData[,1])$p.value}
  if(sum(is.na(MarketChangeLongerPeriod[,SelectedIndicesRegionStage[i],drop=FALSE]))==0){RegressionRegionStage[i,14]<-adf.test(MarketChangeLongerPeriod[,SelectedIndicesRegionStage[i],drop=FALSE])$p.value}
}
for(i in 1:nrow(RegressionRegionSector))
{
  RegressionData<-as.data.frame(cbind(ValueChangeRegionSector[TestPeriod,colnames(ValueChangeRegionSector)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesRegionSector[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  if(sum(!is.na(RegressionData[,1]))>2)
  {
    Regression<-lm(ValueChange ~ Index, RegressionData)
    RegressionRegionSector[i,c(1:11,13)]<-c(summary(Regression)$coefficients[1,-3],summary(Regression)$coefficients[2,-3],summary(Regression)$r.squared,dwtest(Regression)$statistic,dwtest(Regression)$p.value,bptest(Regression)$statistic,bptest(Regression)$p.value,adf.test(RegressionData[,2])$p.value)
    if(sum(is.na(RegressionData[,1]))==0){RegressionRegionSector[i,12]<-adf.test(RegressionData[,1])$p.value}
  }
}
for(i in 1:nrow(RegressionStageSector))
{
  RegressionData<-as.data.frame(cbind(ValueChangeStageSector[TestPeriod,colnames(ValueChangeStageSector)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesStageSector[i],drop=FALSE]))
  colnames(RegressionData)<-c("ValueChange","Index")
  if(sum(!is.na(RegressionData[-nrow(RegressionData),1]))>2)
  {
    Regression<-lm(ValueChange ~ Index, RegressionData)
    RegressionStageSector[i,c(1:11,13)]<-c(summary(Regression)$coefficients[1,-3],summary(Regression)$coefficients[2,-3],summary(Regression)$r.squared,dwtest(Regression)$statistic,dwtest(Regression)$p.value,bptest(Regression)$statistic,bptest(Regression)$p.value,adf.test(RegressionData[,2])$p.value)
    if(sum(is.na(RegressionData[,1]))==0){RegressionStageSector[i,12]<-adf.test(RegressionData[,1])$p.value}
    if(sum(is.na(MarketChangeLongerPeriod[,SelectedIndicesStageSector[i],drop=FALSE]))==0){RegressionStageSector[i,14]<-adf.test(MarketChangeLongerPeriod[,SelectedIndicesStageSector[i],drop=FALSE])$p.value}
  }
}
for(i in 1:nrow(RegressionBucket))
{
  if(!is.na(SelectedIndicesBucket[i]))
  {
    RegressionData<-as.data.frame(cbind(ValueChangeBucket[TestPeriod,colnames(ValueChangeBucket)[i],drop=FALSE],MarketChange[TestPeriod,SelectedIndicesBucket[i],drop=FALSE]))
    colnames(RegressionData)<-c("ValueChange","Index")
    if(sum(!is.na(RegressionData[,1]))>1 & var(RegressionData[,1],na.rm=TRUE)!=0)
    {
      Regression<-lm(ValueChange ~ Index, RegressionData)
      RegressionBucket[i,c(1:11,13)]<-c(summary(Regression)$coefficients[1,-3],summary(Regression)$coefficients[2,-3],summary(Regression)$r.squared,dwtest(Regression)$statistic,dwtest(Regression)$p.value,bptest(Regression)$statistic,bptest(Regression)$p.value,adf.test(RegressionData[,2])$p.value)
      if(sum(is.na(RegressionData[,1]))==0){RegressionBucket[i,12]<-adf.test(RegressionData[,1])$p.value}
      if(sum(is.na(MarketChangeLongerPeriod[,SelectedIndicesBucket[i],drop=FALSE]))==0){RegressionBucket[i,14]<-adf.test(MarketChangeLongerPeriod[,SelectedIndicesBucket[i],drop=FALSE])$p.value}
    }
  }
}

RegressionsAll<-rbind(RegressionOverall,RegressionRegion,RegressionStage,RegressionSector,RegressionRegionStage,RegressionRegionSector,RegressionStageSector,RegressionBucket)
FinalRegressions<-cbind(FinalPredictionGroupAll,RegressionsAll[match(FinalPredictionGroupAll,rownames(RegressionsAll)),])
rownames(FinalRegressions)<-rownames(FinalPredictionGroupAll)

### Calculate Generic Regressions ###
RegressionGeneric<-as.data.frame(matrix(NA,4,16,dimnames = list(c("Generic US BO","Generic EU BO","Generic VC","Generic ROW"),c("Final Group","Constant","Std. Error Constant","p-value Constant","Parameter","Std. Error Parameter","p-value Parameter","R2","DW Statistic","DW p-value","BP Statistic","BP p-value","DF p-value ValueChange","DF p-value Index","DF p-value Index extended","Index"))))
RegressionGeneric["Generic US BO",2:16]<-RegressionsAll["US BO",]
RegressionGeneric["Generic EU BO",2:16]<-RegressionsAll["EU BO",] 

### Value Change VC ###
ValueChangeVCUnweighted<-t(t(colMeans(UPCValueChange[UPCValueChange$Stage=="VC" & (UPCValueChange$Region=="US" | UPCValueChange$Region=="EU"),-c(1:3)],na.rm = TRUE)))
colnames(ValueChangeVCUnweighted)<-"VC"
ValueChangeVCCapitalWeighted<-t(t(colSums(UPCValueCreation[UPCMapping$Stage=="VC" & (UPCMapping$Region=="US" | UPCMapping$Region=="EU"),],na.rm=TRUE)/colSums(UPCMV[UPCMapping$Stage=="VC" & (UPCMapping$Region=="US" | UPCMapping$Region=="EU"),-c(1,ncol(UPCMV))],na.rm=TRUE)))
colnames(ValueChangeVCCapitalWeighted)<-colnames(ValueChangeVCUnweighted)

### Value Change ROW ###
ValueChangeROWUnweighted<-t(t(colMeans(UPCValueChange[UPCValueChange$Region=="AA" | UPCValueChange$Region=="RE",-c(1:3)],na.rm = TRUE)))
colnames(ValueChangeROWUnweighted)<-"ROW"
ValueChangeROWCapitalWeighted<-t(t(colSums(UPCValueCreation[UPCMapping$Region=="AA" | UPCMapping$Region=="RE",],na.rm=TRUE)/colSums(UPCMV[UPCMapping$Region=="AA" | UPCMapping$Region=="RE",-c(1,ncol(UPCMV))],na.rm=TRUE)))
colnames(ValueChangeROWCapitalWeighted)<-colnames(ValueChangeROWUnweighted)

ValueChangeVC<-if(CapitalWeighted){ValueChangeVCCapitalWeighted}else{ValueChangeVCUnweighted}
ValueChangeROW<-if(CapitalWeighted){ValueChangeROWCapitalWeighted}else{ValueChangeROWUnweighted}

PotentialIndices<-names(Indices)[which(substr(names(Indices),1,2) == "US" | substr(names(Indices),1,2) == "EU" & substr(names(Indices),4,5) == "VC" & rep(c(TRUE,FALSE),length(MarketIndices)))+1]
SelectedIndexVC<-names(which.max(as.data.frame(cor(ValueChangeVC[QuartersNeeded[-1],],MarketChange[,PotentialIndices],use="pairwise.complete.obs"))))
RegressionData<-as.data.frame(cbind(ValueChangeVC[TestPeriod ,,drop=FALSE],MarketChange[TestPeriod,SelectedIndexVC,drop=FALSE]))
colnames(RegressionData)<-c("ValueChange","Index")
Regression<-lm(ValueChange ~ Index, RegressionData)
RegressionGeneric["Generic VC",c(2:14,16)]<-c(summary(Regression)$coefficients[1,-3],summary(Regression)$coefficients[2,-3],summary(Regression)$r.squared,dwtest(Regression)$statistic,dwtest(Regression)$p.value,bptest(Regression)$statistic,bptest(Regression)$p.value,adf.test(RegressionData[,1])$p.value,adf.test(RegressionData[,2])$p.value,SelectedIndexVC)
if(sum(is.na(MarketChangeLongerPeriod[,SelectedIndexVC,drop=FALSE]))==0){RegressionGeneric["Generic VC",15]<-adf.test(MarketChangeLongerPeriod[,SelectedIndexVC,drop=FALSE])$p.value}

PotentialIndices<-names(Indices)[which(substr(names(Indices),1,2) == "AA" | substr(names(Indices),1,2) == "RE"  & rep(c(TRUE,FALSE),length(MarketIndices)))+1]
SelectedIndexROW<-names(which.max(as.data.frame(cor(ValueChangeROW[QuartersNeeded[-1],],MarketChange[,PotentialIndices],use="pairwise.complete.obs"))))
RegressionData<-as.data.frame(cbind(ValueChangeROW[TestPeriod ,,drop=FALSE],MarketChange[TestPeriod,SelectedIndexROW,drop=FALSE]))
colnames(RegressionData)<-c("ValueChange","Index")
Regression<-lm(ValueChange ~ Index, RegressionData)
RegressionGeneric["Generic ROW",2:16]<-c(summary(Regression)$coefficients[1,-3],summary(Regression)$coefficients[2,-3],summary(Regression)$r.squared,dwtest(Regression)$statistic,dwtest(Regression)$p.value,bptest(Regression)$statistic,bptest(Regression)$p.value,adf.test(RegressionData[,1])$p.value,adf.test(RegressionData[,2])$p.value,adf.test(MarketChangeLongerPeriod[-c(1:14),SelectedIndexROW,drop=FALSE])$p.value,SelectedIndexROW)

FinalRegressionsAll<-rbind(FinalRegressions,RegressionGeneric)
FinalRegressionsSQL<-cbind(FinalRegressionsAll[,c("Final Group","Constant","Parameter","R2","Index")],FIGI[match(FinalRegressionsAll[,"Index"],FIGI[,"Index"]),"FIGI"],DaysBeforeCQ,FIGI[match(FinalRegressionsAll[,"Index"],FIGI[,"Index"]),"Currency"])
colnames(FinalRegressionsSQL)<-c("Final Group", "Constant","Parameter","R2","Index","FIGI","Number of Days","Currency")


write.csv(FinalRegressionsSQL,paste("FinalRegressionsSQL",LastQuarter,format(Sys.time(),"%Y%m%d_%H%M%S"),".csv"))



