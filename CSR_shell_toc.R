########################################################################
# Author: stata93
# Original Date: 2022-01-01
# Last Modified Date: 2022-08-31
# Description: generate ToC Excel sheet all fn code in a cell + fn sheet
########################################################################
setwd("E:\\R_study\\ClinicalCSR-test\\abreviatedCSR")
# paths, files, ...
if(TRUE){
  rm(list=ls())
  who.maintain<-"stata93 <stata93@163.com"

  dir(loapath <- projpath<- getwd())
  progfile <- 'CSR_shell_toc.R'
  loafile <- "CSR_tfl_LoA.xlsx"

  outpath<-'shells_toc'
  outf <- gsub('LoA', 'TOC', loafile, fixed=T) #tranwrd function

  #the first 2 characters in footnote code
  fncd <- "FN%02d"
}

#only use the run this file in simple way
if(F){
  dev.off()
  dir()
  source(progfile)
}

#libraries,functions, ...
if(TRUE){
  library(haven)
  library(dplyr)
  library(stringr)
  library(openxlsx)
}

#generate tlf numbers
if(TRUE){
  #import loa sheet
  loa <- openxlsx::read.xlsx(file.path(loapath, loafile), sheet='loa')
  #keep only rows containing TLFs
  loa <- loa %>%
    mutate(X1=substr(SECTION1,1,1))%>%
    filter(X1 %in% c("T","L","F"))
  loa$tlfnumber <- loa$SECTION1
  fn <-as.list(loa$ID)
  lengthfn <- length(fn)
  stN <- vector(mode = "list", length = lengthfn) #starting position of footnote
  edN <- vector(mode = "list", length = lengthfn) #ending position of footnote
  ffN <- vector(mode = "list", length = lengthfn) #a empty list to store footnote
  ffM <- vector(mode = "list", length = lengthfn)
}

if(TRUE){
  #extract all footnote
  for(i in 1:nrow(loa)){
    id0<- loa$ID[i]     #get TLF template id
    print( paste(i,': ', id0,'is completed'))  #show log
    fn[[i]] <- openxlsx::read.xlsx(file.path(loapath, loafile), sheet=id0)
    stN[[i]] <- grep('note:', fn[[i]][,1])
    edN[[i]]<-grep('--------', fn[[i]][,2])-1
    ffN[[i]]<-fn[[i]][stN[[i]]:edN[[i]], 2]
  }

  #remove duplicates and NAs
  fttext <- unique(unlist(ffN))
  fttext <- fttext[!is.na(fttext)]
  #check if number of characters exceed 1000 in footnote
  if(is.element(TRUE, nchar(fttext)>1000)){
    print("There are footnote(s) exceed 1000 characterl limit")
  }

  #generate footnote code column
  if(TRUE) {
    df_ft <- data.frame(FTCODE = sprintf(fncd, 1:length(fttext)),
                        FTTEXT = fttext)
    for (i in 1:length(ffN)) {
      ffM[[i]]<- match(ffN[[i]],fttext)
    }
    #remove all NAs
    ffM<- lapply(ffM, function(x) x[!is.na(x)])
    #convert number to bmxx format
    for (i in 1:length(ffM)){
      ffM[[i]] <- sprintf(fncd, ffM[[i]])
      ffM[[i]]<- paste(ffM[[i]], collapse=', ')
    }
  }

  #generate other columns in TOC
  if(TRUE){
    toc<- loa%>%
      select(X1,TITLE1,TITLE3,tlfnumber) %>%
      rename(TYPE = X1)%>%
      mutate(SECTION1=str_sub(tlfnumber,1,-7),
             SECTION2=str_sub(tlfnumber,-5,-5),
             SECTION3=str_sub(tlfnumber,-3,-3),
             SECTION4=str_sub(tlfnumber,-1,-1))
    for(i in 1:nrow(toc)){
      pp <- toc$TITLE3[i]
      tt1 <- toc$TITLE1[i]
      tt1<- gsub('<<','',tt1)
      tt1 <-gsub('>>','',tt1)
      tt1 <- gsub('in [[Population]]','',tt1,fixed=TRUE)
      toc$TITLE1[i]<- tt1
    }
    #add populations to title
    toc$TITLE1 <- paste(toc$TITLE1, toc$TITLE3)
  }
  a<-unlist(ffM,recursive=FALSE);   # str(a)
  #merge FTCODE column into toc
  toc1 <-cbind(toc,unlist(ffM))
  toc1<-toc1 %>%
    rename(FTCODE=`unlist(ffM)`)
  #reorder the columns
  toc1 <- toc1[c("SECTION1", "SECTION2", "SECTION3",
                 "SECTION4","tlfnumber", "TYPE","TITLE1","TITLE3","FTCODE")]

  #output
  list_of_datasets <- list("TOC" = toc1, "Footnote" = df_ft)

  write.xlsx(list_of_datasets, file = file.path(outpath,outf),
             asTable=TRUE, colWidths = c("auto", "auto"), withFilter = c(T, T))
}

################################################
print( paste('The TOC tables are generated in the xlsx file.'))

