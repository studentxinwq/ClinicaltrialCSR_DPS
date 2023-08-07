###########################################################################################
# Author: stata93
# Original Date: 2022-3-22
# Last Modified Date: 2022-08-31
# Description: convert mockup shells into an integrated RTF file
###########################################################################################

## To debug:  loa[i,]
## Alert: before running this code, check the loa table,
# the order must be T->L->F

setwd("E:\\R_study\\ClinicalCSR-test\\abreviatedCSR")

# paths, files, ...
if(TRUE){
  rm(list=ls())
  who.maintain<-"stata93 <stata93@163.com>"

  #Rscirpt for CSR mockup shells generation
  dir(projpath<-getwd())

  dir(progpath<-projpath)
  progfile<-'CSR_mock_shell.R'

  loafile <- mockfile<- mockfileF<- "CSR_tfl_LoA.xlsx"

  endC <- "&&&&End&&&&"
  noteC<- "note: "
  popc1<- "[[Population]]"
  popc2<- "mmmmmm"

  outpath<-'shells_toc'
  outfile4 <- gsub('.xlsx', '.rtf', gsub('LoA', 'shell', loafile))
  oneTFLfile <- TRUE #whether generate an integrated TLF shell file or not
  max_row<-40 #maximum number of rows
}

#only use the run this file in simple way
if(F){
  dev.off()
  dir()
  source(progfile)
}

#libraries,functions, ...
if(TRUE){
  library(openxlsx)
  library(rtf)
  library(png)
  library(grid)

  ###########Function: convert a string into a vector of numbers#########
  s2n <- function(x, sp=','){
     unique(as.numeric(unlist(strsplit(x, split=sp, fixed=TRUE))))
  }

  ###########Function: increase TLF id by 1#########
  add_tid <- function(tlf.id){
    ed.id <- strsplit(tlf.id, split='.', fixed=TRUE)[[1]]
    len.id <- length(ed.id)
    paste(paste(ed.id[-len.id], collapse='.'),
           as.numeric(ed.id[len.id])+1, sep='.')
  }

  ###########Function: remove beginning space in strings#########
  noSpace0 <- function(xx=c("no ", "  mean", "     tes t ")){
    x2 <- xx
    for(i in 1:length(x2)){
      nosp<-T
      while(nosp){
        idx <- regexpr(" ", x2[i], fixed=TRUE)[1]
        if(idx==1){x2[i]<-substring(x2[i],2)}else{nosp<-FALSE}
      }
    }
    x2
  }

  ###########Function: generate a raster figure object#########
  pFig0 <- function(temID=NULL, wb=wb.fig){
    if(is.null(temID)){return(NULL)}
    s0<-substring(wb@.xData$sheet_names,1,1)
    figure_sheet <- wb@.xData$sheet_names[toupper(s0)=='F']
    whf <- which(figure_sheet==temID)
    img<- png::readPNG(sort(wb@.xData$media)[whf])
    return(dim(img)[1:2])
  }

  pFig <- function(temID='f0013', wb=wb.fig){
    s0<-substring(wb@.xData$sheet_names,1,1)
    figure_sheet <- wb@.xData$sheet_names[toupper(s0)=='F']
    whf <- which(figure_sheet==temID)
    img<- png::readPNG(sort(wb@.xData$media)[whf])
    img2<- grid::grid.raster(img,interpolate = TRUE)
    return(img2)
  }

  ###########Function: split table into multiple pages#########
  addT <-function(rtf.obj, tb,
                  cwid=NULL, cjust=NULL, #column width and justify
                  pwid=wdth, phei=hght, #page width and height
                  fsize=8, rownm=FALSE, na.s=' ',
                  maxRow=20,
                  adjByCol=NULL #whether adjusting maxRow by the ?th column
                  ){
    n.tb <- nrow(tb)
    numPg <- ceiling(n.tb/maxRow)
    nxtstt <- 1
    for(i in 1:numPg){
      if(i>1){
        addPageBreak(rtf, width=pwid, height=phei, omi=c(0.5,0.5,0.5,0.5))
      }
      if(!is.null(adjByCol) & numPg>1){
        sttrow <- nxtstt
        endrow<- maxRow-1 + sttrow
        if(endrow>=n.tb-1){ endrow<-n.tb }
        while(endrow<n.tb && endrow>0 &&
              all(tb[endrow+1,adjByCol]=='') ){endrow <- endrow-1}
        nxtstt <- endrow + 1
      }else{
        sttrow <- (i-1)*maxRow +1
        endrow <- min(sttrow+maxRow-1, n.tb)
      }
      addTable(rtf.obj, tb[sttrow:endrow,],
               col.width=cwid, col.justify=cjust,
               font.size=fsize, row.names=rownm, NA.string=na.s)
      if(i==numPg & endrow<n.tb-1){
        while(endrow<n.tb){
          sttrow <- nxtstt
          endrow<- min(maxRow-1 + sttrow, n.tb)
          if(endrow==n.tb-1){ endrow<-n.tb }
          while(endrow<n.tb && endrow>0 &&
                all(tb[endrow+1, adjByCol]=='')){endrow <- endrow-1}
          addPageBreak(rtf, width=pwid, height=phei, omi=c(0.5,0.5,0.5,0.5))
          addTable(rtf.obj, tb[sttrow:endrow,],
                   col.width=cwid, col.justify=cjust,
                   font.size=fsize, row.names=rownm, NA.string=na.s)
          nxtstt <- endrow + 1
        }
      }
    }
  } #~~~end addT
}

#import datasets
if(TRUE){
  loa <- openxlsx::read.xlsx(file.path(progpath, loafile), sheet='loa')
  wb.fig <- openxlsx::loadWorkbook(file.path(progpath, mockfileF))
}

#process and output mockup shells for tables
if(TRUE){
  TT<-LL<-FF<-0

  loa$tlfnumber <- loa$SECTION1 #for TLF ids.
  loa <- loa[!is.na(loa$ID),]
  loa[is.na(loa)] <- ''

  wdth<-11.2; hght<-8.5; ahh<-FALSE;

  #output table, listing and figure into one mockup shells
  rtf <- RTF(file.path(outpath, outfile4),
             width=wdth, height=hght, font.size=10, omi=c(0.5,0.5,0.5,0.5))

  addText(rtf, "Table of Contents\n", bold= T)
  addTOC(rtf)

  #for(i in 1:nrow(loa)){
  for(i in 1:nrow(loa)){
    tt <- loa$TITLE1[i] #get title
    tt.vb <- substring(tt, gregexpr('<<', tt)[[1]][1]+2, gregexpr('>>', tt)[[1]][1]-1)
    pp0<- loa$TITLE3[i] #get basic population
    ft0<- loa$Notes[i]  #get additional footnote
    id0<- loa$ID[i]     #get TLF template id
    tid0<-tid<- loa$tlfnumber[i]   #start TLF id
    print( paste(i,': ', id0,'is completed'))  #show log

    if(!ahh){ #avoid adding pagebreak after a header
      addPageBreak(rtf, width=wdth, height=hght, omi=c(0.5,0.5,0.5,0.5))
    }
    if(id0=='head'){
      addHeader(rtf, title=tt, TOC.level=1)
      ahh<-TRUE
      next
    }else{ahh<-FALSE}

    # read in the TLF template
    tf0 <- openxlsx::read.xlsx(file.path(progpath, loafile), sheet=id0)
    colnames(tf0)<-paste0('X', 1:ncol(tf0))
    tf0[is.na(tf0)]<-''
    nt0wh <- grep('note:', tf0$X1, fixed=TRUE)
    ed0wh <- grep('&&End&', tf0$X1, fixed=TRUE)
    if(length(ed0wh)==0){stop(paste('You forget to inser &&&&End&&&& to', id0))}
    tf <- tf0[1:(nt0wh-1),]      #main table
    #remove begining spaces in strings
    colnames(tf) <- noSpace0(tf0[1,])
    tf <- as.matrix(apply(tf, 2, noSpace0))
    cctf<-unlist(sapply(strsplit(paste0(colnames(tf),"\n"), split='\n'),
                        function(x){x[nchar(x)==max(nchar(x))][1]}))
    bbtf<-apply(tf, 2, function(y){
      unlist(sapply(strsplit(paste0(y,'\n'), split='\n'),
                    function(x){x[nchar(x)==max(nchar(x))][1]}))
    })
    colw <- apply(rbind(cctf,bbtf), 2, function(x){max(nchar(x))})
    colw<-0.85*wdth*colw/sum(colw)
    colw[colw<0.6] <- 0.6
    coladj <- c('L', rep('L',length(colw)-1))

    ftnt <- tf0[nt0wh:(ed0wh),2] #footnote
    ftnt <- gsub('<xxxxxx>', tt.vb, ftnt)
    ftnt.pop<-''

    tt2 <- gsub("<<|>>", "", tt)
    tt2 <- gsub("in [[Population]]", '', tt2, fixed=TRUE) #pp0
    tt2 <- gsub("In [[Population]]", '', tt2, fixed=TRUE)
    tt2 <- gsub(" [[Population]]", '', tt2, fixed=TRUE)
    tt2 <- gsub("in [[pop]]|In [[pop]]", '', tt2, fixed=TRUE)
    if(substring(id0,1,1)!='G')
      tt2 <- paste0(tid,'\n', tt2, '\n', pp0,"\\qc")
    pp02<-paste0(pp0, '\\qc')
    ftnt.pop.k<-grep("mmmmmm", ftnt)[1]
    ftnt[ftnt.pop.k] <- paste(ftnt[ftnt.pop.k], ftnt.pop)
    ftnt<- gsub("mmmmmm", pp0, ftnt, fixed=TRUE)
    tt2 <- gsub('  ', '', tt2, fixed=TRUE)
    ftnt <- gsub('  ', '',ftnt, fixed=TRUE)
    tid0 <- gsub('  ', '', tid0, fixed=TRUE)
    ftnt.pop <- gsub('  ', '', ftnt.pop, fixed=TRUE)

    #convert table into data.frame
    if(nrow(tf)==2 ){
      tf.cnm <- colnames(tf)
      tf_1<-matrix(tf[-1,], nrow=1)
      colnames(tf_1) <- tf.cnm
    }else{tf_1 <- tf[-1,]}

    if(substring(id0,1,1)=='G'){
      #remove double spaces
      setFontSize(rtf, 9)
      addHeader(rtf, title=tt2, TOC.level=2, font.size=10)
      #add table
      addT(rtf, tf_1, cwid=colw, maxRow=max_row, adjByCol=1)
    }

    if(substring(id0,1,1)=='T'){
      #remove double spaces
      setFontSize(rtf, 9)
      addHeader(rtf, title=tt2, #subtitle=pp02,
                TOC.level=2, font.size=10)
      #add table
      addT(rtf, tf_1, cwid=colw, maxRow=max_row, adjByCol=1)
    }

    if(substring(id0,1,1)=='L'){
      setFontSize(rtf, 9)
      addHeader(rtf, title=tt2, #subtitle=pp02,
                TOC.level=2, font.size=10)
      #add table
      addT(rtf, tf_1, cwid=colw, cjust=coladj, maxRow=max_row, adjByCol=1)
    }

    if(substring(id0,1,1)=='F'){
      setFontSize(rtf, 9)
      addHeader(rtf, title=tt2, #subtitle=pp02,
                TOC.level=2, font.size=10)
      plot.fig <- function(){ pFig(id0) }
      increaseIndent(rtf, 1.2)
      picS <- pFig0(id0)/2000
      picW <- min(0.8, max(picS[2], 0.4))*wdth
      picH <- min(picW*picS[1]/picS[2], 0.55*hght)
      picW <- min(picW, picH*picS[2]/picS[1])
      addPlot(rtf, plot.fun=plot.fig, width=picW, height=picH, res=200)
      decreaseIndent(rtf, 1.2)
    }

    #add footnote and notes to programmers
    addNewLine(rtf)
    addText(rtf, paste(ftnt, collapse='\n'))
    if(nchar(gsub(' ', '', ft0, fixed=TRUE))>0) {
      addNewLine(rtf); addText(rtf, ft0)
    }
  }

  done(rtf)
}

################################################
print( paste('The TLF shells are generated in the rtf file.'))



