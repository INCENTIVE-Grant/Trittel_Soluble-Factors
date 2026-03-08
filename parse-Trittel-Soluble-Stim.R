#!/usr/bin/env Rscript
##
## Script to parse the four sheets in Stephanie Trittel's Excel
## spreadsheet into a LONG-FORMAT CSV file. There is a large amount of
## data here, with over 1700 rows x AK columns of MFI sample data
## alone. That plus the same data converted to concentration (pg/ml)
## and 66 rows of bridging samples (MFI & conc) means there is a lot
## to output. However, the data appear very well organized and my eye
## has not detected inconsistancies in the naming, etc., yet.
##
## Note: at this time, the bridging samples are of limited utility for
## correction because though they carefully document the plate on
## which they were analyzed, the plate on which samples were run is
## not documented. I shall attempt to get Stephie to correct this
## oversight.
##
## VERSION HISTORY
## [2025-06-15 MeD] Initial version.
##
##********************************************************************************
library(AnalysisHeader)
library(tibble)
library(readxl)

## Bring in many variables that are used across all parsing:
##    TrialNames, AssayNames, SubAssayNames, SubAssays,
##    SampleTypes, VisitDay (D000 - D365), CSV ColumnNames,
##    KnownStrains, AliasStrains, ExcelColumns.
if( !file.exists('Controlled-Vocab.R') ) {
    cat("Expected to find the file 'Controlled-Vocab.R' in the local directory.\n")
    cat("This is usually accomplished by a soft-link (Linux/Unix/Mac) to the file.\n")
    stop("Missing file: Controlled-Vocab.R")
}
source('Controlled-Vocab.R')

## GLOBAL variables
ProgramName <- 'parse-Trittel-Soluble-Stim.R'
Version <- 'v1.0'

options(warn=1, width=132)

## Enable some DEBUGGING statements if TRUE
DEBUG <- TRUE

## Annotate that DEBUGGING is turned on
if(DEBUG == TRUE) {
    cat("\nDEBUGGING is Enabled.\n")
}

## Expected Excel worksheet names
ExpectedSheets <- c("Net MFl", "Bridging Samples-Net MFI",
                    "Concentration", "Bridging Samples-Concentration")

## "Pretty" lines for dividing the output - double-header or single-header
dhLine <- paste(rep('=', length=(getOption('width')-2)), collapse='')
shLine <- paste(rep('-', length=(getOption('width')-2)), collapse='')

## Collect the table of stims (from data dictionary) here for checking
Stims <- data.frame(
    Trial=c(
        rep('QIV2', times=7),
        rep('QIV3', times=9)),

    Name=c(
        ## Used for QIV2
        'NEG', 'MIX', 'VIC', 'TAS', 'WAS', 'PUK', 'SEB',
        ## Used for QIV3
        'WBSR', 'WBSM', 'WBSV', 'WBST', 'WBSW', 'WBSP', 'WBSS', 'WBSD', 'WBSA'),

    Description=c(
        ## Used for QIV2
        'Unstimulated (negative)',
        'Antigen Mix (a mix of the four vaccine antigens)',
        'Influenza antigen A/Victoria/2570/2019',
        'Influenza antigen A/Tasmania/503/2020',
        'Influenza antigen B/Washington/02/2019 ',
        'Influenza antigen B/Phuket/3073/2013',
        'Maximal stimulation  (Staphylococcal Enterotoxin B)',
        ## Used for QIV3
        'Unstimulated (negative)',
        'Antigen Mix (a mix of the four vaccine antigens)',
        'Influenza antigen A/Victoria/2570/2019',
        'Influenza antigen A/Tasmania/503/2020',
        'Influenza antigen B/Washington/02/2019 ',
        'Influenza antigen B/Phuket/3073/2013',
        'Maximal stimulation  (Staphylococcal Enterotoxin B)',
        'Influenza antigen A/Darwin/9/2021',
        'Influenza antigen B/Austria/1359417/2021'),

    Strain=c(
        "None", "All",
        "A/Victoria/2570/2019 (H1N1)", "A/Tasmania/503/2020 (H3N2)",
        "B/Washington/2/2019", "B/Phuket/3073/2013",
        "Staph. Enterotoxin B",

        "None", "All",
        "A/Victoria/2570/2019 (H1N1)", "A/Tasmania/503/2020 (H3N2)",
        "B/Washington/2/2019", "B/Phuket/3073/2013",
        "Staph. Enterotoxin B",
        "A/Darwin/9/2021 (H3N2)","B/Austria/1359417/2021")
)

stopifnot(Stims$Trial %in% TrialNames,
          Stims$Strain %in% c(KnownStrains, "None", "All", "Staph. Enterotoxin B") )

##********************************************************************************
## Begin main code - process command line arguments.

## Assign my input file name; if I don't, then get the command-line argument
if(interactive())
    inFile <- "Guzmán-Riese-Tritel_Soluble-Factors_2025-06-12.xlsx"

## Check if an 'inFile' object already exists. Useful in debugging, etc.
if( !exists('inFile') ) {
    inFile <- commandArgs(trailingOnly=TRUE)[1]
}

## Check command-line args are ok else display usage message
if(is.na(inFile) || file.exists(inFile) == FALSE) {
    cat(paste0(ProgramName, " - Rscript to convert Stephie Trittel's XLSX results to LONG-format CSV"),
        paste("Version:", Version),
        " ",
        "Usage:",
        paste0("\t", ProgramName, " <input-XLSX-file>"),
        "\nWhere:",
        "\t<input-XLSX-file> = An Excel file (*.xlsx) in Stephie Trittel's format for her assays",
        "\nOutputs",
        "\tOutput CSV file: same name as input with '.xlsx' replaced with '.csv'. No overwriting.",
        "\tLog text file: parse-Trittel-Soluble-Stim_YYYYMMDD.log",
        " ",
        sep='\n')
    stop("Invalid input filename")
}

## Prepare an output file. Base the file name on the input name and on date.
StartTime <- Sys.time()   # Used for both file names and computing run-time.
Today <- format(Sys.time(), "_%Y%m%d")
outName <- paste0(gsub('\\.xlsx$', '', inFile), Today, '.csv')

## Prepare a Log File for logging data via STDOUT & STDERR
logName <- paste0(gsub('\\.R$', '', ProgramName), Today, '.log')
if( !interactive() ) {
    cat("\n*** Redirecting program reporting to Log File:", logName, "\n")
    LogFile <- file(logName, open='wt')
    sink(LogFile)
    sink(LogFile, type='message')

    ## Annotate that DEBUGGING is turned on (here, into the log file)
    if(DEBUG == TRUE)
        cat("\nDEBUGGING is Enabled.\n")
}

## Document the run-time information
print( collectRunInfo(ProgramName, Version) )
cat('Vocabulary Version:', VocabVersion, "\n\n")
cat("Data input & output files:\n",
    "\tInp = ", inFile, "\n",
    "\tOut = ", outName, "\n",
    "\tLog = ", logName, "\n",
    "\n",
    sep='')

##********************************************************************************
##                                    SUBROUTINES
##********************************************************************************
## Layout of sheet called "Net MFI"
##   Row 1 = Analyte measured
##   Row 2 = Some header + Unit information
##   Row 3 - N = data
##
##   Column 1 = Complex SubjectID + Visit + Stim
##   Column 2 = SubjectID or BLANK
##   Column 3 = Beginning of measurements
##' Parse the Sample MFI data into Long Format
##'
##' All of the Sample MFI data are on one worksheet called "Net MFI".
##' Column 1 contains a complex name built of SubjectID + Visit Number + Stimulator Factor
##' Column 2 contains the SubjectID but with many BLANKs (pretty formatting)
##' Column 3-N contains the measurements
parseSampleMFI <- function(File, Sheet) {
    ## FIXME: The MFI and Concentration parsing is so simlar that I should combine them
    ## FIXME: Adjust the columns to avoid the "Stimulation" column.
    cat("\n", dhLine, "\nParsing Sample MFI worksheet.\n", sep='')

    cat("\n\tReading row-1: Analyte names.\n")
    analytes <- unname(unlist(read_xlsx(path=File, sheet=Sheet,
                                        col_names=FALSE, trim_ws=TRUE, n_max=1, skip=0)))
    cat("\n\tReading row-2: Headings and units.\n")
    unitsRow <- unname(unlist(read_xlsx(path=File, sheet=Sheet,
                                        col_names=FALSE, trim_ws=TRUE, n_max=1, skip=1)))
    cat("\n\tReading remaining rows: MFI data.\n")
    dat      <- as.data.frame(read_xlsx(path=File, sheet=Sheet,
                                        col_names=FALSE, trim_ws=TRUE, skip=2))
    cat("\n\tParsing SampleNames --> SubjectID, Visit Number, Stim Treatment\n")
    samples <- parseSampleNames(dat[, 1], dat[, 2])

    ##------------------------------------------------------------
    ## Log the experimental design for MFI
    cat("\n", shLine, "\nEXPERIMENTAL DESIGN for MFI data set:\n\n", sep='')
    cat("For QIV2:\n")
    inx <- samples$Trial == 'QIV2'
    print(table(samples$SubjectID[inx], samples$Stim[inx], samples$Visit[inx]))

    cat("\nFor QIV3:\n")
    inx <- samples$Trial == 'QIV3'
    ## Note in the design, the switching between (WBSA, WBSD) <--> (WBST, WBSW)
    ## as expected for the different vaccines delivered.
    print(table(samples$SubjectID[inx], samples$Stim[inx], samples$Visit[inx]))

    ##------------------------------------------------------------
    ## Build the LONG FORMAT data frame, 'res' for 'result', to return.
    ## Header for Dobaño data set is:
    ##    "SampleType","Trial","SubjectID","Day","Assay","Strain","Protein",
    ##    "StrainProt","Dilution","Value","ValueUnit","Isotype","UreaPresent","PlateID","Well"
    numAnalytes <- length(analytes)
    stopifnot(numAnalytes == ncol(dat)-2)
    numSamples <- nrow(dat)
    N <- numAnalytes * numSamples ## Wide --> Long format
    res <- data.frame(SampleType=rep("Samp", times=N),
                      Trial=rep(samples$Trial, times=numAnalytes),
                      SubjectID=rep(samples$SubjectID, times=numAnalytes),
                      Day=rep(samples$Day, times=numAnalytes),
                      Assay=rep('SolFact', times=N),
                      Strain=rep(NA_character_, times=N),  # Fill-in later from Stims
                      Protein=rep(NA_character_, times=N),
                      StrainProt=rep(NA_character_, times=N),
                      Dilution=rep(NA_real_, times=N),
                      Value=rep(NA_real_, times=N),
                      ValueUnit=rep('MFI', times=N),
                      Stimulation=rep(samples$Stim, times=numAnalytes)
                      )
    ## Report on size of the data frame extracted
    if( DEBUG ) {
        cat("Data frame 'res' is ", nrow(res), " rows x ", ncol(res), " columns.\n",
            "\tTotal storage is ", object.size(res), " bytes.\n", sep='')
        cat("\nAnalyte storage rows in 'res' data frame:\n")
        cat(sprintf("\t%20s %5s %5s\n", "Analyte", "Low", "High"))
    }

    ## Load in the analytes values into "Assay" as "SolFac: <analyte>"
    ##    Assay := "SolFac: <analyte>" so each analyte is a separate assay
    ##    Protein := Stimulation, as is
    ##    Strain  := Strain associated with the Stim
    for(i in 1:numAnalytes) {
        ## Build an index to access a subset for 'res' to store an analyte set of values
        lo <- ( (i - 1) * numSamples ) + 1
        hi <- i * numSamples
        inx <- rep(FALSE, times=N)
        inx[lo:hi] <- TRUE
        if( DEBUG )
            cat(sprintf("\t%20s %5d %5d\n", analytes[i], lo, hi))
        stopifnot(sum(inx) == nrow(samples))

        ## Lookup the items associated with the stim: Strain, etc
        ind <- match(samples$Stim, Stims$Name)
        stopifnot( !is.na(ind) )

        ## Load the results into the 'res' data frame
        res$Strain[inx] <- Stims$Strain[ind]
        res$Protein[inx] <- analytes[i]
        res$Dilution[inx] <- 1.0
        res$Value[inx] <- dat[, i+2]
    }

    ##------------------------------------------------------------
    ## Return the results
    return(res)
}

##------------------------------------------------------------
## Parse the Sample "Concentration" worksheet.
## At first glance, it appears identical to the MFI worksheet.
## Begin then with copy-and-paste code.
parseSampleConc <- function(File, Sheet) {
    cat("\n", dhLine, "\nParsing Sample 'Concentration' worksheet.\n", sep='')

    cat("\n\tReading row-1: Analyte names.\n")
    analytes <- unname(unlist(read_xlsx(path=File, sheet=Sheet,
                                        col_names=FALSE, trim_ws=TRUE, n_max=1, skip=0)))
    cat("\n\tReading row-2: Headings and units.\n")
    unitsRow <- unname(unlist(read_xlsx(path=File, sheet=Sheet,
                                        col_names=FALSE, trim_ws=TRUE, n_max=1, skip=1)))
    cat("\n\tReading remaining rows: Concentration data.\n")
    dat      <- as.data.frame(read_xlsx(path=File, sheet=Sheet,
                                        col_names=FALSE, trim_ws=TRUE, skip=2, na=c("", "OOR>")))
    cat("\n\tParsing SampleNames --> SubjectID, Visit Number, Stim Treatment\n")
    samples <- parseSampleNames(dat[, 1], dat[, 2])

    ##------------------------------------------------------------
    ## Log the experimental design for Concentration
    cat("\n", shLine, "\nEXPERIMENTAL DESIGN for Concentration data set:\n\n", sep='')
    cat("For QIV2:\n")
    inx <- samples$Trial == 'QIV2'
    print(table(samples$SubjectID[inx], samples$Stim[inx], samples$Visit[inx]))

    cat("\nFor QIV3:\n")
    inx <- samples$Trial == 'QIV3'
    ## Note in the design, the switching between (WBSA, WBSD) <--> (WBST, WBSW)
    ## as expected for the different vaccines delivered.
    print(table(samples$SubjectID[inx], samples$Stim[inx], samples$Visit[inx]))

    ##------------------------------------------------------------
    ## Build the LONG FORMAT data frame, 'res' for 'result', to return.
    ## Header for Dobaño data set is:
    ##    "SampleType","Trial","SubjectID","Day","Assay","Strain","Protein",
    ##    "StrainProt","Dilution","Value","ValueUnit","Isotype","UreaPresent","PlateID","Well"
    numAnalytes <- length(analytes)
    stopifnot(numAnalytes == ncol(dat)-2)
    numSamples <- nrow(dat)
    N <- numAnalytes * numSamples ## Wide --> Long format
    res <- data.frame(SampleType=rep("Samp", times=N),
                      Trial=rep(samples$Trial, times=numAnalytes),
                      SubjectID=rep(samples$SubjectID, times=numAnalytes),
                      Day=rep(samples$Day, times=numAnalytes),
                      Assay=rep('SolFact', times=N),
                      Strain=rep(NA_character_, times=N),  # Fill-in later from Stims
                      Protein=rep(NA_character_, times=N),
                      StrainProt=rep(NA_character_, times=N),
                      Dilution=rep(NA_real_, times=N),
                      Value=rep(NA_real_, times=N),
                      ValueUnit=rep('pg/ml', times=N),
                      Stimulation=rep(samples$Stim, times=numAnalytes)
                      )
    ## Load in the analytes values into "Assay" as "SolFac: <analyte>"
    if( DEBUG ) {
        cat("Data frame 'res' is ", nrow(res), " rows x ", ncol(res), " columns.\n",
            "\tTotal storage is ", object.size(res), " bytes.\n", sep='')
        cat("\nAnalyte storage rows in 'res' data frame:\n")
        cat(sprintf("\t%20s %5s %5s\n", "Analyte", "Low", "High"))
    }
    for(i in 1:numAnalytes) {
        ## Build an index to access a subset for 'res' to store an analyte set of values
        lo <- ( (i - 1) * numSamples ) + 1
        hi <- i * numSamples
        inx <- rep(FALSE, times=N)
        inx[lo:hi] <- TRUE
        if( DEBUG )
            cat(sprintf("\t%20s %5d %5d\n", analytes[i], lo, hi))
        stopifnot(sum(inx) == nrow(samples))

        ## Lookup the items associated with the stim: Strain, etc
        ind <- match(samples$Stim, Stims$Name)
        stopifnot( !is.na(ind) )

        ## Load the results into the 'res' data frame
        res$Strain[inx] <- Stims$Strain[ind]
        res$Protein[inx] <- analytes[i]
        res$Dilution[inx] <- 1.0
        res$Value[inx] <- dat[, i+2]
    }

    ##------------------------------------------------------------
    ## Return the results
    return(res)
}

##--------------------------------------------------------------------------------
parseSampleNames <- function(sampleNames, subjectID) {
    stopifnot(is.character(sampleNames))

    ## Fill-down 'subjectID'. We'll need it in the FIXES (#2), below.
    beg <- which( !is.na(subjectID) )
    end <- beg - 1
    end <- c(end[-1], length(subjectID))
    for(i in 1:length(beg))
        subjectID[ (beg[i] + 1):end[i] ] <- subjectID[ beg[i] ]

    ## Apply FIXES. These are based on "consistency" alone and need Scientist confirmation.
    adjustedValue <- rep(FALSE, length(sampleNames))

    ## FIX #1:
    dupName <- 'UIB005-V2-SEB'
    replName <- 'UIB005-V1-SEB'
    inx <- sampleNames == dupName
    if(sum(inx) > 1) {
        cat("\nSample name '", dupName, "' is duplicated in rows: ",
            paste(which(inx), collapse=', '), ".\n", sep='')
        cat("Based on surrounding sample names, the first of the two will be adjusted to '",
            replName, "'.\n", sep='')
        sampleNames[ which(inx)[1] ] <- replName
        adjustedValue[ which(inx)[1] ] <- TRUE
    }

    ## FIX #2: On "Concentration" worksheet, Subject CHU002 has not initial SubjectID in SampleName
    subID   <- 'CHU002'
    badName <- 'V1WBSM1'
    newName <- 'CHU002-V1-WBSM1'
    inx <- (subjectID == subID) & (sampleNames == badName)
    if(sum(inx) > 0) {
        cat("\nSample name '", badName, "' is not formatted correctly in rows: ",
            paste(which(inx), collapse=', '), ".\n", sep='')
        cat("Based on surrounding sample names, it will be adjusted to '",
            newName, "'.\n", sep='')
        sampleNames[inx] <- newName
        adjustedValue[inx] <- TRUE
    }

    ## FIX #3: Several names are simple replacements
    fixData <- data.frame(BadName=c( "CHU001-1V4WBSW1", "UIBß019-V1-PUK", "1V4WBSW1", "UIb039-V1-PUK"),
                          NewName=c( "CHU001-V4WBSW1",  "UIB019-V1-PUK",  "V4WBSW1",  "UIB039-V1-PUK")
                          )
    for(i in 1:nrow(fixData)) {
        inx <- sampleNames == fixData$BadName[i]
        if(sum(inx) == 1) {
            cat("\nSample name '", fixData$BadName[i], "' appears incorrectly formatted in row: ",
                paste(which(inx), collapse=', '), ".\n", sep='')
            cat("Based on surrounding sample names, the name will be adjusted to: '",
                fixData$NewName[i], "'.\n", sep='')
            sampleNames[inx] <- fixData$NewName[i]
            adjustedValue[inx] <- TRUE
        }
    }

    ## Sample names are built with HYPHENs: SubjectID - Visit Number - Stim Treatment
    ## Unfortunately, they are inconsistant:
    ##   1) Sometimes the HYPHEN between Visit and Treatment is dropped.
    ## Tried "strsplit()" without a clean solution. Try RegEx.
    ## Note: on the 'Concentration' worksheet, a further complication begins where
    ##       the SubjectID is dropped for days beyond "V1". Instead, the name is
    ##       shorter and begins with a 'V<number>'.
    inxShort <- grepl('^V[0-9]', sampleNames)
    inxLong  <- !inxShort

    ## Pre-declare returned values to allow indexing into them
    id <- visit <- stim <- sampNum <- rep(NA_character_, length(sampleNames))

    ## Extract out the values from the long names: SubjectID - Visit Number - Stim Treatment
    id[inxLong]      <- gsub('^([a-zA-Z]{3}[0-9]{3}).*$', '\\1',
                             sampleNames[inxLong], perl=TRUE)
    visit[inxLong]   <- gsub('^[a-zA-Z]{3}[0-9]{3}[- ]*(V[0-9]).*$', '\\1',
                             sampleNames[inxLong], perl=TRUE)
    stim[inxLong]    <- gsub('^[a-zA-Z]{3}[0-9]{3}[- ]*V[0-9]-?([A-Z]{3,4})[0-9]?$', '\\1',
                             sampleNames[inxLong], perl=TRUE)
    sampNum[inxLong] <- gsub('^[a-zA-Z]{3}[0-9]{3}[- ]*V[0-9]-?[A-Z]{3,4}([0-9]?)$', '\\1',
                              sampleNames[inxLong], perl=TRUE)

    ## Extract out the values from short names (no SubjectID): Visit Number - Stim Treatment
    if( sum(inxShort) > 0 ) {
        visit[inxShort]   <- gsub('^(V[0-9])-?[A-Z]{3,4}[0-9]?$', '\\1',
                                  sampleNames[inxShort], perl=TRUE)
        stim[inxShort]    <- gsub('^V[0-9]-?([A-Z]{3,4})[0-9]?$', '\\1',
                                  sampleNames[inxShort], perl=TRUE)
        sampNum[inxShort] <- gsub('^V[0-9]-?[A-Z]{3,4}([0-9]?)$', '\\1',
                                  sampleNames[inxShort], perl=TRUE)

    }

    ## This leaves many "holes" in the 'id' for the 'Concentration' worksheet.
    ## Try pulling from subjectID, which has been "filled-down".
    inx <- !is.na(id) & !is.na(subjectID)
    stopifnot( id[inx] == subjectID[inx] )   # Seek to confirm alignment of 'id' and 'subjectID'.
    id <- subjectID

    ## FIX #4: Some errors in Stim and Visit day. Look for duplicates and patterns of stim.
    subjectIDs <- sort(unique(id))
    visits <- sort(unique(visit))
    keys <- c(t(outer(subjectIDs, visits, FUN=paste, sep='-')))
    stimPattern <- matrix(NA_integer_, nrow=length(keys), ncol=nrow(Stims),
                          dimnames=list(Key=keys, Stim=Stims$Name))
    ## Collect the SubjectID x Visit types of patterns.
    ## We're searching for the "typical" pattern in which the Stims are
    ## applied - NEG then SEB then ...
    ## I hope to see a pair (QIV2 and QIV3) of patterns predominate with
    ## a few oddballs, indicating possible errors, thrown in.
    for(i in seq_along(subjectIDs)) {
        for(j in seq_along(visits)) {
            key <- paste(subjectIDs[i], visits[j], sep='-')
            inxKey <- (subjectIDs[i] == id) & (visits[j] == visit)
            stimNames <- stim[inxKey]
            stimPattern[ (key == rownames(stimPattern)) , ] <- match(colnames(stimPattern), stimNames)
        }
    }

    ## Ok - lots of stim patterns collected. What is common and what is uncommon?
    ## Build a set of unique keys to discern
    stimPatternKey <- apply(stimPattern, 1, function(x) paste(x, collapse='_'))
    cat("\nPattern of Stimulus applications computed. Top lines are:\n")
    print(head(stimPattern, 12))
    cat("\nA count of the observed patterns is:\n\tkey= ",
        paste(colnames(stimPattern), collapse='_'), "\n", sep='')
    tb <- table(stimPatternKey)
    print(sort(tb))

    ## Select the odd patterns based on arbitrary cut-off of '20'
    cat("\nOdd patterns are:\n")
    print(tb[ tb < 20 ] )
    oddPatterns <- names(tb[tb<20])
    oddNames <- rownames(stimPattern)[ stimPatternKey %in% oddPatterns ]
    cat("Odd patterns correspond to the following samples:\n\t",
        paste(oddNames, collapse=', '), "\n", sep='')
    cat("\nContext for each of the odd names is shown below:\n")
    inxOrd <- order(id, visit, stim)
    for(nm in oddNames) {
        cat("\n\tName =", nm, "\n")
        nmSplit <- strsplit(nm, '-')[[1]]
        inxSub <- (id[inxOrd] == nmSplit[1])
        tmp <- data.frame(SubjectID=id[inxOrd][inxSub],
                          Visit=visit[inxOrd][inxSub],
                          Stim=stim[inxOrd][inxSub]
                          )
        print(tmp)
    }

    ## Re-derive "Day" from "Visit"
    day <- ifelse(visit == 'V1', "D000",
           ifelse(visit == 'V2', "D003-8",
           ifelse(visit == 'V3', "D030",
           ifelse(visit == 'V4', "D058",
           ifelse(visit == 'V5', "D180",
           ifelse(visit == 'V6', "D365", NA))))))

    ## Guess at clinical trial based in SubjectID
    trial <- ifelse(grepl('^UIB', id), 'QIV2', 'QIV3')

    sampNames <- data.frame(SampleNames=sampleNames,
                            Adjusted=adjustedValue,
                            Trial=trial,
                            SubjectID=id,
                            Visit=visit,
                            Day=day,
                            Stim=stim,
                            SampleNumber=sampNum,
                            ParticipantID=subjectID    # Since I "fill-down" earlier, let's keep it
                            )
    return(sampNames)
}


##********************************************************************************
##                                    MAIN ROUTINE
##********************************************************************************
## Figure out which sheets are present
sheetNames <- excel_sheets(inFile)
stopifnot(sheetNames %in% ExpectedSheets)  # Catch major errors

## Get the "raw data" in MFI - then confirm correct formats
rawData  <- parseSampleMFI(File=inFile, Sheet="Net MFl")
stopifnot(rawData$SampleType %in% SampleTypes,
          rawData$Trial %in% TrialNames,
          rawData$Day %in% VisitDay,
          rawData$Assay %in% AssayNames,
          rawData$Strain %in% KnownStrains,
          colnames(rawData) %in% ColumnNames
          )


concData <- parseSampleConc(File=inFile, Sheet="Concentration")

##--------------------------------------------------------------------------------
## Why are rawData and concData  different sizes?
if(DEBUG) {

    ##keyRaw <- unique(rawData[, c('SubjectID', 'Day', 'Strain', 'Protein')])
    keyRaw <- paste(rawData$SubjectID, rawData$Day, rawData$Strain, rawData$Protein, sep='@')
    cat("Rows x Columns of MFI data:\n")
    print(dim(rawData))
    #print(length(keyRaw))

    keyConc <- paste(concData$SubjectID, concData$Day, concData$Strain, concData$Protein, sep='@')
    cat("Rows x Columns of Concentration data:\n")
    print(dim(concData))
    #print(length(keyConc))
}

##--------------------------------------------------------------------------------
## Output the data to a CSV file

## Create the one big data frame
tmp <- rbind(rawData, concData)
cat(dhLine,
    "\nCombined dataset of MFI and Concentration created.\n",
    "\tDataset is ", format(nrow(tmp), big.mark=','),
    " rows x ", format(ncol(tmp), big.mark=','), " columns.\n",
    "\tIt is ", format(unclass(object.size(tmp)), big.mark=','), " bytes in RAM.\n\n", sep='')

cat("Writing combined dataset to file:", outName, "\n\n")
write.csv(tmp, outName, row.names=FALSE)

##********************************************************************************
## Completed.
EndTime <- Sys.time()
cat(dhLine, "\nCompleted: ", format(EndTime, '%Y-%m-%d %H:%M:%S'), ".\n",
    "Run-time: ", difftime(EndTime, StartTime, units='secs'), " secs.\n", sep='')

if( !interactive() ) {
    sink(type='message')
    sink()
}

## Show 'completed' so we know there was no error.
cat("\nCompleted.\n")
