## TODO
## [ ] Add barcode if available
## [ ] add kinetic_timestamp badsed on start-time
## [ ] Add readerplatenumber (for stacked) based on start time
## [ ] Return headers to the environment as ingestr does (filename_sheetname)
## [ ] data type (absorption)
## XXX are emission files similar?

ingest_icontrol_xlsx <- function(ReaderFile) {
  ## 1. Check file is iControl xlsx
  ## 2. Parse header
  ## 3. Parse all sheets
  
  ## Get list of sheets in workbook: FileSheets
  if(class(try(FileSheets <- readxl::excel_sheets(path = ReaderFile)))=='try-error'){
    futile.logger::flog.error("Files %s is not an xlsx file", ReaderFile)
    stop()
  }
  
  futile.logger::flog.info("Found %s sheets in %s: %s", length(FileSheets), ReaderFile, paste0(FileSheets,collapse=', '), name = "platereader")
  
  ReaderFileDF <- plyr::ldply(FileSheets, function(x) icontrol_sheet(ReaderFile=ReaderFile, Sheet=x)$data)
  ##ReaderFileDF <- purrr::map_df(.x=FileSheets, .f=function(x) icontrol_sheet(ReaderFile=ReaderFile, Sheet=x)$data, .id=NULL)
  SheetDF <- tibble::as_tibble(list(icontrol_sheet_name=FileSheets, icontrol_sheet_number=seq_along(FileSheets)))
  ReaderFileDF <- checkRows(nrow(ReaderFileDF),merge(SheetDF, ReaderFileDF, by='icontrol_sheet_name'))
  tibble::as_tibble(plyr::arrange(ReaderFileDF, icontrol_sheet_number)) ## would like to sort on well_name, but it is well_name_384
}

icontrol_sheet <- function(ReaderFile, Sheet) {
  ## imports a single sheet
  ## example data: 
  futile.logger::flog.trace("icontro_sheet processing %s:%s", ReaderFile, Sheet)
  RawFile <- readxl::read_excel(ReaderFile, Sheet, col_names=FALSE, col_types="text")
  
  if(nrow(RawFile)==0) ## Empty sheet
    return(list(data=NULL, header=NULL))
  
  ## Check it is from i-control
  if(! grepl('i-?control', as.character(RawFile[1,1]), ignore.case=TRUE)) {
    futile.logger::flog.error("Sheet %s of File %s is not an i-control file", Sheet, ReaderFile)
    stop()
  }
  
  ## Find the data
  DataBlock <- which(grepl("^[A-P]\\d{0,2}$",as.character(RawFile[,1,drop=TRUE]))) ## XXX 1536 letters?
  RowNames <- as.character(RawFile[DataBlock,1,drop=TRUE])

  ## Find form of data: list (one row per well) or matrix (one row per row)
  DataTableType <- 'unknown'
  if(all(grepl('^[A-P]$', RowNames)))
    DataTableType <- 'matrix'
  if(all(grepl('^[A-P]\\d{1,2}$', RowNames)))
    DataTableType <- 'list'
  futile.logger::flog.info("Found %s rows of %s data in Sheet %s of File %s", length(RowNames),DataTableType, Sheet, ReaderFile)
  if(DataTableType=='unknown')
    stop("Unknown Dataformat")
  
  ## Check dataBlock is continuous
  if(any(diff(DataBlock)!=1)){
    futile.logger::flog.error("Datablock is not continuous at row(s) %s", paste0(DataBlock[which(diff(DataBlock)!=1)], collapse=' ,'))  
    stop()
  }
    
  ## Header is above data separated by empty line
  ## OBS: readxl 0.1.1 does not handle empty rows well: tidyverse/readxl
  if(packageVersion("readxl") < '1.0.0') {
    futile.logger::flog.error("platereader needs readxl 1.0.0 or newer to parse i-control files")
    stop()
  }
  EmptyLines <- which(apply(as.matrix(RawFile), 1, function(x) all(is.na(x))))
  HeaderBlock <- seq(1, max(EmptyLines[EmptyLines < min(DataBlock)]))
  HeaderData <- RawFile[setdiff(HeaderBlock,EmptyLines),]

  ## Parse HeaderData
  HeaderDF <- icontrol_header(HeaderData)
  
  ## Parse readerData
  if(any(grepl('Kinetic', HeaderDF["parameter"]))) {
##     ReaderDF <- icontrol_kinetic_data(RawFile[c(DataBlock[1]-c(3,2,1),DataBlock),])
    ReaderDF <- icontrol_kinetic_data(readxl::read_xlsx(path=ReaderFile, sheet=Sheet, skip=DataBlock[1]-4,n_max=length(DataBlock), col_names=FALSE))
    DataType <- 'kinetic'
  } else {
    if(DataTableType=='list')
      ReaderDF <- icontrol_endpoint_list(RawFile[DataBlock,])
    if(DataTableType=='matrix') {
      ReaderDF <- icontrol_endpoint_matrix(readxl::read_xlsx(path=ReaderFile, sheet=Sheet, skip=DataBlock[1]-1,n_max=length(DataBlock), col_names=FALSE))
      ReaderTempRaw <- RawFile[DataBlock[1]-2,2] ## Temperatue in one field as: "Temperature: 24.4 °C"
      ReaderTemp <- as.numeric(sub('^Temperature:\\s+([0-9\\.]+)\\s+\\D+$','\\1',dplyr::pull(ReaderTempRaw,1))) ## 24.4 °C even in DK locale!
      ReaderDF <- dplyr::mutate(ReaderDF, reader_chamber_temperature_C=ReaderTemp)
    }
    DataType <- 'endpoint'
  }

  list(data=cbind(reader_file_name=basename(ReaderFile), icontrol_sheet_name=Sheet,ReaderDF), header=cbind(reader_file_name=basename(ReaderFile),icontrol_sheet_name=Sheet,HeaderDF))
}

icontrol_header <- function(HeaderData) {
  ## First row is "parameter", the rest is "value"
  ## Some parametres end on ":". Skip that
  HeaderDF <- dplyr::mutate(setNames(HeaderData[,1], 'parameter'), parameter=sub(':$','',parameter))
  HeaderDF$value <- apply(as.matrix(HeaderData[,-1]),1, function(x) paste0(na.omit(x),collapse=' '))
  HeaderDF
}

icontrol_endpoint_matrix <- function(ReadData) {
  EmptyCols <- apply(as.matrix(ReadData),2,function(x) all(is.na(x)))
  ReadData <- ReadData[,!EmptyCols]
  ## Number of columns should be 1.5 x nrow
  stopifnot(nrow(ReadData)*1.5 == ncol(ReadData)-1)
  ReadMat <- as.matrix(ReadData[,-1]) ## Remove first column (row-name)
  ReadVector <- as.vector(ReadMat) ## col-wise
  WellNames <- sprintf("%s%02d",rep(LETTERS[1:nrow(ReadMat)], ncol(ReadMat)), rep(1:ncol(ReadMat),each=nrow(ReadMat)))
  ReaderDF <- tibble::as_tibble(x=list(well_name=WellNames, endpoint_value=ReadVector, plate_wells=length(ReadVector)))
  plyr::rename(ReaderDF, c(well_name=sprintf("well_name_%s",length(ReadVector)))) ## is sorted by construction
}

icontrol_kinetic_data <- function(ReadData) {
  EmptyCols <- apply(as.matrix(ReadData),2,function(x) all(is.na(x)))
  ReadData <- ReadData[,!EmptyCols]
  ## well_name_384, kinetic_step, kinetic_second, kinetic_timestamp, kinetic_value (to be renamed), reader_chamber_temperature_C, plate_wells (384)
  cycle_data <- plyr::rename(tidyr::gather(ReadData[c(1,2,3),], key="column",value="parameter_value",-X__1), c(X__1="parameter"))
  read_data <- plyr::rename(tidyr::gather(ReadData[-c(1,2,3),], key="column",value="kinetic_value",-X__1), c(X__1="well"))
  preReaderDF <- merge(cycle_data, read_data, by='column')
  ReaderDF_raw <- tidyr::spread(preReaderDF, key="parameter", value="parameter_value") ## ingestr would return this
  HeaderNames <- ReadData[c(1,2,3),1]
  Renamer <- setNames(c("kinetic_step","kinetic_second","reader_chamber_temperature_C",sprintf("well_name_%s",nrow(read_data))),c(dplyr::pull(HeaderNames[,1])),"well")
  ReaderDF <- plyr::rename(dplyr::mutate(ReaderDF_raw, well=normal_wellname(well)), Renamer)
  cbind(plyr::arrange(ReaderDF, kinetic_step, well), plate_wells=nrow(read_data))
}

normal_wellname <- function(well) {
  well <- as.character(well)
  pat <- '^(\\D)(\\d{1,2})$'
  if(any(malformed <- !grepl(pat,well))){
    futile.logger::flog.error("Well name number %s: '%s' does not match pattern '%s'", which(malformed),well[malformed],pat)
    stop()
  }
  sprintf("%s%.2d",sub(pat,'\\1',well), as.integer(sub(pat,'\\2',well)))
}

checkRows <- function(n,x) {
  if(nrow(x) != n)
    stop(sprintf("Expected %s rows. Got %s",n,nrow(x)))
  x
}