#' Import password protected excel files
#'
#' @param path The path to the excel file.
#' @param sheet The sheet that has to be read. Default is \code{NULL}, which causes the first sheet to be imported.
#' @param pw Password of the excel file. Default is \code{NULL} and the user is then prompted to
#' give the password using \code{\link[rstudioapi]{askForPassword}}.
#' @param rmHist Logical, indicates whether the history has to be deleted.
#'
#' @return Returns the excel file as a dataframe.
#' @details This function needs the package \code{RDCOMClient}, which can be installed using: \cr
#' \code{library(devtools)} \cr
#' \code{install_github('omegahat/RDCOMClient')} \cr
read_excel_pw <- function(path, sheet = NULL, pw = NULL, rmHist = T) {
  
  if(!"RDCOMClient" %in% installed.packages())
    stop("Package RDCOMClient is required for this function.")
  
  if(!"RDCOMClient" %in% names(sessionInfo()$otherPkgs))
    require(RDCOMClient, quietly = T)
  
  if (!rstudioapi::hasFun("askForPassword"))
    stop("Masked input is not supported in your version of RStudio; please update to version >= 0.99.879")
  
  if(!file.exists(path))
    stop("File not found.")
  
  if(!is.null(sheet))
    if(!is.character(sheet))
      stop("Provide character variable")
  else if(length(sheet) > 1)
    stop("Only functionality provided to import one sheet")
  
  if(rmHist)
    on.exit(clearhistory())
  
  # import data
  if(is.null(pw))
    pw   = rstudioapi::askForPassword()
  eApp = COMCreate("Excel.Application")
  wk   = eApp$Workbooks()$Open(Filename = path,Password = pw)
  tf   = tempfile()
  if(is.null(sheet))
    wk$Sheets(1)$SaveAs(tf, 3)
  else
    wk$Worksheets(sheet)$SaveAs(tf, 3)
  
  Df = read.delim(sprintf("%s.txt", tf), header = TRUE, sep = "\t")
  
  # Close Excel
  wk$Close(SaveChanges = FALSE)
  eApp$Quit()
  rm(wk, eApp)
  gc()
  
  # Remove tmp files
  unlink(sprintf("%s.txt", tf), recursive = T, force = T)
  unlink(tf, force = T)
  rm(pw)
  Df
}


# Function to remove R history
clearhistory <- function() {
  write("", file=".blank")
  loadhistory(".blank")
  unlink(".blank")
}

#' Function to read all sheets of a password protected excel file
#'
#' @param path The path to the excel file.
#' @param pw Password of the excel file. Default is \code{NULL} and the user is then prompted to
#' give the password using \code{\link[rstudioapi]{askForPassword}}.
#' @param rmHist Logical, indicates whether the history has to be deleted.
#'
#' @return Returns a list with all the excel sheets.
#' @details This function needs the package \code{RDCOMClient}, which can be installed using: \cr
#' \code{library(devtools)} \cr
#' \code{install_github('omegahat/RDCOMClient')} \cr
read_excel_allsheets_pw <- function(path, pw = NULL, rmHist = T) {
  if(!"RDCOMClient" %in% installed.packages())
    stop("Package RDCOMClient is required for this function.")
  
  if(!"RDCOMClient" %in% names(sessionInfo()$otherPkgs))
    require(RDCOMClient, quietly = T)
  
  if (!rstudioapi::hasFun("askForPassword"))
    stop("Masked input is not supported in your version of RStudio; please update to version >= 0.99.879")
  
  if(!file.exists(path))
    stop("File not found.")
  
  if(rmHist)
    on.exit(clearhistory())
  
  # import data
  if(is.null(pw))
    pw   = rstudioapi::askForPassword()
  
  eApp = COMCreate("Excel.Application")
  wk   = eApp$Workbooks()$Open(Filename = path,Password = pw)
  NrSheets    = wk$Sheets()$Count()
  NamesSheets = sapply(1:NrSheets, function(i) wk$Sheets(i)$Name()) 
  
  AllSheets = list() 
  for(i in NamesSheets) {
    tf   = tempfile()
    wk$Worksheets(i)$SaveAs(tf, 3)
    AllSheets[[i]] = read.delim(sprintf("%s.txt", tf), header = TRUE, sep = "\t")
  }
  
  # Close Excel
  wk$Close(SaveChanges = FALSE)
  eApp$Quit()
  rm(wk, eApp)
  gc()
  
  # Remove tmp files
  unlink(sprintf("%s.txt", tf), recursive = T, force = T)
  unlink(tf, force = T)
  rm(pw)
  
  AllSheets
}  


#' Temporary fix for savexlsx function if there are missing values.
#'
#' @param x An object of class dataframe or a list. 
#'
#' @return
#'
#' @examples
#' data('mtcars')
#' mtcars[1, ] = NA
#' test = xlsxFix(mtcars)
#' savexlsx(test)
#' 
#' @details This function has to be used in case of missing values. If not, the missing values will be replaced by large values
#' in the excel file.
#' 
xlsxFix <- function(x) {
  if("list" %in% class(x))
    lapply(x, function(y) {
      as.data.frame(lapply(y, function(z) {
        if(is.factor(z))
          z = as.character(z)
        z[is.na(z)] = ""
        return(z)
      }), check.names = F)
    })
  else
    as.data.frame(lapply(x, function(z) {
      if(is.factor(z))
        z = as.character(z)
      z[is.na(z)] = ""
      return(z)
    }), check.names = F) 
}

#' Write a dataframe or list to an excel file (with or without password).
#'
#' @param Object Input object, must be of type dataframe or list.
#' @param path The path for the excel file. Default is current working directory.
#' @param SheetName Optional, name of the sheet.
#' @param pw Logical, indicates whether the file has to be password protected or not. If \code{TRUE}, the user will be prompted to
#' give the password using \code{\link[rstudioapi]{askForPassword}}.
#'
#' @seealso \code{\link{xlsxFix}}
#' @examples
#' 
#' # Without missing values
#' data(mtcars)
#' savexlsx(mtcars)
#' 
#' # With missing values
#' data('mtcars')
#' mtcars[1, ] = NA
#' test = xlsxFix(mtcars)
#' savexlsx(test)
#' 
savexlsx <- function(Object, path = paste0(getwd(), "/", deparse(substitute(Object)), ".xlsx"),
                     SheetName = NULL, pw = F, rmHist = pw) {
  if(!is.list(Object))
    stop("Only objects of type list or dataframe permitted.")
  if(!is.logical(pw))
    stop("pw can only be TRUE or FALSE.")
  if(!"RDCOMClient" %in% installed.packages())
    stop("Package RDCOMClient is required for this function.")
  
  if(!"RDCOMClient" %in% names(sessionInfo()$otherPkgs))
    require(RDCOMClient, quietly = T)
  
  if (!rstudioapi::hasFun("askForPassword"))
    stop("Masked input is not supported in your version of RStudio; please update to version >= 0.99.879")
  
  if(anyNA(Object) & !("xlsxFix" %in% class(Object)))
    stop("Missing values detected! Use xlsxFix function.")
  
  if(pw & rmHist)
    on.exit(clearhistory())
  
  if(!is.null(SheetName) & !is.character(SheetName))
    stop("The sheet name has to be of type character")
  if(is.list(Object) & !is.data.frame(Object)) {
    savexlsxlist(Object, path, SheetName, pw)
  } else if(is.data.frame(Object)){
    if(is.null(SheetName))
      SheetName = print(deparse(substitute(Object)))
    xls = COMCreate("Excel.Application")
    wb  = xls[["Workbooks"]]$Add(1)
    rdcomexport <- function(x) {
      sh = wb[["Worksheets"]]$Add()
      sh[["Name"]] = SheetName
      exportDataFrame(x, at = sh$Range("A1"))
    }
    rdcomexport(Object)
    xls$Sheets("Sheet1")$Delete()
    path <- gsub("/" , "\\\\" , path)
    if(pw) {
      pw = rstudioapi::askForPassword()
      wb$SaveAs(path, password = pw)
    } else {
      wb$SaveAs(path)
    }
    wb$Close(path)
  }
}

savexlsxlist <- function(.list, path = paste0(getwd(), "/", deparse(substitute(.list)), ".xlsx"),
                         SheetName = names(.list), pw = F, FixFactor = T, RowNames = F) {
  if(is.null(SheetName))
    SheetName = names(.list)
  else if(!all(is.character(SheetName)))
    stop("The sheet name has to be of type character")
  
  if(!is.list(.list))
    stop("Only lists are permitted!")
  
  if(length(SheetName) != length(.list))
    stop("Length of sheet names is not equal to number of entries in the list!")
  
  if(!all(sapply(.list, class) == "data.frame"))
    .list = lapply(.list, as.data.frame)
  
  if(FixFactor) {
    .list = lapply(.list,
                   function(y) {
                     as.data.frame(lapply(y, 
                                          function(z) {
                                            if("factor" %in% class(z))
                                              as.character(z)
                                            else
                                              z
                                          }))
                   })
  }
  xls = COMCreate("Excel.Application")
  wb  = xls[["Workbooks"]]$Add(1)
  rdcomexportList <- function(x) {
    sh = wb[["Worksheets"]]$Add()
    sh[["Name"]] = names(.list)[sapply(.list, function(y) identical(x, y))]
    if(RowNames)
      x = cbind.data.frame(" " = rownames(x), x)
    exportDataFrame(x, at = sh$Range("A1"))
  }
  lapply(.list, rdcomexportList)
  xls$Sheets("Sheet1")$Delete()
  path <- gsub("/" , "\\\\" , path)
  if(pw) {
    pw = rstudioapi::askForPassword()
    wb$SaveAs(path, password = pw)
  } else {
    wb$SaveAs(path)
  }
  wb$Close(path)
}
