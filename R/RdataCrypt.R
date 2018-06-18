#' Load encrypted Rdata files
#'
#' @param filename Name of the encrypted file. E.g. \code{"Test.Rdata.gpg"}.
#' @param path Directory where the file is located. Default is current working directory.
#' @param RmOrig Logical, indicates whether the original file should be removed. Default is \code{TRUE}.
#' @param ... Arguments to be passed to \code{\link{decrypt}}.
#'
#' @return The encrypted Rdata file is loaded into the global environment.
#'
LoadRcrypt <- function(filename = ".Rdata.gpg", path = getwd(), RmOrig = T, ...) {
  if(! "rcrypt" %in% names(utils::sessionInfo()$otherPkgs))
    require(rcrypt, quietly = T)
  
  if(!identical(path, getwd())) {
    OldWd = getwd()
    setwd(path)
    on.exit(setwd(OldWd))
  }
  
  
  if(! file.exists(filename))
    stop("File not found")
  
  if(file.exists(gsub(".gpg", "", filename))) {
    if(!RmOrig) {
      Cmmnd = readline("There already is a non-encrypted file with the same name. Delete? y/n  ")
      if(Cmmnd == "n")
        stop("Function stopped by the user.")
      cat("\n\nRemoving original file...\n\n")
    }
    file.remove(gsub(".gpg", "", filename))
  }
  decrypt(filename, ...)
  load(gsub(".gpg", "", filename), envir = .GlobalEnv)
  file.remove(gsub(".gpg", "", filename))
}

#' Encrypt Rdata files
#'
#' @param filename The name of the Rdata file.
#' @param path The directory where the Rdata file is stored. Default is current working directory.
#' @param RmOld Logical, indicates whether the Rdata file has to be removed.
#' @param askPW Logical, indicates whether the password has to be given using the \code{\link[rstudioapi]{askForPassword}} function.
#' @param ... Arguments to be passed to \code{\link{encrypt}}
#'
#' @return The Rdata file is encrypted and will be saved in the same directory as specified in the \code{path} argument.
EncryptRdata <- function(filename = ".Rdata", path = getwd(), RmOld = T, askPW = T, ...) {
  if(! "rcrypt" %in% names(utils::sessionInfo()$otherPkgs))
    require(rcrypt, quietly = T)
  
  assign("Encr", T, envir = .GlobalEnv)
  
  if(!identical(path, getwd())) {
    OldWd = getwd()
    setwd(path)
    on.exit(setwd(OldWd))
  }
  
  if(file.exists(filename)) {
    if(!RmOld) {
      Cmmnd = readline("There already exists a Rdata file. Remove? y/n")
      if(Cmmnd == "n")
        stop("Function stopped by the user")
    }
    cat("\n\nRemoving file...\n\n")
    on.exit(file.remove(filename))
  }
  
  if(file.exists(paste0(filename, ".gpg"))) {
    if(!RmOld) {
      Cmmnd = readline("There already exists an encrypted Rdata file. Remove? y/n")
      if(Cmmnd == "n")
        stop("Function stopped by the user")
    }
    cat("\n\nRemoving file...\n\n")
    file.remove(paste0(filename, ".gpg"))
  }
  
  
  save.image(filename)
  cat("\n\nEncrypting Rdata...\n\n")
  if(askPW){
    pw = rstudioapi::askForPassword()
    encrypt(filename, passphrase = pw, ...)
  } else {
    encrypt(filename, ...)
  }
  cat("\n\nRemoving temporary image...\n\n")
  if(file.exists(filename))
    file.remove(filename)
}