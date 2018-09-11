#' Open encrypted .RAR files
#'
#' @param pathRAR The path to the encrypted .rar file. For example, \code{"C:/Users/u007/Documents/MyNameisRAR_WinRAR.rar"}.
#' @param pathRARexe Path where the WinRAR executable is located. Default is \code{"C:/Program Files/Winrar"}.
#' @param pathExtract Path where the files have to be extracted to. This folder and all its files will be deleted if 
#' \code{RmFiles = T}.
#' @param pw Logical, indicates whether the .rar file is password protected. If \code{TRUE}, the user will be prompted to give the
#' password.
#' @param RmFiles Logical, indicates whether the extracted files have to be deleted.
#' @param RmTime The time until removal of the extracted files.
#' @param pathRscript The path where the temporary Rscript has to be saved.
#'
#' @return The extracted files will be saved in pathExtract folder and if specified, removed after a certain time.
#' @export
#' @details The extracted files and temporary folder will automatically deleted after the specified time period IF AND
#' ONLY IF the windows command window is not closed. In addition, there will be some additional temporary files left
#' that were created to execute the command (an R-script and a .bat file). These can be removed by using the function
#' RmRemainingFiles.
#'
OpenRAR <- function(pathRAR, pathRARexe = "C:/Program Files/Winrar", pathExtract = paste0(getwd(), "/tmp"), pw = c(T, F),
                    RmFiles = T, RmTime = "00:30:00",
                    pathRscript = getwd()) {
  if(!is.logical(pw))
    stop("pw argument can only be logical")
  if (!rstudioapi::hasFun("askForPassword"))
    stop("Masked input is not supported in your version of RStudio; please update to version >= 0.99.879")
  savehistory(paste0(Sys.Date(), ".Rhistory"))
  
  if(pw)
    pw = rstudioapi::askForPassword()
  
  if(!file.exists(pathRAR))
    stop(paste0("File ", pathRAR, " not found."))
  
  if(!file.exists(pathExtract))
    dir.create(pathExtract)
  
  RmTime = as.difftime(RmTime, format = "%H:%M:%S",units = "secs")[[1]]
  
  if(RmFiles) {
    Cmds = paste0("RmTmpFiles = function() unlink('",pathExtract, "', recursive = T, force = T);",
                  "Sys.sleep(", RmTime, "); RmTmpFiles()")
    writeLines(Cmds, con = paste0(pathRscript, "/RmFiles.R"))
    pathRscriptF = paste0(pathRscript, "/RmFiles.R")
    CMDcmnd = paste0('"',normalizePath(path.expand(file.path(R.home(), "bin", "x64", "R.exe")), winslash = "/"), 
                     '" CMD BATCH ',
                     '"', pathRscriptF, '"')
    writeLines(CMDcmnd, paste0(pathRscript, "/Exec.bat"))
    on.exit(shell.exec(paste0(pathRscript, "/Exec.bat")), add = T)
  }
  
  CMDwinrar = paste0('UnRar x "', pathRAR, '"')
  if(!is.logical(pw))
    CMDwinrar = paste0(CMDwinrar, ' -p', pw)
  rm(pw)
  writeLines(c(paste0("set path=", pathRARexe, "/"), paste0("cd ", pathExtract, "/"), CMDwinrar),
             paste0(pathRscript, "/UnRar.bat"))
  shell.exec(paste0(pathRscript, "/UnRar.bat"))
  on.exit(unlink(paste0(pathRscript, "/UnRar.bat"), recursive = T, force = T), add = T)
  RmRemainingFiles <<- function() lapply(c("Exec.bat", "RmFiles.R", "RmFiles.Rout"), unlink)
  on.exit(clearhistory(), add = T)
  cat("\n\nThere are still some remaining files. Use the function RmRemainingFiles to remove them after the",
      "temporary folder has been removed.\n\n")
}