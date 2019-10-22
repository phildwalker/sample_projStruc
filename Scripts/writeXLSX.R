#+++++++++++++++++++++++++++
# WriteMultTabs
#+++++++++++++++++++++++++++++
# fileName : the path to the output file
# listObj : a list of data to write to the workbook


library(openxlsx)

writeMultTabs <- function(fileName, listObj){
  ## Create a blank workbook
  wb <- createWorkbook()
  
  ## Loop through the list of split tables as well as their names
  ##   and add each one as a sheet to the workbook
  Map(function(data, name){
    addWorksheet(wb, name)
    writeData(wb, name, data)
  }, listObj, names(listObj))
  
  ## Save workbook to working directory
  saveWorkbook(wb, file = fileName, overwrite = TRUE) 
}











