# this script is to help with the emailing of items or status updates -pw
# Thu Sep 12 12:39:32 2019 ------------------------------

# https://freecarrierlookup.com/
# Use this to look up how to send text

library(RDCOMClient)

sendEmail <- function(sendto, message){
  ## init com api
  OutApp <- COMCreate("Outlook.Application")
  ## create an email
  outMail = OutApp$CreateItem(0)
  ## configure  email parameter
  outMail[["To"]] =  sendto
  outMail[["body"]] = message
  
  outMail$Send()
}


sendEmail_html<- function(sendto, rmdFile){
  htmlFile <- paste0("./html/",rmdFile,".html")
  eb <- readr::read_lines(htmlFile,n_max= -1L)
  eb2<-paste(eb, sep="", collapse="") 
  
  OutApp <- COMCreate("Outlook.Application")
  ## create an email
  outMail = OutApp$CreateItem(0)
  ## configure  email parameter
  outMail[["To"]] =  sendto
  # outMail[["Subject"]] <- "Example of html embeded within email"
  outMail[["BodyFormat"]] <- 2
  outMail[["HTMLbody"]] <- eb2
  outMail$Send()
}


sendEmailAttach <- function(sendto, subject, attchmnt){
  ## init com api
  OutApp <- COMCreate("Outlook.Application")
  ## create an email
  outMail = OutApp$CreateItem(0)
  ## configure  email parameter
  outMail[["To"]] =  sendto
  outMail[["subject"]] = subject
  outMail[["Attachments"]]$Add(attchmnt)
  
  outMail$Send()
}


sendEmailAttachCCBody <- function(sendto, subject, attchmnt, message, CCto){
  ## init com api
  OutApp <- COMCreate("Outlook.Application")
  ## create an email
  outMail = OutApp$CreateItem(0)
  ## configure  email parameter
  outMail[["To"]] =  sendto
  outMail[["CC"]] =  CCto
  outMail[["subject"]] = subject
  outMail[["body"]] = message
  outMail[["Attachments"]]$Add(attchmnt)
  
  outMail$Send()
}

