# This script will pull from an ODBC and store the data as a CSV. It will then create an outlook email with your default send signature. The script will then paste a link to the file in the body of the email. You can choose to un comment the email for display and sending.


#install.packages("RODBC")
library("RODBC")
#install.packages("RDCOMClient")
library("RDCOMClient")

#Name Source Tbl from DB
sourcetbl <- "TableName"

#Connect with ODBC
db <- odbcConnect("ODBC-Name")

#Pull Data from DB
tbl <- sqlFetch(db,sourcetbl)

#Define CSV File Location
file_string <- paste0("FileStorageLocation",sourcetbl,".csv")

#Write to CSV
write.csv(tbl, file=file_string, row.names = FALSE)

#Create Mail Item
olMailItem = 0
OutApp <- COMCreate("Outlook.Application")
outMail <- OutApp$CreateItem(olMailItem)

#Hyperlink strings
file_string2 <- shQuote(file_string)
link_string <- paste0("<a href=",file_string2,">",file_string,"</a>")

#Signture Pull
outMail$GetInspector()
signature = outMail[["HTMLBody"]]

#Message Desgin
recipients <- "Email"
subject <- paste0("Link to ",sourcetbl," CSV")
body_string <- paste0("Here is a link to ",sourcetbl,'<p>',"Thanks!",'</p>')
body_string_full <- paste0('<p>',body_string,'</p>','<p>',link_string, signature, '</p>')

#Message Contents
outMail[["Recipients"]]$Add(recipients)
outMail[["Subject"]] = subject
outMail[["HTMLBody"]] = body_string_full

outMail$Display()
# outMail$Send()

#Clear Dataset
rm(list=ls())