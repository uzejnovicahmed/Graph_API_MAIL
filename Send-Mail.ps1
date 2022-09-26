Import-Module .\Graph-Mail.psm1


#$TID = "913a44420-dc82-4675-aa35-837d303299ed"  #not real id 
#$CID = "99e6600c-1a3c-4600-9b44-8fe0b74333b38"  #not real id
#$CSECRET = "adfgasdgaerzhhadfgadffgdsagdfhdfhgfdh" #dumysecret
$TID = "913a0020-dc82-4a35-aa35-837d303299ed"
$CID = "99e6600c-1a3c-4100-9b44-8fe0b7893b38"
$CSECRET = ".GT8Q~tclExvauVnDuS7dMW1gr~~.YZe-68j7dtq"

Initializing-GraphMail -TENANTID $TID -CLIENTID $CID -CLIENTSECRET $CSECRET -MAILSENDER "AdeleV@uzeah.onmicrosoft.com" -MAILRECIPIENT "ahmed.uzejnovic@baseit.at"


Send-GraphMail -HTML_CONTENT $data -Subject "Testsubject"

