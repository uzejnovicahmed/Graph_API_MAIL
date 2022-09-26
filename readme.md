# Send Mail via Graph API

Module to Send Mails via Graph API.

## Getting Started

### Prerequisites

* PS Version 5.1 and above
* APP Registration with Send Mail Permissons

### Installing

Download this repo and import the `Graph-Mail.psm1` module file into your PowerShell session.

```
Import-Module .\Graph-Mail.psm1;
```

## Development

For ease of deployment (copy-paste) all functions are includen in file `Graph-Mail.psm1`.

## Examples

```
#$TID = "943a3020-dc82-4n25-aa35-837d3032asdf99ed" # dummy
#$CID = "97e6600c-1b3c-4277-9k44-83e0b78sdaf93b38" # dummy
#$CSECRET = ".GT8Q~tclExvaWdfadghdfahadf1gr~~.YZe-68j7dtq" #dummy


initializing-GraphMail -TENANTID $TID -CLIENTID $CID -CLIENTSECRET $CSECRET -MAILSENDER "AdeleV@dev.onmicrosoft.com" -MAILRECIPIENT "mail.recipient@dev.onmicrosoft.com"

Send Mail without attachment:

Send-GraphMail -HEADERS $headers

Send Mail with attachemnt:

Send-GraphMail -HEADERS $headers -Attachmentpath C:\temp\test.txt


```

### Usage in PowerShell

Initializize Graph Mail. 

Send-Graph Mail

## Resources

* https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-exceptions?view=powershell-7.2

## Missing features

```
Errorhandling
Attachmenttypes

```
