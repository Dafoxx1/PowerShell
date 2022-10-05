<#
.SYNOPSIS
  Outlook Signature Creation
.DESCRIPTION
Creates html signature blocks using AD and HTML
.INPUTS
User data from AD
Please Change lines 38,39,52,135,157
.OUTPUTS
Saves a copy on C:\Signatures  Copy png file into email signature file
.NOTES
  Version:        4.0
  Author:         James Kimble
  Created:       03/25/2021
  Modified:       10/04/2022
#>
#######Making and removing old files folders to prevent overwrite errors
RD C:\Signature\Signature.png -Force -ErrorAction SilentlyContinue
# Getting Active Directory information for current user
$user = (([adsisearcher]"(&(objectCategory=User)(samaccountname=$env:username))").FindOne().Properties)
MD C:\Signature -ErrorAction SilentlyContinue

################### Manually add User name run command below          ############################
#$Logon = Read-Host -Prompt "Input user logon name"
#$user = (([adsisearcher]"(&(objectCategory=User)(samaccountname=$Logon))").FindOne().Properties)
##################################################################################################

if($user) {
  # Create the signatures folder and sets the name of the signature file
  $folderlocation = $Env:appdata + '\Microsoft\signatures'
  $filename = "Signature"
  $file = "C:\Signature\$filename"
  if(!(Test-Path -Path $folderlocation )){
      New-Item -ItemType directory -Path $folderlocation
  }

  # Company name and logo
  $companyName = "YOUR COMPANY NAME"
  $logo = "WEB ADDRESS TO YOUR COMPANY LOGO" # Please note that if you do include a logo it must be located somewhere on the internet that the public has access to, many users upload it to their website.

  # Get the users properties (These should always be in Active Directory and Unique)
  if($user.name.count -gt 0){$displayName = $user.name[0]}
  if($user.title.count -gt 0){$jobTitle = $user.title[0]}
  if($user.homephone.count -gt 0){$directDial = $user.homephone[0].trimstart("1")}
  if($user.mobile.count -gt 0){$mobileNumber = $user.mobile[0].trimstart("1")}
  if($user.mail.count -gt 0){$email = $user.mail[0]}
  $website = "garbennett.com"
  if($user.telephonenumber.count -gt 0){$telephone = $user.telephonenumber[0].trimstart("1 ")}
  if($user.physicaldeliveryofficename.count -gt 0){$office = $user.physicaldeliveryofficename}
  ########################### Building the Address profile from Address book

  $X = Import-Csv -Path "\\LOCATION TO AN ADDRESS BOOK.csv" 
  $address = $X | ?{$_.Location -like $user.company.Split(' ')[-1]} 
  $street = $address.Street
  $city = $address.City
  $state = $address.State
  $zipcode = $address.Zipcode

     
 
  #########################


  
  # Building HTML
  $signature = 
  @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Signature</title>
    <link rel="stylesheet" href="https://use.typekit.net/emh3aht.css">
    <style>
        body{
            flex: 20%;
            display: flex;
            margin: 0px;
            padding: 0px;
            height: 200px;
        }
        p{
            margin: 10px;
        }
        .logo{
            padding-right: 0px;
            padding-left: 10px;
            padding-bottom: 10px;
            padding-top: 20px;
            width: 280px;
            height: 100px;
        }
        .info{
            font-family: Arial, Helvetica, sans-serif;
            color: #0a2942;
            letter-spacing: 1.5px;
            line-height: 1px;
            width: 450px;
            text-align: left;
            justify-content: Left;
            margin: 0%;
            padding-top: 0px;

        }
        .name{
            font-size: 30px;
            font-weight: lighter;
            margin-top: 15px;
            margin-bottom: 5px;
        }
        .title{
            font-size: 14px;
            padding-top: 0px;
        }
        .adpull{
            line-height: .25px;
            font-size: 13px;
            margin; 0px;
            padding; 0px;
        }
        .website{
            font-size: 10px;
            font-strong
        }
        .adpull p{
            margin; 0px;
            padding; 0px;
        }
    </style>
</head>
<body >
    <div class="logo">
        <img src="LOCATION TO YOUR COMPANY LOGO ON THE WEB" />
    </div>
    <div class="info">
        <p class="name"> 
            $(if($displayName){""+$displayName+"</font>"})
        </p>
        <p class="title">
            $(if($jobTitle){ "$jobTitle" })
        </p>
        <div class="adpull">
            <p>
                $(if($street){$street+"   | " })
                $(if($city){ $city+", " })
                $(if($state){ $state })
                $(if($zipCode){ $zipCode})
                <br>
            
                $(if($telephone){"Direct: "+$telephone+"   | "})
                $(if($mobileNumber){"Mobile: "+"$mobileNumber"})
            </p>
        </div>
    <p class="website">
        COMPANY NAME
    </p>
    </div>
    <div>

    </div>    
</body>
</html>
"@

  
  
  #######################################################

  
  
  # Save the HTML to the signature file
  $style + $signature | out-file "$file.htm" -encoding ascii -Force
}
#### Screen Shot HTML page into PNG file ###########
If ((test-path "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") -eq $true){

    & 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe' --headless --disable-gpu --hide-scrollbars --window-size=850,150 --screenshot=C:\Signature\Signature.png file:C:\Signature\Signature.htm --force
}
Else{
    & 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' --headless --disable-gpu --hide-scrollbars --window-size=850,150 --screenshot=C:\Signature\Signature.png file:C:\Signature\Signature.htm --force
}
