#SYN
#This is a list of commands to perform various Termination activities within AD / AZURE
#
#
$Employee = read-host "Employee Name"
$Attributes = Get-ADUser -filter{Name -like $Employee} -Properties UserPrincipalName,samaccountname,distinguishedName,Name | Select -Property UserPrincipalName,samaccountname,distinguishedName,Name
$SAM = $Attributes.samaccountname

#Hide from Global Access List *Only if setup to pull contact info from AD
Set-ADUser $sam -Enabled $false -replace  @{'msDS-cloudExtensionAttribute1'="HideFromGAL"}

#Remove from AD Groups
$Groups = Get-ADPrincipalGroupMembership -Identity $sam | select -Property name
foreach ($Group in $Groups) {
    If($Group.name -ne "Domain Users"){
        Remove-ADGroupMember -identity $Group.name -Members $Sam 
    } 
}

#Move OU
Move-ADObject $Attributes.DistinguishedName -TargetPath "OU=Terminated,OU=domainName,DC=domainName,DC=topLevel"
$UPN= $Attributes.UserPrincipalName

#Azure (0365 disable sign in)
Connect-AzureAD
Set-azureAduser -ObjectId $UPN -AccountEnabled $false
$AzureID = (Get-AzureADUser -ObjectId $UPN).objectID
$AzureGroups = Get-azureADUserMembership -ObjectId $UPN
Foreach ($AzureGroup in $AzureGroups){
    If($AzureGroup.DisplayName -ne "All Users"){
        Remove-AzureADGroupMember -ObjectId $AzureGroup.ObjectId -MemberId $AzureID
    }
}

#O365 remove licenses
Connect-MsolService
$MSU = Get-MsolUser -UserPrincipalName $UPN
$Licenses = $MSU.Licenses.accountskuid
foreach ($License in $Licenses){
    Set-MSOLuserLicense -UserPrincipalName $UPN -RemoveLicenses $License
}

