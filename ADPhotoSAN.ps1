Function ResizeImage() {
    param([String]$ImagePath, [Int]$Quality = 90, [Int]$targetSize, [String]$OutputLocation)
 
    Add-Type -AssemblyName "System.Drawing"
 
    $img = [System.Drawing.Image]::FromFile($ImagePath)
 
    $CanvasWidth = $targetSize
    $CanvasHeight = $targetSize
 
    #Encoder parameter for image quality
    $ImageEncoder = [System.Drawing.Imaging.Encoder]::Quality
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($ImageEncoder, $Quality)
 
    # get codec
    $Codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where {$_.MimeType -eq 'image/jpeg'}
 
    #compute the final ratio to use
    $ratioX = $CanvasWidth / $img.Width;
    $ratioY = $CanvasHeight / $img.Height;
 
    $ratio = $ratioY
    if ($ratioX -le $ratioY) {
        $ratio = $ratioX
    }
 
    $newWidth = [int] ($img.Width * $ratio)
    $newHeight = [int] ($img.Height * $ratio)
 
    $bmpResized = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
    $graph = [System.Drawing.Graphics]::FromImage($bmpResized)
    $graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
 
    $graph.Clear([System.Drawing.Color]::White)
    $graph.DrawImage($img, 0, 0, $newWidth, $newHeight)
 
    #save to file
    $bmpResized.Save($OutputLocation, $Codec, $($encoderParams))
    $bmpResized.Dispose()
    $img.Dispose()
}
Remove-item "PATH TO ERROR LOG" -ErrorAction SilentlyContinue
#outpath
#Main folder with serveral subfolders. Pictures will go inside a folder with users name then will be moved around depending on status.
$OutputFolder = "\\LOCATION OF FOLDER\_AD Import\"
$Path = "\\LOCATION OF FOLDER\employeeheadshots"
$Processed = "\\LOCATION OF FOLDER\_Processed"
#Part 1 - Query AD to make folder structure for employee photos, looks for all active users and then creates a folder with their display name and if there is a photo with the correct name add them to said folder
$AD = (Get-ADUser -Filter * -SearchBase "OU=Staff,OU=GB,DC=gb,DC=local" -Properties enabled | Sort -Property Name |Select -Property Name).name
Foreach($Employee in $AD)
{
    If((Test-path "$path\_Processed\$employee") -ne $true)
    {
        MD -Path "$path\$employee" -ErrorAction SilentlyContinue      
    }
    #Troubleshooting SOP
    If((test-path ("$path\"+"$employee"+".jpg")) -eq "$true")
    {
        Move-Item -Path ("$path\$employee"+".jpg") -Destination $path\$employee 
    }
}
#Any folders you would like to add should have a _Name to not be processed
$X = GCI $Path | ?{$_.name -notlike "_*"}
If($X -ne $null){
    $X.fullname
    $pictures = GCI $X.fullname -Recurse

    #Part 2 - This looks through all employee folders for a JPG, resizes them and drops them into a repository that then updates AD with photos. 
    Foreach($EmployeePIC in $Pictures)
    {    
        ResizeImage ($EmployeePIC).FullName 150 150 ($OutputFolder+$EmployeePIC)
        #outpath
        $resized = ($OutputFolder + $EmployeePIC)
        If ((Test-path -Path $outputfolder\$EmployeePIC) -eq $true)
        {
            Move-item -Path $EmployeePIC.PSParentPath -Destination $processed
        }

        $Name2 = $EmployeePIC.directory.Name
        $name2
        $ADuser = (Get-ADUser -Filter {(Name -eq $name2)}).samaccountname
        If(((Get-ADuser $aduser -Properties Thumbnailphoto).thumbnailphoto) -eq $null)
        {
            $EmployeePIC.Name | Out-File -FilePath "PATH TO ERROR FOLDER\AD photo Error.txt" -Append -Force 
            $ADphoto = [byte[]](Get-Content $resized -Encoding byte)
            Set-ADUser $ADuser -Replace @{thumbnailPhoto=$ADphoto} 
        }

    }
}

#Identifies terminated users to help reduce overall filesize.
$Termed = (Get-ADUser -Filter * -SearchBase "OU=Terminated,OU=DOMAIN,DC=DOMAIN,DC=TOP LEVEL" -Properties enabled | Sort -Property Name |Select -Property Name).name
ForEach($Term in $Termed){
    If((Test-path "$path\_Processed\$term") -eq $true){
        Move-Item -Path ("$path\_processed\$term") -Destination $path\_Terminated -ErrorAction SilentlyContinue
    }
     If((Test-path "$path\$term") -eq $true){
        Move-Item -Path ("$path\$term") -Destination $path\_Terminated -ErrorAction SilentlyContinue -Force
    }
}
