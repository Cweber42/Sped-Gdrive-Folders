#Sped Gdrive folder creation version 1
#Must be a member of the Gdrive folder and have file stream installed on computer
#Cannot make root level folders in Gdrive via cmdline, must use API or create it manually via the web.
#Creates student folders in the shared drives that are mapped/specified
#Moves existing students between gdrive folders when needed
#Creates new school year folder on July 1 each year.


Start-Transcript -path "G:\Shared drives\Sped-Archive\Script Logs\$(Get-date -format yyyy-MM-dd-HH-mm-ss).log"
#Cognos Download
c:\scripts\CognosDownload.ps1 -report "504-students" -savepath C:\scripts\Sped -ReportStudio -Cognosfolder "Special-Ed Reports"
c:\scripts\CognosDownload.ps1 -report "Sped-Students" -savepath C:\Scripts\Sped -ReportStudio -Cognosfolder "Special-ed Reports"
function RemoveSpecials ([String]$in)
{
 $in = $in -replace("\(","") #Remove ('s
 $in = $in -replace("\)","") #Remove )'s
 $in = $in -replace("\.","") #Remove Periods
 $in = $in -replace("\'","") #Remove Apostrophies
 $in = $in -replace("`"","") #remove double quotes
 $in = $in -replace("\'","") #Remove single quotes
 return $in
}
$currentmonth = (Get-date).month
IF ($currentmonth -eq "7"){
     $curyear = (Get-date).Year
     $nxtyear = (Get-date).year+1
     $currentyear = [String]$curyear +'-'+[String]$nxtyear
}
$Spedstus = Import-Csv "C:\scripts\Sped\SpED-Students.csv"
$504stus = import-csv "C:\scripts\Sped\504-Students.csv"

#Create folders
Foreach ($Spedstu in $Spedstus){
$fname = (Removespecials($Spedstu."Student First Name"))
$lname = (removespecials($spedstu."Student Last Name"))
$id = $spedstu."State Report ID"
$name = $fname + " " + $lname + " " + $id
Switch ($spedstu."Current Building"){
    "25" {$studir = "G:\Shared drives\PS-Sped\"}
    "26" {$studir = "G:\Shared drives\IS-Sped\"}
    "27" {$studir = "G:\Shared drives\HS-Sped\"}
    "28" {$studir = "G:\Shared drives\MS-Sped\"}
    "29" {$studir = "G:\Shared drives\JH-Sped\"}
}

If (Test-Path $studir$name -PathType Container)
    {Write-host "$studir$name already exists" -BackgroundColor Red}
    elseif (Get-ChildItem -Path "G:\Shared drives\PS-Sped" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir" -BackgroundColor Green 
        Get-ChildItem -Path "G:\Shared drives\PS-Sped" -Recurse -Include "$name"| Move-Item -Destination $studir 
        Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name moved" -Body "$Name was moved from G:\Shared drives\PS-Sped to $studir"
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\IS-Sped" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\IS-Sped" -Recurse -Include "$name"| Move-Item -Destination $studir        
        Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name moved" -Body "$Name was moved from G:\Shared drives\IS-Sped to $studir"
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\MS-Sped" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\MS-Sped" -Recurse -Include "$name"| Move-Item -Destination $studir 
        Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name moved" -Body "$Name was moved from G:\Shared drives\MS-Sped to $studir"       
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\JH-Sped" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\JH-Sped" -Recurse -Include "$name"| Move-Item -Destination $studir 
        Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name moved" -Body "$Name was moved from G:\Shared drives\JH-Sped to $studir"       
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\HS-Sped" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\HS-Sped" -Recurse -Include "$name"| Move-Item -Destination $studir   
        Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name moved" -Body "$Name was moved from G:\Shared drives\HS-Sped to $studir"     
    }
    else{
        New-Item -path $studir -ItemType directory -name $name
        New-Item -path "$studir$name\" -ItemType director -name "Old Documents"
        New-Item -path "$studir$name\" -ItemType director -name "SY$currentyear"
        New-Item -path "$studir$name\" -ItemType director -name "Evaluations"
        Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name created" -Body "$Name was created in $studir"
    }#Close If/elseif/Else for 
    If (Test-Path "$studir$name\SY$currentyear" -PathType Container)
    {Write-host "$studir$name\SY$currentyear" already exists -BackgroundColor Red}
    Else{
    New-Item -path "$studir$name\" -ItemType director -name "SY$currentyear"
     }#Close If/Elseif/Else for testing student folder exists/moving/creating folders
}#Close Foreach loop for Sped students
#504 students managed by sped department
Foreach ($504stu in $504stus){
    $fname = (Removespecials($504stu."Student First Name"))
    $lname = (removespecials($504stu."Student Last Name"))
    $id = $504stu."State Report ID"
    $name = $fname + " " + $lname + " " + $id
    Switch ($504stu."Current Building"){
        "25" {$studir = "G:\Shared drives\PS-504\"}
        "26" {$studir = "G:\Shared drives\IS-504\"}
        "27" {$studir = "G:\Shared drives\HS-504\"}
        "28" {$studir = "G:\Shared drives\MS-504\"}
        "29" {$studir = "G:\Shared drives\JH-504\"}
    }
    
    If (Test-Path $studir$name -PathType Container)
        {Write-host "$studir$name already exists" -BackgroundColor Red}
        elseif (Get-ChildItem -Path "G:\Shared drives\PS-504" -Recurse -Include "$name") {
            Write-Host "Moving $name to $studir" -BackgroundColor
            Get-ChildItem -Path "G:\Shared drives\PS-504" -Recurse -Include "$name"| Move-Item -Destination $studir 
            Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name moved" -Body "$Name was moved from G:\Shared drives\PS-504 to $studir"
        }
        elseif (Get-ChildItem -Path "G:\Shared drives\IS-504" -Recurse -Include "$name") {
            Write-Host "Moving $name to $studir"
            Get-ChildItem -Path "G:\Shared drives\IS-504" -Recurse -Include "$name"| Move-Item -Destination $studir        
            Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name moved" -Body "$Name was moved from G:\Shared drives\IS-504 to $studir"
        }
        elseif (Get-ChildItem -Path "G:\Shared drives\MS-504" -Recurse -Include "$name") {
            Write-Host "Moving $name to $studir"
            Get-ChildItem -Path "G:\Shared drives\MS-504" -Recurse -Include "$name"| Move-Item -Destination $studir 
            Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name moved" -Body "$Name was moved from G:\Shared drives\MS-504 to $studir"       
        }
        elseif (Get-ChildItem -Path "G:\Shared drives\JH-504" -Recurse -Include "$name") {
            Write-Host "Moving $name to $studir"
            Get-ChildItem -Path "G:\Shared drives\JH-504" -Recurse -Include "$name"| Move-Item -Destination $studir 
            Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name moved" -Body "$Name was moved from G:\Shared drives\JH-504 to $studir"       
        }
        elseif (Get-ChildItem -Path "G:\Shared drives\HS-504" -Recurse -Include "$name") {
            Write-Host "Moving $name to $studir"
            Get-ChildItem -Path "G:\Shared drives\HS-504" -Recurse -Include "$name"| Move-Item -Destination $studir   
            Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name moved" -Body "$Name was moved from G:\Shared drives\HS-504 to $studir"     
        }
        else{
            New-Item -path $studir -ItemType directory -name $name
            New-Item -path "$studir$name\" -ItemType director -name "Old Documents"
            New-Item -path "$studir$name\" -ItemType director -name "SY$currentyear"
            New-Item -path "$studir$name\" -ItemType director -name "Evaluations"
            Send-MailMessage -SmtpServer "smtp-relay.gmail.com" -From "fromalert@domain.com" -To 'alertemail1@domain.com','alertemail2@domain.com' -Subject "Student folder for $name created" -Body "$Name was created in $studir"
        } #Close If/Elseif/Else for testing student folder exists/moving/creating folders
        If (Test-Path "$studir$name\SY$currentyear" -PathType Container)
        {Write-host "$studir$name\SY$currentyear" already exists}
        Else{
        New-Item -path "$studir$name\" -ItemType director -name "SY$currentyear"
         } #close IF/else for Currentyear folder
    }#Close foreach loop for 504 students
Stop-Transcript    
    
