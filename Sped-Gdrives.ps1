#Sped Gdrive folder creation version 1
#Must be a member of the Gdrive folder and have file stream installed on computer
#Cannot make root level folders in Gdrive via cmdline, must use API or create it manually via the web.
#Creates student folders in the shared drives that are mapped/specified
#Moves existing students between gdrive folders when needed
#Creates new school year folder on July 1 each year.

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
    {Write-host "$studir$name already exists"}
    elseif (Get-ChildItem -Path "G:\Shared drives\PS-Sped" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\PS-Sped" -Recurse -Include "$name"| Move-Item -Destination $studir        
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\IS-Sped" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\IS-Sped" -Recurse -Include "$name"| Move-Item -Destination $studir        
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\MS-Sped" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\MS-Sped" -Recurse -Include "$name"| Move-Item -Destination $studir        
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\JH-Sped" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\JH-Sped" -Recurse -Include "$name"| Move-Item -Destination $studir        
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\HS-Sped" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\HS-Sped" -Recurse -Include "$name"| Move-Item -Destination $studir        
    }
    else{
        New-Item -path $studir -ItemType directory -name $name
        New-Item -path "$studir$name\" -ItemType director -name "Old Documents"
        New-Item -path "$studir$name\" -ItemType director -name "SY$currentyear"
        New-Item -path "$studir$name\" -ItemType director -name "Evaluations"
    }
    If (Test-Path "$studir$name\SY$currentyear" -PathType Container)
    {Write-host "$studir$name\SY$currentyear" already exists}
    Else{
    New-Item -path "$studir$name\" -ItemType director -name "SY$currentyear"
     }
    
    
}

Foreach ($504stu in $504stus){
    $fname = (Removespecials($504stu."Student First Name"))
    $lname = (Removespecials($504stu."Student Last Name"))
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
    {Write-host "$studir$name already exists"}
    elseif (Get-ChildItem -Path "G:\Shared drives\PS-504" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\PS-504" -Recurse -Include "$name"| Move-Item -Destination $studir        
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\IS-504" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\IS-504" -Recurse -Include "$name"| Move-Item -Destination $studir        
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\MS-504" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\MS-504" -Recurse -Include "$name"| Move-Item -Destination $studir        
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\JH-504" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\JH-504" -Recurse -Include "$name"| Move-Item -Destination $studir        
    }
    elseif (Get-ChildItem -Path "G:\Shared drives\HS-504" -Recurse -Include "$name") {
        Write-Host "Moving $name to $studir"
        Get-ChildItem -Path "G:\Shared drives\HS-504" -Recurse -Include "$name"| Move-Item -Destination $studir        
    }
    else{
        New-Item -path $studir -ItemType directory -name $name
        New-Item -path "$studir$name\" -ItemType director -name "Old Documents"
        New-Item -path "$studir$name\" -ItemType director -name "SY$currentyear"
        New-Item -path "$studir$name\" -ItemType director -name "Evaluations"
    }
    If (Test-Path "$studir$name\SY$currentyear" -PathType Container)
    {Write-host "$studir$name\SY$currentyear" already exists}
    Else{
    New-Item -path "$studir$name\" -ItemType director -name "SY$currentyear"
     }
}