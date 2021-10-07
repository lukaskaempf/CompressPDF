#============================================================================================
#
# LANG			: Powershell 
# NAME			: CompressPDF
# AUTHOR		: Lukas Kämpf
# DATE  		: 26.08.2021
# Description 	: 
# Version       : 1.0
#
#============================================================================================

#============================================================================================
# Variable Declarations
#============================================================================================

$errors = @{}

#============================================================================================
# FUNCTIONS
#============================================================================================

$ErrorActionPreference = "Stop"

function compress{
    param(
    [string]$location,
    [string]$tempoutput,
    [string]$currentfolder)
    try{
        gswin64 -sDEVICE=pdfwrite -dBATCH -dNOPAUSE -dNOSAFER -q -dPDFSETTINGS=/ebook -sOutputFile="$tempoutput" $location
        Start-Sleep 1
        try{
            Move-Item -Path $tempoutput $location -Force
        }
        catch{
            Write-Host("error")
            $errors.add($currentfolder, $location)
        }
        
        return "Success"
        }
    catch{
        return "Failure"
        }
}


#============================================================================================
# Main Body
#============================================================================================

$folders = Get-ChildItem 
$counter = 0
foreach ($folder in $folders){
    $path = Get-Location
    $folderpath = "$path\$folder"
    $files = Get-ChildItem $folderpath
    foreach ($file in $files){
        
        $currentfile= "$folderpath\$file"
        $temp = "$folderpath\temp.pdf"
        

        $res = compress -location $currentfile -tempoutput $temp -currentfolder $folderpath
        $errors.Count
        if ($res -eq "Failure"){
            Write-Host("$res ----> $currentfile")
        }
    $counter ++
    [int] $currentProzent = (($counter/$folders.Count) * 100)
    Write-Progress -Activity "Processing $folder" -CurrentOperation("$currentProzent %") 
    }
    
}
repair
repairItemsWithOneFile

function repair{
    $todelete = @{}
    foreach ($error1 in $errors.GetEnumerator()){
        Write-Host("ttttt $($error1.Name)")
        $files = Get-ChildItem $error1.Name
        if ($files.Count -eq 2){
            $tempoutput = "$($error1.Name)\$($files[1])"
            $location = "$($error1.Name)\$($files[0])"
            try{
                Move-Item -Path $tempoutput $location -Force
                
                $todelete.add($error1.Name, "test $($error1.Name)")
               
            }
            catch{Write-Host("fehler")}
        }
    }
    
    foreach ($delete in $todelete.GetEnumerator()){
        Write-Host($delete.Name)
        $errors.Remove("$($delete.Name)")
    }
    
}


function repairItemsWithOneFile{
    $todelete = @{}
    foreach ($error1 in $errors.GetEnumerator()){
        $files = Get-ChildItem $error1.Name
        if ($files.Count -eq 1){
            Write-Host("ttttt $($error1.Name)")
            #immer temp.pdf
            $split = $($error1.Name).Split("\")
            $superiorfolder = $($split[-2]).SubString(1,2)
            $newItemName = "O1_$($superiorfolder+$split[-1])_1"
            try{
          
               
               Rename-Item -Path "$($error1.Name)\temp.pdf" -NewName "$($error1.Name)\$newItemName.pdf"
                
               $todelete.add($error1.Name, "test $($error1.Name)")
               
            }
            catch{Write-Host("fehler")}
        }
    }
    
    foreach ($delete in $todelete.GetEnumerator()){
        Write-Host($delete.Name)
        $errors.Remove("$($delete.Name)")
    }
    
}


