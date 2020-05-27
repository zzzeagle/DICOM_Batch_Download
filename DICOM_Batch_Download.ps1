###########################  Read Me #######################################
### You can use this tool to download cases from a PACS server
### You will need to have DCMTK installed on your computer
### Your computer will need to have query and retrieve access from PACS
### Modify the parameters below to fit your project
#############################################################################


##############################
####    Parameter Files   ####
##############################
. ".\PACS_info.ps1"
. ".\Project_info.ps1"

##############################
#### Script - Do not edit ####
##############################
function Try-Command($Command, $Parameters, $queryfile) {
     $i = 0
     Write-Host $command $parameters
     do{
        $output = & $Command $Parameters $queryfile
        #Start-Sleep -Seconds 15
        $i++
        }until(($output | Select-String -Pattern "0x0000: Success" -Quiet) -or ($i -eq 20))
        
        if(($output | Select-String -Pattern "0x0000: Success" -Quiet)){
            if($queryfile){Remove-Item $queryfile}
            }
        else{
            $error = ""
            $error = "Error on command: "+$command+$parameters+$queryfile
            Add-Content $errorLog -Value $error
            return "Error"
        }
 }

function Find-DICOM($Query){
    $findexe = "C:\Program Files\dcmtk\bin\findscu.exe" 
    $findparam = @(
        '-S',
        '-d',
        '-aet', $MYAE,
        '-aec', $PACSAE,
        $PACSIP,
        $PACSPORT)
    $findParameters = $findparam + $Query
    Try-Command $findexe $findParameters
}

function Move-DICOM($Query, $QueryFile){
    $moveexe = "C:\Program Files\dcmtk\bin\movescu.exe" 
    $param = @(
        '-S',
        '-d',
        '+xa',
        '-aet', $MYAE,
        '-aec', $PACSAE,
        $PACSIP,
        $PACSPORT,
        '--port', $myPort)
    $moveParameters = $param + $Query
    Try-Command $moveexe $moveParameters $QueryFile
}

function Create-UniqueDirectory($Path, $Name){
        $i = ''
        $name = $name  -replace '[^A-Za-z0-9_.]', ''
        $fullPath = $path +"\"+$name
        while(Test-Path ($fullPath+$i)){
            $i = $i+1 -as [int]
            }
        $folder = New-Item -ItemType directory -Path ($fullPath+$i)
        return $folder.FullName
}

$ProjectPath = $ProjectFolder +"\"+ $ProjectName
$errorLog = $projectPath+ "errors.txt"
#Check if path exists, if it does, error, if not create it.
If(Test-Path $ProjectPath -PathType Container){
    Throw "Project already exists. Choose a different name"
    }Else{
    New-Item -ItemType directory -Path $ProjectPath
    }
If(-not(Test-Path $AccList -PathType leaf)){
    Throw "File does not exist, check path"
    }
$(
#Create a folder for each accession and get the study level query file
$accessions = Import-Csv $AccList -Header "Accession"

#For each accession in the input list
ForEach ($accession in $accessions){
    
    #Create Study Folder
    $StudyFolder = Create-UniqueDirectory $ProjectPath $accession.Accession
    $FindResponseFolder = Create-UniqueDirectory $StudyFolder "responses"
    
    #Set Find Parameters
    $findParam = @(
            '-k', 'QueryRetrieveLevel=STUDY',
            '-k', -join ('AccessionNumber=', $accession.Accession),
            '-X',
            '-od', $FindResponseFolder
        )
    #Execute Find Command
    Find-DICOM $findParam


    $rsps = get-childitem $FindResponseFolder
    #For each response from the Stufy level find.
    #A study can have more than one StudyUID
    ForEach ($rsp in $rsps){
             $seriesDir = $StudyFolder
             #New-Item -ItemType directory -Path $seriesDir
             $findParam = @(
                $rsp.FullName,
                '-k', 'QueryRetrieveLevel=SERIES',
                '-k', 'SeriesInstanceUID',
                '-k', 'SeriesDescription',
                '-X',
                '-od', $studyFolder)

             Find-DICOM $findParam
    

            #For each series retrieve the images into a folder
            $seriesQF = Get-ChildItem -path $studyfolder\* -Include *.dcm
            ForEach ($series in $seriesQF){
                $seriesJSON = & "$DCMTK\dcm2json.exe" $series.FullName -fc
                $seriesJSON = ConvertFrom-Json $seriesJSON
                $seriesName = $seriesJSON."0008103E".value

                if ($seriesName -eq $NULL)
                    {$seriesOutput = $seriesJSON."0020000E".value}
                else
                    {$seriesOutput = $seriesName}
        
                #$seriesOutput = $seriesOutput.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'

                $outputDir = Create-UniqueDirectory $series.DirectoryName $seriesOutput

                $moveParam = @(
                    $studyQF,
                    '-k', 'QueryRetrieveLevel=SERIES',
                    $series.FullName,
                    '-od', $outputDir)

                Move-DICOM $moveParam
                Remove-Item $series
                #Remove-Item $series
    }

}
    Remove-Item $FindResponseFolder -Recurse

    if(Test-Path $errorLog){
        Write-Host "Job finished with errors. Please review the error log" -BackgroundColor Red
    }
}) *>&1 > $ProjectPath'log.txt'