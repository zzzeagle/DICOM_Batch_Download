##############################
### PACS Downloader        ###
###                        ###
### Author: Zachary Eagle  ###
### Tested with:           ###
###   PowerShell 7.1       ###
###   DCMTK 3.6.6          ###
##############################

###########################  Read Me #########################################################
### You can use this tool to download cases from a PACS server
### You will need to have DCMTK 3.6 or greater installed on your computer
### Your computer will need to have query and retrieve access from PACS
### To run this script, call it from the command line with the following four parameters
### 1) Computer variables file. The location of a ps1 file that has the following variables defined:
###    a) $DCMTK = The path to the DCMTK binaries
###    b) $myIP = your computers IP
###    c) $myAE = your computers AE title
###    d) $myPort = your computers port
###    e) $PACSIP = the PACS IP
###    f) $PACSAE = the PACS AE
###    g) $PACSPORT = the PACS Port
### 2) The path to your CSV containing the accession numbers
### 3) The project folder path
### 4) The project name
###############################################################################################


##############################
######## Parameters  #########
##############################
param ($computervariables, $AccList, $ProjectPath, $ProjectName)
### Import variables
. $computervariables


#DCMTK Binaries
$findbin = Join-Path $DCMTK "findscu" 
$movebin = Join-Path $DCMTK "movescu"
$dcm2jsonbin = Join-Path $DCMTK "dcm2json"


#Find the DICOM value given the DCMTK output and the tag number
function Find-DICOMTagValue($output, $tag){
    $matches = $null
    $regex = "\("+$tag+"\).{0,20}\[(.*?)\]"
    if ([string]$output -match $regex){
        return $Matches[1].Trim()
    }else {
        return $null
    }
}

#Find the series in the Study
function Find-Series($accession, $Query){
    $findparam = @(
        '-S',
        #'-v',
        '--timeout',  '60',
        '--dimse-timeout', '60'
        '-aet', $MYAE,
        '-aec', $PACSAE,
        $PACSIP,
        $PACSPORT)
    $findParameters = $findparam + $Query
    
    $findcommand = "'"+$findbin+"'"+$findParameters + " 2>&1"
    $output = Invoke-Expression "& $findcommand"
}

###Find the study
function Find-Study($accession, $Query){
    $findparam = @(
        '-S',
        #'-v',
        '-k', 'NumberOfStudyRelatedInstances',
        '--timeout',  '60',
        '--dimse-timeout', '60',
        '-aet', $MYAE,
        '-aec', $PACSAE,
        $PACSIP,
        $PACSPORT)
    $findParameters = $findparam + $Query
    $findcommand = "'" +$findbin+"' " + $findParameters + " 2>&1"
    $output = Invoke-Expression "& $findcommand"

    $SRI = Find-DICOMTagValue $output "0020,1208"
    $IA = Find-DICOMTagValue $output "0008,0056"

    return $IA, $SRI
}

###Move the study
function Move-DICOM($accession, $Query, $QueryFile){
    $param = @(
        '-S',
        '-d',
        '+xa',
        '--timeout',  '60',
        '--dimse-timeout', '60'
        '-aet', $MYAE,
        '-aec', $PACSAE,
        $PACSIP,
        $PACSPORT,
        '--port', $myPort)
    $moveParameters = $param + $Query + $QueryFile

    $output = ""
    $movecommand = "'" + $movebin +"' "+ $moveParameters + " 2>&1"
    $moveoutput = Invoke-Expression "& $movecommand"
    $statusRegex = "DIMSE Status(?:[^:]*\:){2}(.*)D:"
    $errorRegex = "[E|F]:(.*)"
    $moveoutput = [string]$moveoutput

    if($moveoutput -match $statusRegex){

        return $Matches[1]
    }elseif($moveoutput -match $errorRegex){
        Write-Host "ERROR"
        return $Matches[1].Trim()
    }else{
        return $null
    }  
}
function New-UniqueDirectory($Path, $Name){
        $i = ''
        $name = $name  -replace '[^A-Za-z0-9_.]', ''
        $fullPath = $Path +"\"+$name
        while(Test-Path ($fullPath+$i)){
            $i = $i+1 -as [int]
            }
        $folder = New-Item -ItemType directory -Path ($fullPath+$i)
        return $folder.FullName
}

function Write-Log($accession, $StudyStatus, $StudyInstances, $InstancesInFolder, $MoveMessage){
    $logEntry = [PSCustomObject]@{
        Accession = $accession
        StudyStatus = $StudyStatus
        StudyInstances = $StudyInstances
        InstancesInFolder = $InstancesInFolder
        MoveMessage = $MoveMessage
    }

    return $logEntry
    }

function New-ProjectFolder($ProjectPath, $ProjectName){
    $NewPath = Join-Path $ProjectPath $ProjectName
    #Check if path exists, if it does, error, if not create it.
    If(Test-Path $NewPath -PathType Container){
        Throw "Project already exists. Choose a different name"
        }Else{
        $folder = New-Item -ItemType directory -Path $NewPath
        return $folder
        }
}

function Clear-ScriptVariables(){
    Clear-variable -name log,
                    missedExams, 
                    missedExamsObject, 
                    logObject, 
                    accessions, 
                    computervariables, 
                    AccList, 
                    ProjectPath, 
                    ProjectName, 
                    findbin,
                    movebin, 
                    dcm2jsonbin,
                    projectfolder
}



If(-not(Test-Path $AccList -PathType leaf)){
    Clear-ScriptVariables
    Throw "File does not exist, check path"
}else{
    $accessions = Import-Csv $AccList -Header "Accession"
}

If($ProjectName -eq $null){
    Clear-ScriptVariables
    Throw "Missing parameter. Please enter four parameters. 1)Computer Variables 2)Accession List 3)Project Path 4)Project Name"
}

#Create Project Folder
$ProjectFolder = New-ProjectFolder $ProjectPath $ProjectName

#Prepare log files
$log = Join-Path $ProjectFolder "log.csv"
$missedExams = Join-Path $ProjectFolder "missedexams.csv"
$missedExamsObject = @()
$logObject = @()

#For each accession in the input list
ForEach ($accession in $accessions){
    write-host "Study Start: " -NoNewline
    get-date -Format "HH:mm:ss.fff" | Write-Host
    #Create Study Folder
    $StudyFolder = New-UniqueDirectory $ProjectFolder.FullName $accession.Accession
    $FindResponseFolder = New-UniqueDirectory $StudyFolder "responses"
    
    #Set Find Parameters
    $findParam = @(
            '-k', 'QueryRetrieveLevel=STUDY',
            '-k', -join ('AccessionNumber=', $accession.Accession),
            '-X',
            '+sr',
            '-od', $FindResponseFolder
        )
    #Execute Find Command, Return the Study Online Status and number of instances
    $StudyStatus, $StudyInstances = Find-Study $accession.Accession $findParam
    write-host "Find Study: " -NoNewline
     get-date -Format "HH:mm:ss.fff" | Write-Host

    $rsps = get-childitem $FindResponseFolder
    #For each response from the Study level find.
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

             Find-Series $accession.Accession $findParam
             write-host "Find Series: " -NoNewline
             get-date -Format "HH:mm:ss.fff" | Write-Host
    

            #For each series retrieve the images into a folder
            $seriesQF = Get-ChildItem -path $studyfolder\* -Include *.dcm
            ForEach ($series in $seriesQF){
                $seriesJSON = & $dcm2jsonbin $series.FullName -fc
                $seriesJSON = ConvertFrom-Json $seriesJSON
                $seriesName = $seriesJSON."0008103E".value

                if ($seriesName -eq $NULL)
                    {$seriesOutput = $seriesJSON."0020000E".value}
                else
                    {$seriesOutput = $seriesName}
        
                $outputDir = New-UniqueDirectory $series.DirectoryName $seriesOutput

                $moveParam = @(
                    $studyQF,
                    '-k', 'QueryRetrieveLevel=SERIES',
                    $series.FullName,
                    '-od', $outputDir)

                $MoveStatus = Move-DICOM $accession.Accession $moveParam
                write-host "Move Study: " -NoNewline
                get-date -Format "HH:mm:ss.fff" | Write-Host
                Remove-Item $series
            }
    }
    Remove-Item $FindResponseFolder -Recurse -Force

    #Count files downloaded
    $FilesDownloaded = (Get-ChildItem $StudyFolder -Recurse -File | Measure-Object).Count
    $StudyInstances = $StudyInstances.trim()
    
    #Uncomment if you want to delete incomplete studies
    #If the files downloaded doesn't match the number PACS reported, delete the folder and log the exam to the missing exams log
    #if(($FilesDownloaded -ne $StudyInstances) -or ($StudyStatus -ne "ONLINE")){
    #   Remove-Item $StudyFolder -Recurse
    #    $missedexam = [PSCustomObject]@{
    #        Accession = $accession.Accession
    #     }
    #     $missedExamsObject += $missedexam
    #    }

    #Create log entry
    $studyLog = Write-Log $accession.Accession $StudyStatus $StudyInstances $FilesDownloaded $MoveStatus
    Write-Host $studyLog
    $logObject += $studyLog
    write-host "Study Finish: " -NoNewline
    get-date -Format "HH:mm:ss.fff" | Write-Host
}

#Save log file
$logObject | Export-CSV $log -NoTypeInformation

#Save missed exams list
if($missedExamsObject.length -gt 0){
    $missedExamsObject | Export-CSV $missedExams -NoTypeInformation
}


Clear-ScriptVariables
