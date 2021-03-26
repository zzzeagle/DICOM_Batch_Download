# DICOM_Batch_Download
A script that downloads a list of studies from PACS. This version can now run on Windows, Mac and Linux.

Requirements:

-Powershell 7 or greater (https://docs.microsoft.com/en-us/powershell/scripting/whats-new/what-s-new-in-powershell-70?view=powershell-7.1)

-DCMTK 3.6 or greater (https://dicom.offis.de/dcmtk.php.en)

To use this script, call the script followed by the following 4 parameters.
1. Computer variables file. The location of a ps1 file that has the following variables defined:
   - $DCMTK = The path to the DCMTK binaries
   - $myIP = your computers IP
   - $myAE = your computers AE title
   - $myPort = your computers port
   - $PACSIP = the PACS IP
   - $PACSAE = the PACS AE
   - $PACSPORT = the PACS Port
2. The path to your CSV containing the accession numbers
3. The project folder path
4. The project name
