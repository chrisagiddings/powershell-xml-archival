# Written by Chris Giddings
# Updated May 27th 2014
#
# This script takes All XML files in specific directories from the past 24 hours,
# creates a zip file to hold them and inserts the files one by one into the zip,
# prior to deleting the originals.
###############################################################################
#                       DECLARE GLOBAL VARIABLES
###############################################################################
$global:DataMigrationPath = "C:\DataMigration";
#
# DERIVED GLOBAL VARIABLES
$global:DataPath = "$DataMigrationPath\Data";
#
# SUB-DERIVATIVE VARIABLES
$global:FamilyTree = "$DataPath\FamilyTree";
$global:GDSN = "$DataPath\GDSN";
$global:GDSN_CINXML = "$GDSN\CINXML";
$global:BackInTime = "-90";
#
# DROP ZONE PATH VARIABLES
$CINXML_DROP = "$GDSN_CINXML";
$FAMILY_TREE_DROP = "$FamilyTree";

###############################################################################
#   VERIFY IF GIVEN DIRECTORY EXISTS
###############################################################################
Function CheckDirectoryStructure($StructureItem)
{
    if (!(Test-Path -path $StructureItem))
    {
        New-Item $StructureItem -type directory;
        sleep 1;
        Write-Output "New directory created: $StructureItem.";
    }
    Else
    {
        Write-Output "$StructureItem already exists.";
    }
}
###############################################################################
#   VERIFY REQUIRED DIRECTORY STRUCTURES EXIST
###############################################################################
Function CheckBaseStructure
{
    CheckDirectoryStructure "$DataMigrationPath";
    CheckDirectoryStructure "$DataPath";
    CheckDirectoryStructure "$FamilyTree";
    CheckDirectoryStructure "$GDSN";
    CheckDirectoryStructure "$GDSN_CINXML";
}
###############################################################################
#   CHECK FOR DATED DIRECTORIES IN STRUCTURE
###############################################################################
$TodaysDate = Get-Date -format yyyMMdd;
$ThisYear = Get-Date -format yyy;
$ThisMonth = Get-Date -format MM;
#Write-Output "$TodaysDate";

$YearInFamilyTree ="$FamilyTree\$ThisYear";
$MonthInFamilyTree = "$YearInFamilyTree\$ThisMonth";

$YearInCINXML = "$GDSN_CINXML\$ThisYear";
$MonthInCINXML = "$YearInCINXML\$ThisMonth";

Function CheckOrCreateDatedItem($DatedItem)
{
    CheckDirectoryStructure "$DatedItem";
}
###############################################################################
#   VERIFY REQUIRED DIRECTORY STRUCTURES EXIST
###############################################################################
Function CheckCustomStructure
{
    CheckOrCreateDatedItem "$YearInFamilyTree";
    CheckOrCreateDatedItem "$MonthInFamilyTree";

    CheckOrCreateDatedItem "$YearInCINXML";
    CheckOrCreateDatedItem "$MonthInCINXML";
}
###############################################################################
#   CREATE NEW EMPTY ZIP FILE NAMED FOR DATE [NAME]_<DATE>.zip
###############################################################################
Function CreateNewDatedZip($zipName)
{
    set-content $zipName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18));
    (dir $zipName).IsReadOnly = $false;

    Start-Sleep -milliseconds 500;
}
###############################################################################
#   INSERT EXISTING DATED XML FILE INTO EXISTING DATED ZIP FILE
#
#   ITTERATE OVER ARRAY TO DETERMINE FILE & INFO
###############################################################################
Function InsertDatedFileIntoDatedZip($fileName, $zipName)
{
    if (!(Test-Path($zipName)))
    {
        Write-Output "Critical Error: The archive $zipFileName does not exist.";
        Exit 1
    }
    
    $shellApplication = New-Object -com shell.application;
    $zipPackage = $shellApplication.NameSpace($zipName);
    
    # Syntax is odd. Basically, copy (file) into the preceeding zip.
    $zipPackage.CopyHere($fileName);

    # Take a nap in case PoweShell hates sub-shells.
    Start-Sleep -milliseconds 500;
}
###############################################################################
#   DELETE A FILE OR FOLDER
###############################################################################
function DeleteFileOrFolder($PathToItem)
{ 
    if (Test-Path $PathToItem)
    {
        Remove-Item ($PathToItem) -Force -Recurse;
    }
}
###############################################################################
#   ITERATIVELY GO BACK AND PROCESS FILES UP TO $BackInTime DAYS OLD
###############################################################################
Function ProcessAndArchiveFiles
{
    Write-Output "Looking as far back as $BackInTime days.";

    # UNTIL WE REACH BACK AS FAR AS WE'VE CONFIGURED
    while ($CurrentDaysBack -ge $BackInTime)
    {
        $DateToMatch=(Get-Date).AddDays($CurrentDaysBack).ToString("yyyMMdd");
###############################################################################
#   DETECT & PROCESS CINXML FILES
###############################################################################
#   CINXML Convention       CINyyyyMMddHHmmssSSS.xml
###############################################################################
        # foreach($file in Get-ChildItem "$CINXML_DROP\*.xml")
        # {
        #     $zipName = "$MonthInCINXML\CIN_$DateToMatch.zip";

        #     if ( $file.name -like "*$DateToMatch*.xml" )
        #     {
        #         Write-Output "";
        #         Write-Output "Matching file found: $file";

        #         if ( Test-Path -Path "$zipName" )
        #         {
        #             Write-Output "CINXML archive, $zipName already exists. Matching files will be inserted.";

        #             Write-Output "Archiving $file into $zipName";
        #             InsertDatedFileIntoDatedZip "$file" "$zipName";

        #             DeleteFileOrFolder "$file";
        #         }
        #         else
        #         {
        #             Write-Output "Creating new CINXML archive: $zipName"
        #             CreateNewDatedZip $zipName;

        #             Write-Output "Archiving $file into $zipName";
        #             InsertDatedFileIntoDatedZip "$file" "$zipName";

        #             # DeleteFileOrFolder "$file";
        #         }
        #     }
        # }
###############################################################################
#   DETECT & PROCESS FAMILYTREE FILES
###############################################################################
#   FamilyTree Convention:  SyncFamilyTreeyyyyMMddHHmmssSSS.xml
###############################################################################
        foreach($file in Get-ChildItem "$FAMILY_TREE_DROP\*.xml")
        {
            $zipName = "$MonthInFamilyTree\FamilyTree_$DateToMatch.zip"

            if ($file.name -like "*$DateToMatch*.xml")
            {
                Write-Output "";
                Write-Output "Matching file found: $file";

                if (Test-Path -Path "$zipName")
                {
                    Write-Output "FamilyTree archive, $zipName already exists. Matching files will be inserted.";

                    Write-Output "Archiving $file into $zipName";
                    InsertDatedFileIntoDatedZip "$file" "$zipName";

                    # DeleteFileOrFolder "$file";
                }
                else
                {
                    Write-Output "Creating new FamilyTree archive: $zipName"
                    CreateNewDatedZip $zipName;

                    Write-Output "Archiving $file into $zipName";
                    InsertDatedFileIntoDatedZip "$file" "$zipName";

                    # DeleteFileOrFolder "$file";
                }
            }
        }
        $CurrentDaysBack = $CurrentDaysBack - 1;
    } 
}
###############################################################################
#
#
#   HERE IS WHERE WE ACTUALLY INVOKE FUNCTIONALITY
#
#
###############################################################################
CheckBaseStructure
CheckCustomStructure

Write-Output "";
Write-Output "BEGINNING ARCHIVAL PROCESS"

$CurrentDaysBack = -1;
ProcessAndArchiveFiles

Write-Output "";
Write-Output "ARCHIVAL PROCESS COMPLETED"

Exit 0
