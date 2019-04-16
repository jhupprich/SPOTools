<#
  .SYNOPSIS
  Reports folder names and relative URLs for all folders in a document library

  .DESCRIPTION
   Used to enumerate all folders in a selected document library
   Saves workbook is same directory as script. This script requires the ImportExcel module for PowerShell.
   https://github.com/dfinke/ImportExcel
   https://www.powershellgallery.com/packages/ImportExcel/4.0.11

   SharePoint Online Client Components SDK is also needed on host computer:
   https://www.microsoft.com/en-us/download/details.aspx?id=42038

  .EXAMPLE
  .\Iterate-SpoFolders.ps1 

  .OUTPUTS
  'SPO Folders.xlsx' or 'SPO Folders.csv' in root directory

  .NOTES
#>

## load CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$global:arr01 = @()

## functions
function Iterate-SpoFiles ([Microsoft.SharePoint.Client.Folder]$folder) 
    {
  
        $context.Load($folder.Folders)
        $context.ExecuteQuery()

        foreach($fold in $folder.Folders  | Where-Object {$_.Name -ne "Forms"})
            {
                Write-Host -ForegroundColor Magenta $fold.Name
                $fold.ServerRelativeUrl
                $attr = @(New-Object System.Object)
                $attr | Add-Member -Type NoteProperty -Name 'FolderName' -Value $fold.Name
                $attr | Add-Member -Type NoteProperty -Name 'RelativeURL' -Value $fold.ServerRelativeUrl
                $global:arr01 += $attr

                Iterate-SpoFiles -Folder $fold



            }

        
    }


## create context - change these variables!
$admin = 'admin@whatever.com'         #change
$pass = ConvertTo-SecureString 'SomePassword' -AsPlainText -Force      #change
$cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin,$pass)
$siteUrl='https://tenant.sharepoint.com/sites/companySite'    #change
$docLib = 'Documents'   #change

$context = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$context.Credentials = $cred
$library = $context.web.Lists.GetByTitle($docLib)
$context.Load($library)
$context.Load($library.RootFolder)
$context.ExecuteQuery()

Iterate-SpoFiles -Folder $library.RootFolder


## if you aren't installing the ImportExcel module - use the second line to eport to CSV
$arr01 | Export-Excel -AutoSize 'SPO Folders.xlsx'  #change if needed
#$arr01 | Export-Csv  'SPO Folders.csv'  -NoTypeInformation  
