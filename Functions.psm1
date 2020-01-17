<#

	.NOTES
		Version:		1.4
#>

function log-write {


    Param([string]$Logstring)

    Add-Content $logfile -Value "$(get-date -f MM/dd/yyyy_hh:mm:ss): $Logstring"

}

Function Run-Process{

    <#

	.NOTES
		Changelog:
            1.2         Added Ran Command Line to Output Object
    		1.1         Initial Relase
    #>

    Param(
        [Parameter(Mandatory=$true)]
        [string] $FilePath,

        [string] $Arguments
    )
    Try{
    $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
    $ProcessInfo.FileName = $FilePath
    $ProcessInfo.RedirectStandardError = $true
    $ProcessInfo.RedirectStandardOutput = $true
    $ProcessInfo.UseShellExecute= $false
    $ProcessInfo.Arguments = $Arguments
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $ProcessInfo
    $process.Start() | out-null
    
    

      $Output = New-Object -TypeName psobject -Property @{'Data'= $process.StandardOutput.ReadToEnd(); 'CMDLine'= "$FilePath $arguments"; 'ProcError' = $process.StandardError.ReadToEnd(); 'ExitCode'= $process.ExitCode }
     $process.WaitForExit()
    }



    Catch{
    $Output = New-Object -TypeName psobject -Property @{'Data'= $null; 'CMDLine'= "$FilePath $arguments"; 'ProcError' = $_.Exception.toString(); 'ExitCode'= 4123 }

    Return $Output
    }    
     
    return $Output
}

Function Create-Log{

<#

	.NOTES
		Version:		1.1
#>

	Param(
			[Parameter(Mandatory=$true)]
			[String] $path,
			[Parameter(Mandatory=$true)]
			[string] $Name
			
		)
	$fullPath = "$path\$Name"
	if(Test-Path $fullPath){

		Clear-Content -Path $

	}
	
}

Function Create-shortcut{
<#

	.NOTES
		Version:		1.3
#>
    param(
        [parameter(Mandatory=$true)]
        [string]$ShortcutPath, #Path where to place Shortcut
        [parameter(Mandatory=$true)]
        [string]$Name, #Name to be used for Shortcut
        [parameter(Mandatory=$true)]
        [string]$Target, #Where the Shortcut will point to
        [bool]$UseURL, #Will this shortcut be used to point to a URL  
        [string]$IconPath #path where an icon is stored
    )
    $ext = "lnk"
    if($UseURL){
        $ext="url"
    }

    $WshShell = New-Object -ComObject WScript.Shell
    $shortcut = $WshShell.CreateShortcut("$ShortcutPath\$Name.$ext")
    $shortcut.TargetPath = $Target
    if($IconPath){
    $shortcut.Iconlocation = "$IconPath, 0"
    }
    $shortcut.Save()

    return $shortcut
}


function Get-FileName($initialDirectory)
{  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “All files (*.*)| *.*”
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

function Get-DepartmentUsers {
    param (

        [string]$costCenterNumber
    )
    try{
    get-aduser -Filter * -Properties hnnyorgteam,hnnyorggroup | where {$_.hnnyorgteam -like "$costCenterNumber*"} |Select-Object -Property Name,SAmAccountNAme | export-csv c:\temp\$($costCenterNumber).csv }
    Catch{
        Write-Output $_.Exception


    }

    Write-Output "$costCenterNumber.csv in C:\Temp"

}



Function Test-AdminCred{
<#
    .SYNOPSIS
        TEst User Credentials for ADMIN Rigts
    .NOTES
        VERSION:      1.0.0
#>

    param(
        [string]$ComputerName #for Remote Machines
    )

    if(!($ComputerName)){

        $ComputerName = $env:COMPUTERNAME
    }


    try{
        
        $silent = Test-Path \\$ComputerName\admin$ -ErrorAction Stop

    }
    catch{
    
        $AdminRightsTest=$_
        if ($AdminRightsTest.Exception -like "*Access is denied*" -or $a -eq $false) {

            Write-Warning "Insufficient privileges to run this Device Action"

            Write-Output "Typing Y or Yes will prompt for new credentials`nTyping N or No will terminate this script`n`nThe username format should be as follows:`nDOMAIN\USERNAME"

            $Input=Read-Host "Prompt for new credentials? (Yes/No)"

            if($Input -eq "Yes" -or $Input -eq "Y"){
                while($Input -ne "Yes" -and $Input -ne "Y" -and $Input -ne "No" -and $Input -ne "N") {
                
                    Write-Host "`n`n"

                    Write-Warning "Invalid option selected"

                    Write-Output "Typing Y or Yes will prompt for new credentials`nTyping N or No will terminate this script`n`nThe username format should be as follows:`nDOMAIN\USERNAME"

                    $Input=Read-Host "Prompt for new credentials? (Yes/No)"

                }



            }while($Input -ne "Yes" -and $Input -ne "Y" -and $Input -ne "No" -and $Input -ne "N") {
                
                Write-Host "`n`n"

                Write-Warning "Invalid option selected"

                Write-Output "Typing Y or Yes will prompt for new credentials`nTyping N or No will terminate this script`n`nThe username format should be as follows:`nDOMAIN\USERNAME"

                $Input=Read-Host "Prompt for new credentials? (Yes/No)"

            }

            if ($Input -eq "No" -or $Input -eq "N") {
            
                Write-Output "Operation cancelled by user`nTerminating Device Action..."

                Start-Sleep -Seconds 2

                Exit

            }

            if ($Input -eq "Yes" -or $Input -eq "Y") {

                $AdminCred = try {Get-Credential -Message "Please type in your administrative credentials in the format DOMAIN\USERNAME" -ErrorAction SilentlyContinue} catch {}

                if ($AdminCred -eq $null) {

                    Write-Output "Operation cancelled by user`nTerminating Device Action..."

                    Start-Sleep -Seconds 2

                    Exit

                }

                try {

                    $b=New-PSDrive -Name Temp -PSProvider FileSystem -Root "\\$Computer\admin$" -Credential $AdminCred -ErrorAction Stop

                }

                catch {

                    $AdminRightsTest2=$_

                    if ($AdminRightsTest2.Exception -like "*Access is denied*") {

                        Write-Warning "The credentials provided are still insufficient.`nPlease try running Script: ""Add Standard Admins to Machine""`nTerminating Device Action..."

                        Pause-Script 'J' $modifier 'Ctrl + J' $true

                        Exit

                    }

                    else {

                        $AdminRightsTest2

                        Write-Warning "There was an unexpected error in checking for administrative privileges`nMake sure the device is online and try again"

                        Pause-Script 'J' $modifier 'Ctrl + J' $true

                        Exit

                    }

                }
           



    }



}
    }
}