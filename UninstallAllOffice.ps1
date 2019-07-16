
Param(
       
       [Parameter(Mandatory=$true)]
       [string]$DependencyDir
  

)


Function Get-UninstallString{
    
    Param(
        [Parameter(Mandatory=$true)]
        [string]$DisplayName
    )


    $get = Get-ChildItem REGISTRY::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Select-Object -Property Name 
    $get += Get-ChildItem REGISTRY::HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Select-Object -Property Name 
    [array]$end = $null

    foreach($x in $get){

        If($set = Get-ItemProperty "REGISTRY::$($x.Name)" |Where-Object {$_.DisplayName -like "*$($DisplayName)*"}) {
            
            $set | Add-Member Path $x.Name
            $end += $set
            
  
        }

    
    }
    
    return $end
}

[array]$officeVersions = @("Access 97", "edition 2003", "2010", "2013","2016")
[string]$logfile = "C:\ProgramData\HNInstallerLogs\RemovalofOffice_$(Get-Date -Format MMddyy_HHmmss).Log"

Write-Output "Check $logfile for Verbose Logs"

foreach($version in $officeVersions){

    $isInstalled = $null
    
    switch ($version) {
        "edition 2003"{$isInstalled = Get-UninstallString "Microsoft Office Professional $($version)" ;break}
        "Access 97" {$isInstalled = Get-UninstallString $version; break}
        Default {$isInstalled = Get-UninstallString "Microsoft Office Professional Plus $($version)"}
    }
 
    if($isInstalled){

        $version |Out-File -FilePath $logfile -Append

        switch($version){
            
            "Access 97"{ #custom thin app of Access 97 used for legacy DB's
                
                Write-Output "Found Acess 97 starting uninstall" |Out-File $logfile -Append

                Start-Process -FilePath Msiexec.exe -ArgumentList "/x{F6A4850C-42B7-454E-B3CA-3E4BC3B6B337} /qn" -PassThru -NoNewWindow -Wait| Out-File $logfile -Append
    
                Break
            }

            "Edition 2003"{


                Write-Output "Found Office 2003 starting uninstall" |Out-File $logfile -Append

                Start-Process -FilePath Msiexec.exe -ArgumentList "/x{90110409-6000-11D3-8CFE-0150048383C9} /qn" -PassThru -NoNewWindow -Wait| Out-File $logfile -Append

                break
            
            
            }

            "2010"{
                
                if(Get-UninstallString "Microsoft Project Standard 2010"){

                    Write-Output "Found Project 2010 starting uninstall" |Out-File $logfile -Append

                    Try{Copy-Item -Path $DependencyDir\Uninstallproject.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\UninstallProject.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                                                                      
                    Try{Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\setup.exe' -ArgumentList "/uninstall PrjStd /Config uninstallproject.xml" -NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

                }

                if(Get-UninstallString "Microsoft Visio Premium 2010"){
                    
                    Write-Output "Found Visio 2010 starting uninstall" |Out-File $logfile -Append

                    Try{Copy-Item -Path $DependencyDir\uninstallvisio.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\uninstallvisio.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                     
                    Try{Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\setup.exe' -ArgumentList "/uninstall VISIO /config uninstallvisio.xml "-NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                }

                Write-Output "Found Office 2010 starting uninstall" |Out-File $logfile -Append
                
                Try{ Copy-Item -Path $DependencyDir\Uninstall2010.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\Uninstall2010.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                    
                Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\setup.exe' -ArgumentList "/uninstall ProPlus /config uninstall2010.xml" -NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                
                break
                        
            }

            "2013"{
                
                if(Get-UninstallString "Microsoft Project Standard 2013"){
                    
                    Write-Output "Found Project 2013 starting uninstall" |Out-File $logfile -Append
                    
                    Try{ Copy-Item -Path $DependencyDir\Uninstallproject.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\UninstallProject.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                                        
                    Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\setup.exe' -ArgumentList "/uninstall PrjStd /Config uninstallproject.xml" -NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

                }
               
                if(Get-UninstallString "Microsoft Visio Standard 2013"){
                    
                    Write-Output "Found Visio 2013 starting uninstall" |Out-File $logfile -Append
                    
                    Try{ Copy-Item -Path $DependencyDir\uninstallvisio.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\uninstallvisio.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

                    Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\setup.exe' -ArgumentList "/uninstall VisStd /config uninstallvisio.xml "-NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

                }
                
                Write-Output "Found Office 2013 starting uninstall" |Out-File $logfile -Append
                
                Try{ Copy-Item -Path $DependencyDir\Uninstall2010.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\Uninstall2010.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

                Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\setup.exe' -ArgumentList "/uninstall ProPlus /config uninstall2010.xml" -NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }


                break

            }

            "2016"{
                if(Get-UninstallString "Microsoft Project Standard 2016"){
                    
                    Write-Output "Found Project 2016 starting uninstall" |Out-File $logfile -Append

                    Try{ Copy-Item -Path $DependencyDir\Uninstallproject.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\Office Setup Controller\UninstallProject.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                    
                    Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\Office Setup Controller\setup.exe' -ArgumentList "/uninstall PrjStd /Config uninstallproject.xml" -NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

                }

                if(Get-UninstallString "Microsoft Visio Standard 2016"){
                    
                    Write-Output "Found Visio 2016 starting uninstall" |Out-File $logfile -Append

                    Try{ Copy-Item -Path $DependencyDir\uninstallvisio.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\Office Setup Controller\uninstallvisio.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

                    Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\Office Setup Controller\setup.exe' -ArgumentList "/uninstall VisStd /config uninstallvisio.xml "-NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

                }

                Write-Output "Found Office 2016 starting uninstall" |Out-File $logfile -Append

                Try{ Copy-Item -Path $DependencyDir\uninstall16.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\Office Setup Controller\Uninstall16.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                
                Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\Office Setup Controller\setup.exe' -ArgumentList "/uninstall PROPLUS /config uninstall16.xml" -NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }


                break
            }
            
            default{Write-Output " There is no Default" | Out-File $logfile -Append }

        }

    }

}

#region check for One Off Visio and Projects 2013 & 2016

#region 2010
if(Get-UninstallString "Microsoft Project Standard 2010"){

    Write-Output "Found Project 2010 starting uninstall" |Out-File $logfile -Append

    Try{Copy-Item -Path $DependencyDir\Uninstallproject.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\UninstallProject.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                                                                      
    Try{Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\setup.exe' -ArgumentList "/uninstall PrjStd /Config uninstallproject.xml" -NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

}

if(Get-UninstallString "Microsoft Visio Premium 2010"){
                    
    Write-Output "Found Visio 2010 starting uninstall" |Out-File $logfile -Append

    Try{Copy-Item -Path $DependencyDir\uninstallvisio.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\uninstallvisio.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                     
    Try{Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\setup.exe' -ArgumentList "/uninstall VISIO /config uninstallvisio.xml "-NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
}
#endregion

#region 2013
 if(Get-UninstallString "Microsoft Project Standard 2013"){
                    
    Write-Output "Found Project 2013 starting uninstall" |Out-File $logfile -Append
                    
    Try{ Copy-Item -Path $DependencyDir\Uninstallproject.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\UninstallProject.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                                        
    Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\setup.exe' -ArgumentList "/uninstall PrjStd /Config uninstallproject.xml" -NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

}
               
if(Get-UninstallString "Microsoft Visio Standard 2013"){
                    
    Write-Output "Found Visio 2013 starting uninstall" |Out-File $logfile -Append
                    
    Try{ Copy-Item -Path $DependencyDir\uninstallvisio.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\uninstallvisio.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

    Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\Office Setup Controller\setup.exe' -ArgumentList "/uninstall VisStd /config uninstallvisio.xml "-NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

}
#endregion

#region 2016
if(Get-UninstallString "Microsoft Project Standard 2016"){
                    
    Write-Output "Found Project 2016 starting uninstall" |Out-File $logfile -Append

    Try{ Copy-Item -Path $DependencyDir\Uninstallproject.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\Office Setup Controller\UninstallProject.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }
                    
    Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\Office Setup Controller\setup.exe' -ArgumentList "/uninstall PrjStd /Config uninstallproject.xml" -NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

}

if(Get-UninstallString "Microsoft Visio Standard 2016"){
                    
    Write-Output "Found Visio 2016 starting uninstall" |Out-File $logfile -Append

    Try{ Copy-Item -Path $DependencyDir\uninstallvisio.xml -Destination 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\Office Setup Controller\uninstallvisio.xml' -Force -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

    Try{ Start-Process -FilePath 'C:\Program Files (x86)\Common Files\microsoft shared\OFFICE16\Office Setup Controller\setup.exe' -ArgumentList "/uninstall VisStd /config uninstallvisio.xml "-NoNewWindow -Wait -ErrorAction Stop |Out-File $logfile -Append}Catch{ Add-Content -Value $_.Exception.Message -Path $logfile }

}
#endregion
#endregion
