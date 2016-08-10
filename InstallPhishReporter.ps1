########################################################
#####      Install Phish Reporter in Outlook       #####
#####      1) Install Junk Reporter                #####
#####      3) Copy ribbon files to pc		   #####
#####	   2) Update BccEmailAddress Key           #####
#####      4) Restart Outlook                      #####
#####      Author: Spencer Alessi                  #####
#####      Modified: 8/10/2016                     #####  
########################################################

# You will need to modify the items in [ ] to fit your environment before running the script

$junkReporterPath = "[path to your JunkReporter folder]"

$computers = Get-ADComputer -Filter * -SearchBase "[your ad search query]" | select-object Name

Foreach ($pc in $computers) {
    
    If (Test-Connection -ComputerName $pc.Name -Count 1 -Quiet) {
        Try {
           
            ########################################################
            #####      1) Install Junk Reporter                #####
            ########################################################
            
            Write-Host "Copying .msi & .bat files..."

            copy "$junkReporterPath\JunkReporter.msi" \\$($pc.Name)\c$
            copy "$junkReporterPath\JunkReporter.bat" \\$($pc.Name)\c$

            Write-Host "installing..."
			psexec -u someadmin -p force \\$pc.Name -s -d cmd.exe /c "c:\junkreporter.bat"
        
            #wait 30 seconds for install to finish
            Start-Sleep -m 30000                     
                               
            

            ########################################################
            #####      3) Copy ribbon files to pc              #####
            #####      *Note: Ribbon files are profile based.  #####
            #####      Pushing these files may be better off   #####
            #####      done via group policy or a login script #####
            ########################################################

            # Get the current logged on username and push the outlook ribbon files to their profile
            Write-Host "Finding logged on user..."
            $explorerprocesses = @(Get-WmiObject -ComputerName $pc.Name -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue)
            if ($explorerprocesses.Count -eq 0)
            {
                write-host $pc.Name": No explorer process found / Nobody interactively logged on"
            } else {
                foreach ($i in $explorerprocesses)
                {
                    $username = $i.GetOwner().User
                    $userPath = "Users\$($username)\AppData\Local\Microsoft\Office"
                    Write-Host "Copying " "$junkReporterPath\Client Files\olkmailread.officeUI" " to: " \\$($pc.Name)\c$\$($userPath)                        
                    Copy-Item -Force "$junkReporterPath\Client Files\olkmailread.officeUI" -Destination \\$($pc.Name)\c$\$($userPath)
                    Copy-Item -Force "$junkReporterPath\Client Files\olkexplorer.officeUI" -Destination \\$($pc.Name)\c$\$($userPath)
                }
            }
             

            
            ########################################################
            #####      2) Update Registry Key                  #####
            ########################################################
        
            $regPath = "SOFTWARE\\wow6432node\\Microsoft\\Junk E-mail Reporting\\Addins"
            $Property = "BccEmailAddress"
            $Value = "[yoursecurityteam@somedomain.com]"
            $key = Get-ItemPropertyValue -Path $path -Name BccEmailAddress
            $currentPC = $pc.Name

            Write-Host "Looking for the correct registry key..."
            If ($key -eq $null) {
                Write-Host "key not found"
            } Else {
                Write-Host "Updating the registry key..."
                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $currentPC) 
                $regKey= $reg.OpenSubKey($regPath,$true) 
                $regKey.SetValue($Property,$Value,[Microsoft.Win32.RegistryValueKind]::String)
            }
                        
        } Catch {
            write-host $pc.Name": Could not connect"
        }
    } Else {
        write-host $pc.Name": Could not connect"
    }

#}
