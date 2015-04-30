# VBox lab control PowerShell
# Goal: Streamlined interface to launching VMs headless or not, with user directed networking config
# requires Oracle Virtualbox installed to default location

# hardcoded defaults
$defaultVMnoGUI =
$defaultVMyesGUI =  
$defaultNetwork = 


# background working. get list of VMs from vbox($_long), truncate to readable(loop of $_short)
$VBoxManage='C:\Program Files\Oracle\VirtualBox\VBoxManage.exe'
$allVMs_long = & $VBoxManage list vms
$allVMs_short = @()

foreach ($vm in $allVMs_long) {
    $vm = $vm|%{$_.split('"')[1]}
    $allVMs_short += $vm
}

$xAppName = "VBox Lab Menus"
$global:xExitSession = $false

function LoadMenuSystem() {
	$xMenu1 = 0
	$xMenu2 = 0
	$xValidSelection = $false
	
	
	# Main menu to select 1. VM no GUI, 2. VM yes GUI, 3. Exit
	while ($xMenu1 -lt 1 -or $xMenu1 -gt 3) {
		CLS  # clear screen
		Write-Host "`n`t`t`t ---=======---" -Fore Magenta
		Write-Host "`n`t`t`t VBox Lab Menu" -Fore Magenta
		Write-Host "`n`t`t`t ---=======---`n" -Fore Magenta
		Write-Host "`t`tPlease select your choice below `n" -Fore Cyan
		Write-Host "`t`t`t1. Start VM Headless" -Fore Cyan
		Write-Host "`t`t`t2. Start VM with GUI" -Fore Cyan
		Write-Host "`t`t`t3. Quit and Exit`n" -Fore Cyan
		
		$xMenu1 = Read-Host "`t`tEnter Menu Option(1-3)"
		if ($xMenu1 -lt 1 -or $xMenu1 -gt 3) {
			Write-Host "`tPlease select one of the listed options.`n" -Fore Red; start-Sleep -Seconds 1
		}
	}
	Switch ($xMenu1) {
		1 {
			# submenu 1 - list all VMs(hardcode a default)
			$quit_option = $allVMs_short.length + 1
			while ($xMenu2 -lt 1 -or $xMenu2 -gt $quit_option) {
				CLS # clear screen
				Write-Host "`t`t`t ---=======---" -Fore Magenta
				Write-Host "`n`t`t`t VBox Lab Menu" -Fore Magenta
				Write-Host "`n`t`t`t ---=======---`n" -Fore Magenta
				Write-Host "`t`tSelect VM to start HEADLESS `n" -Fore Cyan
				
				# Loop(?) to generate menu from available VMs
				# menu 1 selection, networking options choice (host only default)
				for ($i=0; $i -lt $allVMs_short.length; $i++) {
					$pretty = $i+1
					write-host "`t`t"$pretty $allVMs_short[$i]
				}

				Write-Host "`t`t"$quit_option "..`n"
				
				$xMenu2 = Read-Host "`t`tEnter Option Number"
				if ($xMenu2 -lt 1 -or $xMenu2 -gt $quit_option -or $xMenu2 -isnot [int]) {
					Write-Host "`tPlease select one of the listed options.`n" -Fore Red; start-Sleep -Seconds 1
				}
			}
			
			Switch ($xMenu2) {
				$xMenu2 {
					if ($xMenu2 -lt $quit_option) {
						$switch_start = $allVMs_short[$xMenu2 - 1]
						Write-Host "`t`tStarting " $switch_start "Headless" 
						& "$VBoxManage" startvm "$switch_start" --type headless
					}
				}
				default{break}
			}
		}
		2 {
			# submenu 2 - list all VMs(hardcode a default)
			$quit_option = $allVMs_short.length + 1
			while ($xMenu2 -lt 1 -or $xMenu2 -gt $quit_option) {
				CLS # clear screen
				Write-Host "`n`t`t`t ---=======---" -Fore Magenta
				Write-Host "`n`t`t`t VBox Lab Menu" -Fore Magenta
				Write-Host "`n`t`t`t ---=======---`n" -Fore Magenta
				Write-Host "`t`tSelect VM to start in a window `n" -Fore Cyan
				
				# Loop(?) to generate menu from available VMs
				# menu 1 selection, networking options choice (host only default)
				for ($i=0; $i -lt $allVMs_short.length; $i++) {
					$pretty = $i+1
					write-host "`t`t"$pretty $allVMs_short[$i]
				}

				Write-Host "`t`t"$quit_option "..`n"
				
				$xMenu2 = Read-Host "`t`tEnter Option Number"
				if ($xMenu2 -lt 1 -or $xMenu2 -gt $quit_option -or $xMenu2 -isnot [int]) {
					Write-Host "`tPlease select one of the listed options.`n" -Fore Red; start-Sleep -Seconds 1
				}
			}
			Switch ($xMenu2) {
				$xMenu2 {
					if ($xMenu2 -lt $quit_option) {
						$switch_start = $allVMs_short[$xMenu2 - 1]
						Write-Host "`t`tStarting" $switch_start "Windowed" 
						& "$VBoxManage" startvm "$switch_start" 
					}
				}
				default{break}
			}
		}
		default {$global:xExitSession = $true;break}
	}
}

LoadMenuSystem
if ($xExitSession) {
	exit-pssession
}else {
	.\vbox_lab.ps1
}

