    <#
        .DESCRIPTION
        Back and Restore MineCraft Worlds

        .LINK
        Online version: http://www.fabrikam.com/extension.html

        .LINK
    #>

Add-Type -assembly "system.io.compression.filesystem"
Add-Type -AssemblyName System.Windows.Forms 
$sw = (Get-WmiObject -Class Win32_DesktopMonitor).ScreenWidth[1]
$sh = (Get-WmiObject -Class Win32_DesktopMonitor).ScreenHeight[1]
$backupfolder = ([IO.Path]::Combine(([Environment]::GetFolderPath("MyDocuments")), 'MinecraftWorld-Backups', 'minecraftWorlds'))
$defaultMineCraftdir = "$env:LOCALAPPDATA\Packages\Microsoft.MinecraftUWP_8wekyb3d8bbwe\LocalState\games\com.mojang\minecraftWorlds"

if(get-process Minecraft.Windows){Stop-Process -Name Minecraft.Windows}

$Menu = [ordered]@{
  "View" = 'View Saved MineCraft Worlds For this Profile - Sorted on Last Save Date'
  "Export" = 'Export MineCraft Worlds For this Profile'
  "Import" = 'Import MineCraft Worlds For this Profile'
  }

  $Result = $Menu | Out-GridView -PassThru  -Title 'Make a  selection'
  Switch ($Result)  {
  {$Result.Name -eq "View"} {
        $folders = gci $defaultMineCraftdir | sort LastWriteTime -Descending | % {
        $o = [PSCustomObject]@{
        LevelName = (gc "$defaultMineCraftdir\$($_.name)\levelname.txt")
        FolderName = ("$defaultMineCraftdir\$($_.name)")
        }
        $o
        }
        $folders  | Out-GridView -Title "MineCraft World Safegames"
  }

  {$Result.Name -eq "Export"} {'Export MineCraft Worlds For this Profile'
        if(!(Test-Path $backupfolder)){
        New-Item -ItemType Directory -Path $backupfolder
        }

        $Form = New-Object system.Windows.Forms.Form
        $Form.ClientSize = "$($sw / 2),$($sh / 3)"
        $Form.Text = "Exporting all of your Worlds, please wait."
        $Label = New-Object System.Windows.Forms.Label
        $Form.Controls.Add($Label)
        $Label.Text = "Do not close this screen untill all is done...
        Busy exporting all of your Worlds to: 
        $backupfolder"
        $Label.Font = "Microsoft Sans Serif,18,style=Bold"
        $Label.ForeColor = "Green"
        $Label.AutoSize = $True
        $Form.Visible = $True
        $Form.Update()

        gci "$env:LOCALAPPDATA\Packages\Microsoft.MinecraftUWP_8wekyb3d8bbwe\LocalState\games\com.mojang\minecraftWorlds" | % {
        $Source = $_.FullName
        $shortfoldername = $_.Name
        $LevelName = (gc "$env:LOCALAPPDATA\Packages\Microsoft.MinecraftUWP_8wekyb3d8bbwe\LocalState\games\com.mojang\minecraftWorlds\$($_.name)\levelname.txt") -replace '\W','-'
        $zipFileName = "$backupfolder\" + "$LevelName" + "__" + $shortfoldername + ".zip"

            if(Test-Path $zipFileName){
            Remove-Item $zipFileName -Force
            }

        write-host "Creating $zipFileName"
        [io.compression.zipfile]::CreateFromDirectory($Source, $zipFileName)
        }
        $Form.Close()
  }

  {$Result.Name -eq "Import"} {'Import MineCraft Worlds For this Profile'
            $Form = New-Object system.Windows.Forms.Form
            $Form.ClientSize = "$($sw / 2),$($sh / 3)"
            $Form.Text = "Exporting all of your Worlds, please wait."
            $Label = New-Object System.Windows.Forms.Label
            $Form.Controls.Add($Label)
            $Label.Text = "Do not close this screen untill all is done...
            Importing all of your Worlds to: 
            $defaultMineCraftdir"
            $Label.Font = "Microsoft Sans Serif,18,style=Bold"
            $Label.ForeColor = "Green"
            $Label.AutoSize = $True
            $Form.Visible = $True
            $Form.Update()


           if((gci $backupfolder).count -gt 1){
           gci $backupfolder | % {
                    $Source = $_.FullName
                    $shortfoldername = $_.Name
                    write-host "Importing: $($_.FullName) to $($defaultMineCraftdir)"
                    $destinationFilename = ((($shortfoldername -split "__")[1]) -replace ".zip","")
                    $destination = Join-Path $defaultMineCraftdir $destinationFilename
                    Remove-Item -Recurse -Force $destination
                    [io.compression.zipfile]::ExtractToDirectory($Source, $destination)
                }
        }

            $Form.Close()
  }   

} 