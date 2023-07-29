# 獲取捷徑信息
function Get-Shortcut {
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string]$Path
    )

    # 檢測路徑
    if ($Path) {
        [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
        $Path = [System.IO.Path]::GetFullPath($Path)
        if (!(Test-Path -PathType:Leaf $Path)) { Write-Error "Error:: Path `"$Path`" does not exist" -ErrorAction:Stop }
    }

    # 讀取檔案
    $File = Get-Item -Path $Path
    
    # 讀取 ShellLink 物件
    $shell = New-Object -ComObject Shell.Application
    $dir = $shell.NameSpace($File.DirectoryName)
    $item = $dir.Items().Item($File.Name)
    $link = $item.GetLink()

    # Return the ShellLinkObject
    return $link
} # (Get-Shortcut "C:\Users\User\Desktop\お　大粒焼帆立貝　80ｇ-320x480.jpg.lnk")

# 修改捷徑信息
function Set-Shortcut {
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string]$Path,
        [Parameter(ParameterSetName = "")]
        [string]$TargetPath,
        [Parameter(ParameterSetName = "")]
        [string]$Description,
        [Parameter(ParameterSetName = "")]
        [string]$Arguments,
        [Parameter(ParameterSetName = "")]
        [string]$WorkingDirectory,
        [Parameter(ParameterSetName = "")]
        [int]$Hotkey,
        [Parameter(ParameterSetName = "")]
        [int]$ShowCommand
    )

    # 檢測路徑
    if ($Path) {
        [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
        $Path = [System.IO.Path]::GetFullPath($Path)
        if (!(Test-Path -PathType:Leaf $Path)) { Write-Error "Error:: Path `"$Path`" does not exist" -ErrorAction:Stop }
    }

    # 讀取 ShellLink 物件
    $link = Get-Shortcut -Path $Path

    # 更新捷徑屬性
    if ($PSBoundParameters.ContainsKey('TargetPath')) {
        $link.Path = $TargetPath
    }
    if ($PSBoundParameters.ContainsKey('Description')) {
        $link.Description = $Description
    }
    if ($PSBoundParameters.ContainsKey('Arguments')) {
        $link.Arguments = $Arguments
    }
    if ($PSBoundParameters.ContainsKey('WorkingDirectory')) {
        $link.WorkingDirectory = $WorkingDirectory
    }
    if ($PSBoundParameters.ContainsKey('Hotkey')) {
        $link.Hotkey = $Hotkey
    }
    if ($PSBoundParameters.ContainsKey('ShowCommand')) {
        $link.ShowCommand = $ShowCommand
    }

    # 保存變更
    $link.Save()

    return $link
} # Set-Shortcut -Path "C:\Users\User\Desktop\マルセイバターサンド.jpg.lnk" -Description "AAA"

# 新增捷徑
function New-Shortcut {
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string]$TargetPath,
        [Parameter(ParameterSetName = "")]
        [string]$Path,
        [Parameter(ParameterSetName = "")]
        [string]$Description,
        [Parameter(ParameterSetName = "")]
        [string]$Arguments,
        [Parameter(ParameterSetName = "")]
        [string]$WorkingDirectory,
        [Parameter(ParameterSetName = "")]
        [int]$Hotkey,
        [Parameter(ParameterSetName = "")]
        [int]$ShowCommand
    )
    
    # 檢測 TargetPath
    if ($TargetPath) {
        [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
        $TargetPath = [System.IO.Path]::GetFullPath($TargetPath)
        if (!(Test-Path -PathType:Leaf $TargetPath)) { Write-Error "Error:: Target path `"$TargetPath`" does not exist" -ErrorAction:Stop }
    }

    # 讀取檔案
    $File = Get-Item -Path $TargetPath

    # 如果 Path 為空，則使用當前工作目錄和目標檔案的檔名作為路徑
    if (!$Path) { $Path = Join-Path -Path (Get-Location) -ChildPath ($File.BaseName + ".lnk") }
    if (!$WorkingDirectory) { $WorkingDirectory = $File.DirectoryName }
    
    # 創建空捷徑檔案
    $randomPath = Join-Path -Path (Get-Location) -ChildPath ([IO.Path]::GetRandomFileName() + ".lnk")
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($randomPath)
    $Shortcut.Save()
    Move-Item -Path $randomPath -Destination $Path -Force

    # 使用 Set-Shortcut 函數設定捷徑的屬性
    Set-Shortcut -Path $Path -TargetPath $TargetPath -Description $Description -Arguments $Arguments -WorkingDirectory $WorkingDirectory -Hotkey $Hotkey -ShowCommand $ShowCommand
} # New-Shortcut -TargetPath "C:\Users\User\Desktop\お　大粒焼帆立貝　80ｇ-320x480.jpg"


# 創建捷徑
function CreateShortcut {
    param (
        # 捷徑屬性
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $TargetPath,
        [Parameter(ParameterSetName = "")]
        [string] $Arguments,        # 捷徑的參數
        [Parameter(ParameterSetName = "")]
        [string] $WorkingDirectory, # 捷徑的開始位置
        [Parameter(ParameterSetName = "")]
        [string] $Description,      # 捷徑的描述
        # 輸出位置
        [Parameter(ParameterSetName = "")]
        [string] $Path
    )
    
    # 檢測 TargetPath
    [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
    if ($TargetPath) {
        $TargetPath = [System.IO.Path]::GetFullPath($TargetPath)
        if (!(Test-Path -PathType:Leaf $TargetPath)) { Write-Error "Error:: Target path `"$TargetPath`" does not exist" -ErrorAction:Stop }
    } $File = Get-Item -Path $TargetPath

    # 修正 Path 路徑
    if ($Path) { $Path = [System.IO.Path]::GetFullPath($Path) }
    if (!$Path) { # 路徑為空自動補上當前路徑
        $Path = Join-Path -Path (Get-Location) -ChildPath ($File.BaseName + ".lnk")
    } else { # 路徑參數存在
        $Extension = [IO.Path]::GetExtension($Path)
        if (!$Extension) {
            $Path = $Path + "\$($File.BaseName).lnk"
            $DirPath = Split-Path $Path
            if (!(Test-Path $DirPath)) { New-Item $DirPath -ItemType:Directory -Force |Out-Null }
        } elseif($Extension -ne '.lnk') {
            $Path = $Path + '.lnk'
        }
    }

    # 工作目錄 預設為目標檔案所在目錄
    if (!$WorkingDirectory) { $WorkingDirectory = $File.DirectoryName }
    
    # 創建空捷徑檔案
    $randomPath = Join-Path -Path (Get-Location) -ChildPath ([IO.Path]::GetRandomFileName() + ".lnk")
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($randomPath)
    $Shortcut.Save()
    Move-Item -Path $randomPath -Destination $Path -Force

    # 讀取 ShellLink 物件
    $File = Get-Item -Path $Path
    $shell = New-Object -ComObject Shell.Application
    $dir = $shell.NameSpace($File.DirectoryName)
    $item = $dir.Items().Item($File.Name)
    $link = $item.GetLink()

    # 修改屬性
    if ($TargetPath)       { $link.Path = $TargetPath }
    if ($Arguments)        { $link.Arguments = $Arguments }
    if ($WorkingDirectory) { $link.WorkingDirectory = $WorkingDirectory }
    if ($Description)      { $link.Description = $Description }
    $link.Save()

    # 回傳物件
    return $link
}
# CreateShortcut ".\README.md"
# CreateShortcut ".\README.md" -Path "aaa.ink"
# CreateShortcut ".\README.md" -Path "bbb.exe"
# CreateShortcut ".\README.md" -Path DirName
# CreateShortcut ".\README.md" -Description "Test Description"
