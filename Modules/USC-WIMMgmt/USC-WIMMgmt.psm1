function Get-WimInfo {
    Param($Release='1709',$Index,$Root='appdev')
    Begin{
        $Build = Get-ReleaseBuild $Release
        $Root = Switch ($Root) {
            appdev {'\\usc.internal\dfs\appdev\SCCMPackages\OperatingSystems'}
            cap {'\\wsp-configmgr01\DeploymentShare$\Captures'}
        }
        echo $Root
        if ($Index) {
            $Index="/Index:$Index"
        }
    }

    Process {
        Dism.exe /Get-Wiminfo /WimFile:"$Root\$Build" $Index
    }
}

function Get-WimIndex {
    Param([Parameter(Mandatory=$True)]$Path)
    Get-WindowsImage -ImagePath $Path | Select-Object -Expand ImageIndex
}

function Get-WimIndexVer {
    Param([Parameter(Mandatory=$True)]$Image)
    $Ver = $null
    If ($Image.SPBuild -gt $Image.SPLevel) {
        $Ver = $Image.SPBuild 
    } else {
        $Ver = $Image.SPLevel
    }
    return $Ver
}

function Get-WimIndexDesc {
    Param([Parameter(Mandatory=$True)]$Image)
    return $Image.ImageDescription
}

function Get-WimIndexName {
    Param([Parameter(Mandatory=$True)]$Image)
    return $Image.ImageName
}

function Get-ReleaseBuild {
    Param([Parameter(Mandatory=$True)]$Release,[switch]$Number)

    $Num = Switch ($Release) {
        1607 {'14393'}
        1703 {'15063'}
        1709 {'16299'}
        1803 {'17134'}
    }
    $Build = "Microsoft Windows 10 x64 $Num.wim"
    If ($Number) {
        return $Num
    } else {
        return $Build
    }
}

function Get-BuildRelease {
    Param($Build)

    Switch ($Build) {
       14393 {'1607'}
       15063 {'1703'}
       16299 {'1709'}
       17134 {'1803'}
    }
}

function Copy-WimIndex {
    Param($SourceRoot='\\wsp-configmgr01\DeploymentShare$\Captures',
    $DestRoot='\\usc.internal\dfs\appdev\SCCMPackages\OperatingSystems',
    [Parameter(Mandatory=$True)]$Release,
    $Index,[switch]$WhatIf)
    
    $Build = Get-ReleaseBuild $Release
    
    If ((Test-Path "$SourceRoot\$Build") -and (Test-Path "$DestRoot\$Build")) {
        If (-Not $Index) {
            $Index = Get-WimIndex -Path "$SourceRoot\$Build" | Select-Object -Last 1
        }
        $Image = Get-WindowsImage -ImagePath "$SourceRoot\$Build" -Index $Index
        $BuildVer = Get-WimIndexVer -Image $Image
        $DestName = "Microsoft Windows 10 x64 $(Get-ReleaseBuild $Release -Number) $BuildVer"
        If ($WhatIf) {

            echo Dism /Export-Image /SourceImageFile:"$SourceRoot\$Build" /SourceIndex:$Index /DestinationImageFile:"$DestRoot\$Build" /DestinationName:"$DestName"
            Update-WimIndexDesc $Release -Whatif
        } else {
            Dism /Export-Image /SourceImageFile:"$SourceRoot\$Build" /SourceIndex:$Index /DestinationImageFile:"$DestRoot\$Build" /DestinationName:"$DestName"
            Update-WimIndexDesc $Release
        }
    }
}

Function Update-WimIndexDesc {
    Param([Parameter(Mandatory=$True)]$Release,$Index,$Root='appdev',
    [switch]$Whatif)
    $Root = Switch ($Root) {
        appdev {'\\usc.internal\dfs\appdev\SCCMPackages\OperatingSystems'}
        cap {'\\wsp-configmgr01\DeploymentShare$\Captures'}
    }

    $Build = Get-ReleaseBuild $Release
    $Path = "$Root\$Build"
    If (-Not $Index) {
        $Index = Get-WimIndex -Path "$Path" | Select-Object -Last 1
    }
    $Image = Get-WindowsImage -ImagePath $Path -Index $Index
    $ImageX = Find-Imagex
    $Name = Get-WimIndexName -Image $Image
    $CurrDesc = Get-WimIndexDesc -Image $Image
    $CmdArgs = "/INFO ""$Path"" $Index ""$Name"" ""$Release"""
    If ($Whatif) {
        Echo "Updating description from $CurrDesc to $Release"
        echo "ImageX $CmdArgs"
    } else {
        Start-Process $ImageX -argumentlist $CmdArgs -Wait -NoNewWindow
    }
}

function Find-Imagex {
    "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\DISM\imagex.exe"
}

function Copy-AllWimImages {
    $Images = Get-ChildItem -Path (cap -show) | ? PSIsContainer -eq $False
    ForEach ($ImageFile in $Images) {
        $Build = $ImageFile.BaseName -split ' ' | select -last 1
        $Release = Get-BuildRelease $Build
        $Index = Get-WimIndex -Path $ImageFile.FullName | Select-Object -Last 1
        $Image= Get-WindowsImage -ImagePath $ImageFile.FullName -Index $Index
        $CapVersion = Get-WimIndexVer -Image $Image
        $AppdevImageFile = "$(wim -show)\$($ImageFile.Name)"
        $AppdevIndex = Get-WimIndex -Path "$AppdevImageFile" | Select-Object -Last 1
        $AppdevImage = Get-WindowsImage -ImagePath "$AppdevImageFile" -Index $AppdevIndex
        $AppdevVersion = Get-WimIndexVer -Image $AppdevImage
        If ($CapVersion -gt $AppdevVersion) {
            Copy-WimIndex -Release $Release 
        }

    }

}

function cap{
    Param([switch]$Show)
    $Loc = '\\wsp-configmgr01\DeploymentShare$\Captures'
    if ($Show) {
        $Loc
    } else {
        sl $Loc
    }
}

function wim{
    Param([switch]$Show)
    $Loc = '\\usc\dfs\appdev\sccmpackages\Operatingsystems'
    if ($Show) {
        $Loc
    } else {
        sl $Loc
    }
}
