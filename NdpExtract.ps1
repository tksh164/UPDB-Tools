[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $FilePath
)

#
#
#
function Write-Log
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $LogFilePath,

        [Parameter(Mandatory = $true)]
        [string] $Message
    )

    Add-Content -LiteralPath $LogFilePath -Value $Message -Encoding UTF8
}

#
#
#
function Select-ChildItem
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [System.IO.FileInfo[]] $ChildItem
    )

    Begin {}

    Process
    {
        $returnObject = [PSCustomObject] @{
            Type      = 'other'
            ChildItem = $ChildItem
        }    

        if ($ChildItem.FullName -match '.+_gac_.+')
        {
            $returnObject.Type = 'gac'
        }
        else
        {
            # for amd64
            if (($ChildItem.FullName -match '.+_amd64') -or   # *_amd64
                ($ChildItem.FullName -match '.+\.amd64') -or  # *.amd64
                ($ChildItem.FullName -match '.+_amd64\..+'))  # *_amd64.*
            {
                $returnObject.Type = 'amd64'
            }

            # for x86
            elseif (($ChildItem.FullName -match '.+_x86') -or   # *_x86
                    ($ChildItem.FullName -match '.+\.x86') -or  # *.x86
                    ($ChildItem.FullName -match '.+_x86\..+'))  # *_x86.*
            {
                $returnObject.Type = 'x86'
            }
        }

        $returnObject
    }

    End {}
}

#
#
#
function Rename-SpecificFile
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]] $ChildItem
    )

    $fileName = [System.IO.Path]::GetFileName($ChildItem.FullName)

    if (($fileName -match '.+_ini_amd64') -or ($fileName -match '.+_ini_x86'))
    {
        $newFileName = $fileName.Replace('_ini_amd64', '.ini').Replace('_ini_x86', '.ini')
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif (($fileName -match '.+\.targets\.amd64') -or ($fileName -match '.+\.targets\.x86'))
    {
        $newFileName = $fileName.Replace('.targets.amd64', '.targets').Replace('.targets.x86', '.targets')
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif (($fileName -match '.+\.targets_amd64') -or ($fileName -match '.+\.targets_x86'))
    {
        $newFileName = $fileName.Replace('.targets_amd64', '.targets').Replace('.targets_x86', '.targets')
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif (($fileName -eq 'AttributionFile_amd64') -or ($fileName -eq 'AttributionFile_x86'))
    {
        $newFileName = 'AttributionFile'
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif (($fileName -match '.+_man_amd64') -or ($fileName -match '.+_man_x86'))
    {
        $newFileName = $fileName.Replace('_man_amd64', '.man').Replace('_man_x86', '.man')
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif (($fileName -eq 'ie_browser_amd64') -or ($fileName -eq 'ie_browser_x86'))
    {
        $newFileName = 'ie.browser'
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif (($fileName -match '.+_sql_amd64') -or ($fileName -match '.+_sql_x86'))
    {
        $newFileName = $fileName.Replace('_sql_amd64', '.sql').Replace('_sql_x86', '.sql')
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif (($fileName -match '.+_nlp_amd64') -or ($fileName -match '.+_nlp_x86'))
    {
        $newFileName = $fileName.Replace('_nlp_amd64', '.nlp').Replace('_nlp_x86', '.nlp')
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif (($fileName -match '.+\.overridetasks_amd64') -or ($fileName -match '.+\.overridetasks_x86'))
    {
        $newFileName = $fileName.Replace('.overridetasks_amd64', '.overridetasks').Replace('.overridetasks_x86', '.overridetasks')
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif (($fileName -match '.+_xsd_amd64') -or ($fileName -match '.+_xsd_x86'))
    {
        $newFileName = $fileName.Replace('_xsd_amd64', '.xsd').Replace('_xsd_x86', '.xsd')
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif ($fileName -eq 'WpfEtwMan')
    {
        $newFileName = 'wpf-etw.man'
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif ($fileName -eq 'netfx45_upgradecleanup_x86')
    {
        $newFileName = 'netfx45_upgradecleanup.inf'
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }
    elseif (($fileName -match '.+\.CompositeFont_amd64') -or ($fileName -match '.+\.CompositeFont_x86'))
    {
        $newFileName = $fileName.Replace('.CompositeFont_amd64', '.CompositeFont').Replace('.CompositeFont_x86', '.CompositeFont')
        Rename-Item -LiteralPath $ChildItem.FullName -NewName $newFileName
        return $true
    }

    return $false
}


#
# Create a log directory and a log file.
#

$logDirPath = Join-Path -Path ([System.IO.Path]::GetDirectoryName($FilePath)) -ChildPath ('{0}_log' -f [System.IO.Path]::GetFileNameWithoutExtension($FilePath))
[void] (New-Item -ItemType Directory -Path $logDirPath)

$logFileName = [System.IO.Path]::GetFIleName($FilePath)
$logFilePath = Join-Path -Path $logDirPath -ChildPath ('{0}.log' -f $logFileName)

Write-Verbose -Message 'A log file created.'
Write-Log -LogFilePath $logFilePath -Message ('LogDir: "{0}"' -f $logDirPath)
Write-Log -LogFilePath $logFilePath -Message ('LogFile: "{0}"' -f $logFilePath)

#
# Extract package.
#

$extractDirectory = Join-Path -Path ([System.IO.Path]::GetDirectoryName($FilePath)) -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($FilePath))
$argList = ('/x:"{0}"' -f $extractDirectory),'/q'

Write-Verbose -Message 'The Package extracting...'
Write-Log -LogFilePath $logFilePath -Message ('ExtractDirectory: "{0}"' -f $extractDirectory)

Start-Process -FilePath $FilePath -ArgumentList $argList -Wait -WindowStyle Hidden

#
# Retrieve MSP file.
#

$mspFilePaths = Get-ChildItem -LiteralPath $extractDirectory -Filter '*.msp'
foreach ($mspFilePath in $mspFilePaths)
{
    Write-Log -LogFilePath $logFilePath -Message ('MspFilePath: "{0}"' -f $mspFilePath)

    $p = [System.IO.Path]::GetDirectoryName($mspFilePath.FullName)
    $cp = [System.IO.Path]::GetFileNameWithoutExtension($mspFilePath.FullName)

    # Make directory for extract MSP package.
    $mspExtractDirectory = New-Item -ItemType Directory -Path (Join-Path -Path $p -ChildPath $cp)

    Write-Verbose -Message 'A directory created for the MSP package extracting.'
    Write-Log -LogFilePath $logFilePath -Message ('MspExtractDirectory: "{0}"' -f $mspExtractDirectory)

    break
}

#
# Extract MSP package.
#

$msixExePath = Join-Path -Path $PSScriptRoot -ChildPath 'msix2\release\msix.exe'

$argList = ($mspFilePath.FullName),'/out',($mspExtractDirectory.FullName),'/ext'

Write-Verbose -Message 'The MSP Package extracting...'

Start-Process -FilePath $msixExePath -ArgumentList $argList -Wait -WindowStyle Hidden

#
# Retrieve CAB file.
#

$cabFilePaths = Get-ChildItem -LiteralPath $mspExtractDirectory -Filter '*.cab'

$cabFilePaths | ForEach-Object {

    Write-Log -LogFilePath $logFilePath -Message ('CabFilePath: "{0}"' -f $_.FullName)

    $p = [System.IO.Path]::GetDirectoryName($_.FullName)
    $cp = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)

    #
    # Make directory for extract CAB package.
    #

    $cabExtractDirectory = New-Item -ItemType Directory -Path (Join-Path -Path $p -ChildPath $cp)

    Write-Verbose -Message 'A directory created for the CAB package extracting.'
    Write-Log -LogFilePath $logFilePath -Message ('CabExtractDirectory: "{0}"' -f $cabExtractDirectory)

    #
    # Extract CAB package.
    #

    $argList = ('"{0}"' -f $_.FullName),'-F:*',('"{0}"' -f $cabExtractDirectory.FullName)

    Write-Verbose -Message 'The CAB Package extracting...'

    Start-Process -FilePath 'C:\Windows\System32\expand.exe' -ArgumentList $argList -Wait -WindowStyle Hidden

    #
    # Make architecture directories.
    #

    $amd64Dir = New-Item -ItemType Directory -Path (Join-Path -Path ($cabExtractDirectory.FullName) -ChildPath 'amd64')
    $x86Dir = New-Item -ItemType Directory -Path (Join-Path -Path ($cabExtractDirectory.FullName) -ChildPath 'x86')

    Write-Verbose -Message 'Architecture directories created.'
    Write-Log -LogFilePath $logFilePath -Message ('Amd64Dir: "{0}"' -f $amd64Dir)
    Write-Log -LogFilePath $logFilePath -Message ('X86Dir: "{0}"' -f $x86Dir)

    #
    # Copy specific archtecture module files.
    #

    Write-Verbose -Message 'Filtering module files by architecture...'

    Get-ChildItem -LiteralPath $cabExtractDirectory -File | Select-ChildItem | ForEach-Object {

        if ($_.Type -eq 'amd64')
        {
            Copy-Item -LiteralPath $_.ChildItem.FullName -Destination $amd64Dir.FullName
            Write-Log -LogFilePath $logFilePath -Message ('AMD64: "{0}" -> "{1}"' -f $_.ChildItem.FullName,$amd64Dir.FullName)

            #for ($i = 0; $i -lt 3; $i++)
            #{
                try {
                    Remove-Item -Path $_.ChildItem.FullName -ErrorAction Stop
                    #break
                }
                catch {
                    Write-Log -LogFilePath $logFilePath -Message $Error[0].ErrorDetails
                    Write-Log -LogFilePath $logFilePath -Message $Error[0].FullyQualifiedErrorId
                    Write-Log -LogFilePath $logFilePath -Message $Error[0].ScriptStackTrace
                }

                #Start-Sleep -Seconds 1
            #}
        }
        elseif ($_.Type -eq 'x86')
        {
            Copy-Item -LiteralPath $_.ChildItem.FullName -Destination $x86Dir.FullName
            Write-Log -LogFilePath $logFilePath -Message ('X86: "{0}" -> "{1}"' -f $_.ChildItem.FullName,$x86Dir.FullName)

            #for ($i = 0; $i -lt 3; $i++)
            #{
                try {
                    Remove-Item -Path $_.ChildItem.FullName -ErrorAction Stop
                    #break
                }
                catch {
                    Write-Log -LogFilePath $logFilePath -Message $Error[0].ErrorDetails
                    Write-Log -LogFilePath $logFilePath -Message $Error[0].FullyQualifiedErrorId
                    Write-Log -LogFilePath $logFilePath -Message $Error[0].ScriptStackTrace
                }

                #Start-Sleep -Seconds 1
            #}
        }
        elseif ($_.Type -eq 'gac')
        {
            Write-Log -LogFilePath $logFilePath -Message ('GAC: "{0}"' -f $_.ChildItem.FullName)
            Remove-Item -Path $_.ChildItem.FullName
            Write-Log -LogFilePath $logFilePath -Message ('Delete: "{0}"' -f $_.ChildItem.FullName)

        }
        else
        {
            Write-Log -LogFilePath $logFilePath -Message ('{0}: "{1}"' -f $_.Type,$_.ChildItem.FullName)
            Write-Log -LogFilePath $logFilePath -Message ('Cannot judge the architecture: "{0}"' -f $_.ChildItem.FullName)
        }
    }

    #
    # Rename files.
    #

    Write-Verbose -Message 'Renaming module files...'

    Get-ChildItem -LiteralPath ($amd64Dir.FullName,$x86Dir.FullName) -File | ForEach-Object {

        # Show progress.
        Write-Host -Object '#' -NoNewline

        # Rename by pre-defined file names.
        if (Rename-SpecificFile -ChildItem $_)
        {
           Write-Log -LogFilePath $logFilePath -Message ('Rename: "{0}"' -f $_.FullName)
        }
        else
        {
            $ExtractAppExePath = Join-Path -Path $PSScriptRoot -ChildPath 'ExtractApp\ExtractApp.exe'

            # Rename by the original file name.
            $stdOutFilePath = Join-Path -Path $logDirPath -ChildPath ('{0}.out.txt' -f ([System.IO.Path]::GetFileName($_.FullName)))
            #$stdErrFilePath = Join-Path -Path $logDirPath -ChildPath ('{0}.err.txt' -f ([System.IO.Path]::GetFileName($_.FullName)))
            $argList = ('"{0}"' -f $_.FullName)
            $proc = Start-Process -FilePath $ExtractAppExePath -ArgumentList $argList -Wait -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdOutFilePath #-RedirectStandardError $stdErrFilePath

            Write-Log -LogFilePath $logFilePath -Message ('Rename: "{0}"' -f $_.FullName)

            if ($proc.ExitCode -eq 0)
            {
                Remove-Item -Path $stdOutFilePath #,$stdErrFilePath
            }
            else {
                Write-Log -LogFilePath $logFilePath -Message ('RenameExitCodeNonZero: {0}, "{1}"' -f $proc.ExitCode, $stdOutFilePath)
            }
        }
    }    
}
