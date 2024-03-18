<#
.SYNOPSIS
    This script runs the Applicaiton Beanchmark Tests in different environment.
.DESCRIPTION
    This script chooses an appropriate defaults for most of the parameters if not provided.
.NOTES
    File Name      : RunBCPTTests.ps1
    Copyright 2020 - Microsof Corporation
.EXAMPLE
    RunBCPTTests.ps1  -Environment PROD -AuthorizationType AAD -Credential $Credential -SandboxName sandbox -TestRunnerPage 50010 -SuiteCode ONE -BCPTTestRunnerInternalFolderPath .\ALTestRunner\Internal
.EXAMPLE
    RunBCPTTests.ps1 -Environment OnPrem -AuthorizationType Windows -TestRunnerPage 50010 -SuiteCode THREE -BCPTTestRunnerInternalFolderPath .\ALTestRunner\Internal
#>
param
(
    [ValidateSet("PROD","OnPrem")]
    [string] $Environment,
    [ValidateSet("AAD","Windows","NavUserPassword")]
    [string] $AuthorizationType,
    [pscredential] $Credential,
    [pscredential] $Token,
    [string] $SandboxName,
    [string] $ServiceUrl,
    [int] $TestRunnerPage,
    [Parameter(Mandatory=$true)]
    [string] $SuiteCode,
    [string] $BCPTTestRunnerInternalFolderPath,
    [string] $ClientId,
    [switch] $SingleRun,
    [switch] $SkipDelayBeforeStart,
    [string] $CompanyName
)

<#
This funciton verifies the passed parameters and sets the defaults for the missing ones. The parameters are saved as global variables for further use
#>
function VerifyAndSetup-Parameters(
    [string] $Environment,
    [string] $AuthorizationType,
    [pscredential] $Credential,
    [pscredential] $Token,
    [string] $SandboxName,
    [string] $ServiceUrl,
    [int] $TestRunnerPage = $script:DefaultTestRunnerPage,
    [string] $SuiteCode,
    [string] $BCPTTestRunnerInternalFolderPath,
    [string] $ClientId
)
{
    # If -Environment is not specified then pick the default
    if ($Environment -eq '')
    {
        Write-Host "-Environment parameter is not provided. Defaulting to $script:DefaultEnvironment"

        $global:Environment = $script:DefaultEnvironment
    }
    else
    {
        $global:Environment = $Environment
    }

    # Depending on the Environment make sure necessary parameters are also specified
    switch ($global:Environment)
    {
        # PROD works only with AAD authorizatin type and OnPrem works on all 3 Authorization types
        'PROD'
        {
            if ($AuthorizationType -ne 'AAD')
            {
                throw "Only Authorization type 'AAD' can work in -Environment $Environment."
            }
            else
            {
                if ($AuthorizationType -eq '')
                {
                    Write-Host "-AuthorizationType parameter is not provided. Defaulting to $script:DefaultAuthorizationType"
                    $global:AuthorizationType = $script:DefaultAuthorizationType
                }
                else
                {
                    $global:AuthorizationType = $AuthorizationType
                }
            }

            if ($SandboxName -eq '')
            {
                Write-Host "-SandboxName parameter is not provided. Defaulting to $script:DefaultSandboxName"
                $global:SandboxName = $script:DefaultSandboxName
            }
            else
            {
                $global:SandboxName = $SandboxName
            }

            if ($ClientId -eq '')
            {
                Write-Error -Category InvalidArgument -Message 'ClientId is mandatory in the PROD environment'
            }
            else
            {
                $global:ClientId = $ClientId
            }
        }
        'OnPrem'
        {
            if ($AuthorizationType -eq '')
            {
                Write-Host "-AuthorizationType parameter is not provided. Defaulting to $script:DefaultAuthorizationType"
                $global:AuthorizationType = $script:DefaultAuthorizationType
            }
            else
            {
                $global:AuthorizationType = $AuthorizationType
            }

            # OnPrem, -ServiceUrl should be provided else default is selected. On other environments, the Service Urls are built
            if ($ServiceUrl -eq '')
            {
                Write-Host "Valid ServiceUrl is not provided. Defaulting to $script:DefaultServiceUrl"
                $global:ServiceUrl = $script:DefaultServiceUrl
            }
            else
            {
                $global:ServiceUrl = $ServiceUrl
            }
            if (-Not [string]::IsNullOrEmpty($CompanyName) -And $global:ServiceUrl -inotmatch "company=")
            {
                $global:ServiceUrl = "$global:ServiceUrl&company=$CompanyName"
            }
        }
    }

    switch ($global:AuthorizationType)
    {
        # -Credential or -Token should be specified if authorization type is AAD.
        "AAD"
        {
            if ($Credential -eq $null -and $Token -eq $null)
            {
                throw "Parameter -Credential or -Token should be defined when selecting 'AAD' authorization type."
            }
            if ($Credential -ne $null -and $Token -ne $null)
            {
                throw "Specify only one parameter -Credential or -Token when selecting 'AAD' authorization type."
            }
        }
        # -Credential should be specified if authorization type is NavUserPassword.
        "NavUserPassword"
        {
            if ($Credential -eq $null)
            {
                throw "Parameter -Credential should be defined when selecting 'NavUserPassword' authorization type."
            }
        }
        "Windows"
        {
            if ($Credential -ne $null)
            {
                throw "Parameter -Credential should not be defined when selecting 'Windows' authorization type."
            }
        }
    }

    # Make sure all necessary scripts and assemblies are in the specified folder
    if ($BCPTTestRunnerInternalFolderPath -eq '')
    {
        Write-Host "Valid BCPTTestRunnerInternalFolderPath is not provided. Defaulting to $script:BCPTTestRunnerInternalFolderPath"
        $global:BCPTTestRunnerInternalFolderPath = $script:DefaultBCPTTestRunnerInternalFolderPath
    }
    else
    {
        $global:BCPTTestRunnerInternalFolderPath = $BCPTTestRunnerInternalFolderPath
    }

    if ([System.IO.Path]::IsPathRooted($global:BCPTTestRunnerInternalFolderPath) -eq $false)
    {
        Write-Host "The passed path for -BCPTTestRunnerInternalFolderPath seem to be relative path, converting to absolute."
        $global:BCPTTestRunnerInternalFolderPath = Convert-Path -Path $global:BCPTTestRunnerInternalFolderPath
        Write-Host "Absolute path is $global:BCPTTestRunnerInternalFolderPath"
    }

    if ((Test-Path -Path (Join-Path $global:BCPTTestRunnerInternalFolderPath $script:BCPTTestRunnerInternalScriptName)) -eq $false)
    {
         throw "Cannot find $script:BCPTTestRunnerInternalScriptName in location $global:BCPTTestRunnerInternalFolderPath"
    }

    if ((Test-Path -Path (Join-Path $global:BCPTTestRunnerInternalFolderPath $script:ClientAssembly1)) -eq $false)
    {
         throw "Cannot find $script:ClientAssembly1 in location $global:BCPTTestRunnerInternalFolderPath"
    }

    if ((Test-Path -Path (Join-Path $global:BCPTTestRunnerInternalFolderPath $script:ClientAssembly2)) -eq $false)
    {
         throw "Cannot find $script:ClientAssembly2 in location $global:BCPTTestRunnerInternalFolderPath"
    }

    if($AuthorizationType -eq "AAD")
    {
        if ((Test-Path -Path (Join-Path $global:BCPTTestRunnerInternalFolderPath $script:ClientAssembly3)) -eq $false)
        {
            throw "Cannot find $script:ClientAssembly2 in location $global:BCPTTestRunnerInternalFolderPath"
        }
    }

    if ($TestRunnerPage -eq 0)
    {
        Write-Host "Valid TestRunnerPage is not provided. Defaulting to $script:DefaultTestRunnerPage"
        $global:TestRunnerPage = $script:DefaultTestRunnerPage
    }
    else
    {
        $global:TestRunnerPage = $TestRunnerPage
    }
}

$ErrorActionPreference = "Stop"

$global:Environment = ''
$global:AuthorizationType = ''
$global:SandboxName = ''
$global:ClientId = ''
$global:ServiceUrl = ''
$global:TestRunnerPage = 0
$global:BCPTTestRunnerInternalFolderPath = ''

$script:DefaultEnvironment = "OnPrem"
$script:DefaultAuthorizationType = 'Windows'
$script:DefaultSandboxName = "sandbox"
$script:DefaultServiceUrl = 'http://localhost:48900'
$script:DefaultTestRunnerPage = '149002'
$script:DefaultBCPTTestRunnerInternalFolderPath = $PSScriptRoot

$script:BCPTTestRunnerInternalScriptName = "BCPTTestRunnerInternal.psm1"
$script:ClientAssembly1 = "Microsoft.Dynamics.Framework.UI.Client.dll"
$script:ClientAssembly2 = "NewtonSoft.Json.dll"

VerifyAndSetup-Parameters -Environment $Environment -AuthorizationType $AuthorizationType -Credential $Credential -Token $Token -SandboxName $SandboxName -ServiceUrl $ServiceUrl -TestRunnerPage $TestRunnerPage -SuiteCode $SuiteCode -BCPTTestRunnerInternalFolderPath $BCPTTestRunnerInternalFolderPath -ClientId $ClientId

$BCPTRTestRunnerInternalScriptPath = (Join-Path $global:BCPTTestRunnerInternalFolderPath $script:BCPTTestRunnerInternalScriptName)
Write-Host "Importing AL Test Runner from $BCPTRTestRunnerInternalScriptPath"
Import-Module $BCPTRTestRunnerInternalScriptPath -WarningAction SilentlyContinue -Force

# Fetch no. of tasks, tasks life time and no. of tests from the client
$NoOfTasks,$TaskLifeInMins,$NoOfTests = Get-NoOfIterations -DisableSSLVerification -Environment $global:Environment -AuthorizationType $global:AuthorizationType -Credential $Credential -Token $Token -SandboxName $global:SandboxName -ServiceUrl $global:ServiceUrl -TestRunnerPage $global:TestRunnerPage -SuiteCode $SuiteCode -ClientId $global:ClientId -CompanyName $CompanyName

$TaskLifeInMins += 10;

# Run-BCPTTestsInternal -DisableSSLVerification -Environment $global:Environment -AuthorizationType $global:AuthorizationType -Credential $Credential -Token $Token -SandboxName $global:SandboxName -ServiceUrl $global:ServiceUrl -TestRunnerPage $global:TestRunnerPage -SuiteCode $SuiteCode -SessionTimeoutInMins $TaskLifeInMins -ClientId $global:ClientId

$task = {
    param(
        [string] $Environment,
        [string] $AuthorizationType,
        [pscredential] $Credential,
        [pscredential] $Token,
        [string] $SandboxName,
        [string] $ServiceUrl,
        [int] $TestRunnerPage,
        [string] $SuiteCode,
        [string] $PathToRunnerScript,
        [int] $SessionTimeoutInMins,
        [string] $ClientId,
        [switch] $SingleRun
    )
        
    Write-Host "Importing AL Test Runner from $PathToRunnerScript"
    Import-Module $PathToRunnerScript -WarningAction SilentlyContinue -Force

    Run-BCPTTestsInternal -DisableSSLVerification -Environment $Environment -AuthorizationType $AuthorizationType -Credential $Credential -Token $Token -SandboxName $SandboxName -ServiceUrl $ServiceUrl -TestRunnerPage $TestRunnerPage -SuiteCode $SuiteCode -SessionTimeoutInMins $SessionTimeoutInMins -ClientId $ClientId -SingleRun:$SingleRun -CompanyName $CompanyName
}

# Determine how many sessions to spawn based on single run mode
if ($SingleRun.IsPresent)
{
    $NoOfSessions = $NoOfTests
}
else
{
    $NoOfSessions = $NoOfTasks
}
    
# Spawn sessions
for ($i = 0; $i -lt $NoOfSessions; $i++)
{
    Start-Job -ScriptBlock $task -ArgumentList $global:Environment,$global:AuthorizationType,$Credential,$Token,$global:SandboxName,$global:ServiceUrl,$global:TestRunnerPage,$SuiteCode,$BCPTRTestRunnerInternalScriptPath,$TaskLifeInMins, $global:ClientId, $SingleRun

    if ($SkipDelayBeforeStart -eq $false)
    {sleep -Seconds 2}
}

While (@(Get-Job | Where { $_.State -eq "Running" -and $_.PSJobTypeName -eq "BackgroundJob"}).Count -ne 0) {
    Write-Verbose "Waiting for background jobs..."
    Start-Sleep -Seconds 3
}

$Data = ForEach ($Job in (Get-Job | Where { $_.PSJobTypeName -eq "BackgroundJob"}))
{  
    try {
        Receive-Job $Job
    }
    catch { 
        Write-Host $_ -ForegroundColor Red
    }
    finally {
        Remove-Job $Job
    }
}

# SIG # Begin signature block
# MIInwwYJKoZIhvcNAQcCoIIntDCCJ7ACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBtt7apVoI6Xg6r
# Pjch46TPGNkD4mYgF8Ok46d4tI7cQqCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
# esGEb+srAAAAAANOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI5WhcNMjQwMzE0MTg0MzI5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDdCKiNI6IBFWuvJUmf6WdOJqZmIwYs5G7AJD5UbcL6tsC+EBPDbr36pFGo1bsU
# p53nRyFYnncoMg8FK0d8jLlw0lgexDDr7gicf2zOBFWqfv/nSLwzJFNP5W03DF/1
# 1oZ12rSFqGlm+O46cRjTDFBpMRCZZGddZlRBjivby0eI1VgTD1TvAdfBYQe82fhm
# WQkYR/lWmAK+vW/1+bO7jHaxXTNCxLIBW07F8PBjUcwFxxyfbe2mHB4h1L4U0Ofa
# +HX/aREQ7SqYZz59sXM2ySOfvYyIjnqSO80NGBaz5DvzIG88J0+BNhOu2jl6Dfcq
# jYQs1H/PMSQIK6E7lXDXSpXzAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUnMc7Zn/ukKBsBiWkwdNfsN5pdwAw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMDUxNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAD21v9pHoLdBSNlFAjmk
# mx4XxOZAPsVxxXbDyQv1+kGDe9XpgBnT1lXnx7JDpFMKBwAyIwdInmvhK9pGBa31
# TyeL3p7R2s0L8SABPPRJHAEk4NHpBXxHjm4TKjezAbSqqbgsy10Y7KApy+9UrKa2
# kGmsuASsk95PVm5vem7OmTs42vm0BJUU+JPQLg8Y/sdj3TtSfLYYZAaJwTAIgi7d
# hzn5hatLo7Dhz+4T+MrFd+6LUa2U3zr97QwzDthx+RP9/RZnur4inzSQsG5DCVIM
# pA1l2NWEA3KAca0tI2l6hQNYsaKL1kefdfHCrPxEry8onJjyGGv9YKoLv6AOO7Oh
# JEmbQlz/xksYG2N/JSOJ+QqYpGTEuYFYVWain7He6jgb41JbpOGKDdE/b+V2q/gX
# UgFe2gdwTpCDsvh8SMRoq1/BNXcr7iTAU38Vgr83iVtPYmFhZOVM0ULp/kKTVoir
# IpP2KCxT4OekOctt8grYnhJ16QMjmMv5o53hjNFXOxigkQWYzUO+6w50g0FAeFa8
# 5ugCCB6lXEk21FFB1FdIHpjSQf+LP/W2OV/HfhC3uTPgKbRtXo83TZYEudooyZ/A
# Vu08sibZ3MkGOJORLERNwKm2G7oqdOv4Qj8Z0JrGgMzj46NFKAxkLSpE5oHQYP1H
# tPx1lPfD7iNSbJsP6LiUHXH1MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGaMwghmfAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCggbIwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEINqCoyHaWG5Y29BMIvgn0n/C
# Ir4EoFtW59QmoK7hXanXMEYGCisGAQQBgjcCAQwxODA2oBiAFgBCAEMAUABUAEwA
# aQBiAHIAYQByAHmhGoAYaHR0cDovL3d3dy5taWNyb3NvZnQuY29tMA0GCSqGSIb3
# DQEBAQUABIIBAEppba/6GjQ8lPHALOSVyU2+ILXH3jeP1eH3H82og3PLmkmQX50T
# JM8+Ef5OTf/YPhAdUvSrB/gi7NAPoVU47nRfl8sFO/511rY9/+wpK99VgmAu67wr
# n99LHrDn/UttxVIr1SHpcAaNQvKFs/f4Lfk0g+W6vm0m+mDd9ccObYyR77XBmvOC
# 5SJirGaxI0DF6+EYj7btQQWgfCYs5gOS1Iy6lNGgZeXaBOKWerqV2eObJbfXb0BZ
# cDK30PpzIaEXm0IXhpChklAMxhS5zYKWADLl71+tOyNaVCddxpOBs43CrD2/qJj1
# BsTSQxiK0z0bYGzSpU16sFS9jmEElBV+oM6hghcpMIIXJQYKKwYBBAGCNwMDATGC
# FxUwghcRBgkqhkiG9w0BBwKgghcCMIIW/gIBAzEPMA0GCWCGSAFlAwQCAQUAMIIB
# WQYLKoZIhvcNAQkQAQSgggFIBIIBRDCCAUACAQEGCisGAQQBhFkKAwEwMTANBglg
# hkgBZQMEAgEFAAQgPY4CLxI9UEWZrbHDUpz1SfFTwlKG7ZPN0Yjbs3UQBaACBmUv
# 82PsExgTMjAyMzEwMjAwODIxNTQuNTY1WjAEgAIB9KCB2KSB1TCB0jELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0
# IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNT
# IEVTTjoxNzlFLTRCQjAtODI0NjElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgU2VydmljZaCCEXgwggcnMIIFD6ADAgECAhMzAAABta0a39eFcG0TAAEAAAG1
# MA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4X
# DTIyMDkyMDIwMjIxMVoXDTIzMTIxNDIwMjIxMVowgdIxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5k
# IE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046MTc5
# RS00QkIwLTgyNDYxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZp
# Y2UwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCXCwq4ZV6Cwo2JhcXB
# T4JIee3Zrs99bDbI/CJb5hGUQmwfeJsUiI9+T/JuUynwkZqe277sOL/7d7nBY3+d
# slQgxSMRYmdzwkyOg+YcJawy0k65H/VG7hprVGvFC0h6WrT3xMGXiTs2wQTzuXuo
# oWbLyLAY1COTCvRsy7gEnDLsFReJcd6A5JT33at9DjE1mjHOe/jZMtCxz02ZwGrl
# ayWSpUBNxzw9J+AEsSl0bUOnbCo3DgmskXAk0DVt/NNJxgAFzmyZFkGCw/gmIr/w
# JWuWhyF4TJCieObesW22uiMCt5JSeLEAu72kOuwEjgNQ8YbfighuJ4jWioWX/GTs
# D7u4zyTijyJ8xVY1NpzNs+V0Ni2fqEGt7uvblEQPi55wLE/wPHLfhg9QSVaWU8/c
# sBenlwzBGPH4RbOMS0gsQpe4Bx/GtcJPDJiGblc3MIJliHj+AXZTbL9th96Qqlqg
# VjCShl5PpnGhBtfkp+2CP61mQVXbVS9kiQnpvfhr/jh22uwrp2JOZWc1Mz0no/oS
# GWTax9xTBscgYvyZip0o6LJ3hHE4KZK9wU2RkSCuTnjyYcJsNtPI73FLM3USRbFu
# nyf8sdWdn1+nmEgsae2QzNoPlpdEJDIdIiGw+1HtD+sqgfKMWqC0Y71+2fFi/HM2
# tsp4wLPCZyrAytfodzF7RJ4pnwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFGxQhtmC
# hlVdVdVZBfQ6TVU7ZsdTMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1Gely
# MF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lv
# cHMvY3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNy
# bDBsBggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBD
# QSUyMDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCQUY6nKMpX
# azawD7BOoPAN2GnSYWs+2JTis3c6idNapvzkzpYfX1z8/nXvG6MsKKH8eWU/nEpa
# ZecAhFXX81AOQkEtJ0tNv81C1xPUVYZDsIxOeuf20tnGogW0pXKW9A3KHfUL+qQL
# xCY2nLNqQ7Qbfy44aA6Qn0SricD7tDBV+hsOWCBa4SnN0WdF3JcvfaA6pK+suMUq
# m/goRWpoFOJrEJZkOUiEbdiBTNrtycblnZID994yr8iJXToDNRp1nDpeBxOupeYS
# YTS0yZ2XjgwwULfAZT0BV4UJYp0P1Y6dgDbMeD5cXYtBW2jRVi0Ut3pzoNAVyYHf
# wa88IpVhKfyvXQShe5E36BDuEtyKbXZ6w6dbgXhscKYdZNVZ4AaA+JIi0cdF8Yos
# RRmai1U50U9HzCK4ANP98dJcGSR2kvXa2+AQYQt5POfMW6VnXpv82/p21uBJFmt5
# 6wkE0qlON1iO78aqeUCSl+UvonTDGT3nv9RVieaONFjrWNf3RAZCYHOb2+z/7ZuP
# TZfH9tdfLy+rnuOY1dDa585ombecBwPao5pLJcQ6P2aqEy3i128yMeGI0+V1+PRD
# sqrUeB1EGspaTMJA2Li2zwdEkGKk1pWlZ9TsFxGfwF5jN0ugjPBEsO1q3PpoGGje
# gP9tcetlqDGmszkywwH+tV9vlefVhLv+4TCCB3EwggVZoAMCAQICEzMAAAAVxedr
# ngKbSZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRp
# ZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4
# MzIyNVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3
# DQEBAQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qls
# TnXIyjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLA
# EBjoYH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrE
# qv1yaa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyF
# Vk3v3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1o
# O5pGve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg
# 3viSkR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2
# TPYrbqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07B
# MzlMjgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJ
# NmSLW6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6
# r1AFemzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+
# auIurQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3
# FQIEFgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl
# 0mWnG1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUH
# AgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0
# b3J5Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMA
# dQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAW
# gBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8v
# Y3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRf
# MjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRw
# Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEw
# LTA2LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL
# /Klv6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu
# 6WZnOlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5t
# ggz1bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfg
# QJY4rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8s
# CXgU6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCr
# dTDFNLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZ
# c9d/HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2
# tVdUCbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8C
# wYKiexcdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9
# JZTmdHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDB
# cQZqELQdVTNYs6FwZvKhggLUMIICPQIBATCCAQChgdikgdUwgdIxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBF
# U046MTc5RS00QkIwLTgyNDYxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2WiIwoBATAHBgUrDgMCGgMVAI0wn2vXVFmPQ9a7e6T5pAcXcixVoIGD
# MIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwDQYJKoZIhvcNAQEF
# BQACBQDo3GvTMCIYDzIwMjMxMDIwMTEwMDM1WhgPMjAyMzEwMjExMTAwMzVaMHQw
# OgYKKwYBBAGEWQoEATEsMCowCgIFAOjca9MCAQAwBwIBAAICAMswBwIBAAICEiww
# CgIFAOjdvVMCAQAwNgYKKwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgC
# AQACAwehIKEKMAgCAQACAwGGoDANBgkqhkiG9w0BAQUFAAOBgQBh03/edviC98Tf
# 9PrODNkVIQcnShGF9I6tb0tl1+Tw/AY3sOKI+88CbiumtUt0I7GZEmAUKriLwecn
# 8pJDSGwqpFI4aNXnOV6cKL142sbNV9ej82JfvM5hT3Zsi43KcJIFjV4qBJXdC4md
# oNFZ0n0qrOQpgP1aOo66SbFKId07sjGCBA0wggQJAgEBMIGTMHwxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQSAyMDEwAhMzAAABta0a39eFcG0TAAEAAAG1MA0GCWCGSAFl
# AwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcN
# AQkEMSIEIMDyv1aHBjQOuwQyJwYjvFZ4ZGxIlxhsgwkhEkfUfm+2MIH6BgsqhkiG
# 9w0BCRACLzGB6jCB5zCB5DCBvQQgJ8oNNS1oZxaJ9hzc5WcimntiSfRLwlyVXOuU
# CAXxyIMwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAA
# AbWtGt/XhXBtEwABAAABtTAiBCA6bfkReLVuonWP76MnOSo/xWVBcAFqz1Fo0RKp
# cmPuqTANBgkqhkiG9w0BAQsFAASCAgA50FXpCTlbGLk+CCAV5frVf7f9iXWx5MqF
# bzQr8Lae+mv1ZHbjhLHzDyH4mBZBJl0Vh5jhMtfAd4JFklJ6HUFK1sGDWpjSz7PO
# aRJKu0sRWgD2SYCRR15U51txiT9I0/koaBQzq9ABDDhh4oc/BXPk0FqyH+cTQDLa
# /fRfPsu5MqsvXqXUOXE9ZMOwh2veCd+XaMiowuF7hYbMBvK4tAHCB4Q5Rt43S97L
# QuW3TDi7rCBHZR1XFm+EieDMIxLqTC/3Oi2jQ18RwZn09hyVtIzvnFE9GdWPHClA
# QT92PY8wGECh4FbqKLcJ7YJq7XPZX3r2yLhPHY2yMH8YEEs3koqh4e1Ey8sy+gp0
# Xbc7eDXP/Lt0tMb4pNDbZJHhOh7AzwAnB9X6AxLZqp/oQOvGoans9GoXfOx4iGc+
# +ycYmdxT/hDffuRHW8yt5PJXGX2EO9FE7Fy/JWXOwwLkd1buektYKSQb0gu6EE2T
# ZS1L01ngEyEkJYgDjYUB+WU7HAKjivRcZKWKnb7526FqCMo6tyZ3Ns1X69Jxss90
# IThEwNVInUXawvixHVrczT0lLicxjkkj663wfehcbxopmrYMTqUnBSIBN0MAIh/l
# +NoZVu5QeSoDgWMm8Whfk9KJ5BS4fmqEELEL8W6Uga/0lvb/Jmq1Rv6tOSC/YUOf
# iz9XR+bJ/A==
# SIG # End signature block
