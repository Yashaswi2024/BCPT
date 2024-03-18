function Run-AlTestsInternal
(
    [string] $TestSuite = $script:DefaultTestSuite,
    [string] $TestCodeunitsRange = "",
    [string] $TestProcedureRange = "",
    [string] $ExtensionId = "",
    [int] $TestRunnerId = $global:DefaultTestRunner,
    [ValidateSet('Windows','NavUserPassword','AAD')]
    [string] $AutorizationType = $script:DefaultAuthorizationType,
    [string] $TestPage = $global:DefaultTestPage,
    [switch] $DisableSSLVerification,
    [Parameter(Mandatory=$true)]
    [string] $ServiceUrl,
    [Parameter(Mandatory=$false)]
    [pscredential] $Credential,
    [bool] $Detailed = $true,
    [array] $DisabledTests = @(),
    [ValidateSet('Disabled', 'PerRun', 'PerCodeunit', 'PerTest')]
    [string] $CodeCoverageTrackingType = 'Disabled',
    [string] $CodeCoverageOutputPath = "$PSScriptRoot\CodeCoverage",
    [string] $CodeCoverageExporterId,
    [switch] $CodeCoverageTrackAllSessions,
    [string] $CodeCoverageFilePrefix,
    [bool] $StabilityRun
)
{
    $ErrorActionPreference = $script:DefaultErrorActionPreference
   
    Setup-TestRun -DisableSSLVerification:$DisableSSLVerification -AutorizationType $AutorizationType -Credential $Credential -ServiceUrl $ServiceUrl -TestSuite $TestSuite -TestCodeunitsRange $TestCodeunitsRange -TestProcedureRange $TestProcedureRange -ExtensionId $ExtensionId -TestRunnerId $TestRunnerId -TestPage $TestPage -DisabledTests $DisabledTests -CodeCoverageTrackingType $CodeCoverageTrackingType -CodeCoverageTrackAllSessions:$CodeCoverageTrackAllSessions -CodeCoverageOutputPath $CodeCoverageOutputPath -CodeCoverageExporterId $CodeCoverageExporterId -StabilityRun $StabilityRun
            
    $testRunResults = New-Object System.Collections.ArrayList 
    $testResult = ''
    $numberOfUnexpectedFailures = 0;

    do
    {
        try
        {
            $testStartTime = $(Get-Date)
            $testResult = Run-NextTest -DisableSSLVerification:$DisableSSLVerification -AutorizationType $AutorizationType -Credential $Credential -ServiceUrl $ServiceUrl -TestSuite $TestSuite
            if($testResult -eq $script:AllTestsExecutedResult)
            {
                return [Array]$testRunResults
            }
 
            $testRunResultObject = ConvertFrom-Json $testResult
            if($CodeCoverageTrackingType -ne 'Disabled') {
                $null = CollectCoverageResults -TrackingType $CodeCoverageTrackingType -OutputPath $CodeCoverageOutputPath -DisableSSLVerification:$DisableSSLVerification -AutorizationType $AutorizationType -Credential $Credential -ServiceUrl $ServiceUrl -CodeCoverageFilePrefix $CodeCoverageFilePrefix
            }
       }
        catch
        {
            $numberOfUnexpectedFailures++

            $stackTrace = $_.Exception.StackTrace + "Script stack trace: " + $_.ScriptStackTrace 
            $testMethodResult = @{
                method = "Unexpected Failure"
                codeUnit = "Unexpected Failure"
                startTime = $testStartTime.ToString($script:DateTimeFormat)
                finishTime = ($(Get-Date).ToString($script:DateTimeFormat))
                result = $script:FailureTestResultType
                message = $_.Exception.Message
                stackTrace = $stackTrace
            }

            $testRunResultObject = @{
                name = "Unexpected Failure"
                codeUnit = "UnexpectedFailure"
                startTime = $testStartTime.ToString($script:DateTimeFormat)
                finishTime = ($(Get-Date).ToString($script:DateTimeFormat))
                result = $script:FailureTestResultType
                testResults = @($testMethodResult)
            }
        }
        
        $testRunResults.Add($testRunResultObject) > $null
        if($Detailed)
        {
            Print-TestResults -TestRunResultObject $testRunResultObject
        }
    }
    until((!$testRunResultObject) -or ($NumberOfUnexpectedFailuresBeforeAborting -lt $numberOfUnexpectedFailures))

    throw "Expected to end the test execution, something went wrong with returning test results."      
}

function CollectCoverageResults {
    param (
        [ValidateSet('PerRun', 'PerCodeunit', 'PerTest')]
        [string] $TrackingType,
        [string] $OutputPath,
        [switch] $DisableSSLVerification,
        [ValidateSet('Windows','NavUserPassword','AAD')]
        [string] $AutorizationType = $script:DefaultAuthorizationType,
        [Parameter(Mandatory=$false)]
        [pscredential] $Credential,
        [Parameter(Mandatory=$true)]
        [string] $ServiceUrl,
        [string] $CodeCoverageFilePrefix
    )
    try{
        $clientContext = Open-ClientSessionWithWait -DisableSSLVerification:$DisableSSLVerification -AuthorizationType $AutorizationType -Credential $Credential -ServiceUrl $ServiceUrl
        $form = Open-TestForm -TestPage $TestPage -ClientContext $clientContext
        do {
            $clientContext.InvokeAction($clientContext.GetActionByName($form, "GetCodeCoverage"))

            $CCResultControl = $clientContext.GetControlByName($form, "CCResultsCSVText")
            $CCInfoControl = $clientContext.GetControlByName($form, "CCInfo")
            $CCResult = $CCResultControl.StringValue
            $CCInfo = $CCInfoControl.StringValue
            if($CCInfo -ne $script:CCCollectedResult){
                $CCInfo = $CCInfo -replace ",","-"
                $CCOutputFilename = $CodeCoverageFilePrefix +"_$CCInfo.dat"
                Write-Host "Storing coverage results of $CCCodeunitId in:  $OutputPath\$CCOutputFilename"
                Set-Content -Path "$OutputPath\$CCOutputFilename" -Value $CCResult
            }
        } while ($CCInfo -ne $script:CCCollectedResult)
       
        $clientContext.CloseForm($form)
    }
    finally{
        if($clientContext){
            $clientContext.Dispose()
        }
    }
}

function Print-TestResults
(
    $TestRunResultObject
)
{              
    $startTime = Convert-ResultStringToDateTimeSafe -DateTimeString $TestRunResultObject.startTime
    $finishTime = Convert-ResultStringToDateTimeSafe -DateTimeString $TestRunResultObject.finishTime
    $duration = $finishTime.Subtract($startTime)
    $durationSeconds = [Math]::Round($duration.TotalSeconds,3)

    switch($TestRunResultObject.result)
    {
        $script:SuccessTestResultType
        {
            Write-Host -ForegroundColor Green "Success - Codeunit $($TestRunResultObject.name) - Duration $durationSeconds seconds"
            break;
        }
        $script:FailureTestResultType
        {
            Write-Host -ForegroundColor Red "Failure - Codeunit $($TestRunResultObject.name) -  Duration $durationSeconds seconds"
            break;
        }
        default
        {
            if($codeUnitId -ne "0")
            {
                Write-Host -ForegroundColor Yellow "No tests were executed - Codeunit $"
            }
        }
    }

    if($TestRunResultObject.testResults)
    {
        foreach($testFunctionResult in $TestRunResultObject.testResults)
        {
            $durationSeconds = 0;
            $methodName = $testFunctionResult.method

            if($testFunctionResult.result -ne $script:SkippedTestResultType)
            {
                $startTime = Convert-ResultStringToDateTimeSafe -DateTimeString $testFunctionResult.startTime
                $finishTime = Convert-ResultStringToDateTimeSafe -DateTimeString $testFunctionResult.finishTime
                $duration = $finishTime.Subtract($startTime)
                $durationSeconds = [Math]::Round($duration.TotalSeconds,3)
            }

            switch($testFunctionResult.result)
            {
                $script:SuccessTestResultType
                {
                    Write-Host -ForegroundColor Green "   Success - Test method: $methodName - Duration $durationSeconds seconds)"
                    break;
                }
                $script:FailureTestResultType
                {
                    $callStack = $testFunctionResult.stackTrace
                    Write-Host -ForegroundColor Red "   Failure - Test method: $methodName - Duration $durationSeconds seconds"
                    Write-Host -ForegroundColor Red "      Error:"
                    Write-Host -ForegroundColor Red "         $($testFunctionResult.message)"
                    Write-Host -ForegroundColor Red "      Call Stack:"                    
                    if($callStack)
                    {
                        Write-Host -ForegroundColor Red "         $($callStack.Replace(';',"`n         "))"
                    }
                    break;
                }
                $script:SkippedTestResultType
                {
                    Write-Host -ForegroundColor Yellow "   Skipped - Test method: $methodName"
                    break;
                }
            }
        }
    }            
}

function Setup-TestRun
(
    [switch] $DisableSSLVerification,
    [ValidateSet('Windows','NavUserPassword','AAD')]
    [string] $AutorizationType = $script:DefaultAuthorizationType,
    [Parameter(Mandatory=$false)]
    [pscredential] $Credential,
    [Parameter(Mandatory=$true)]
    [string] $ServiceUrl,
    [string] $TestSuite = $script:DefaultTestSuite,
    [string] $TestCodeunitsRange = "",
    [string] $TestProcedureRange = "",
    [string] $ExtensionId = "",
    [int] $TestRunnerId = $global:DefaultTestRunner,
    [string] $TestPage = $global:DefaultTestPage,
    [array] $DisabledTests = @(),
    [ValidateSet('Disabled', 'PerRun', 'PerCodeunit', 'PerTest')]
    [string] $CodeCoverageTrackingType = 'Disabled',
    [string] $CodeCoverageOutputPath = "$PSScriptRoot\CodeCoverage",
    [string] $CodeCoverageExporterId,
    [switch] $CodeCoverageTrackAllSessions,
    [bool] $StabilityRun
)
{
    Write-Host "Setting up test run: $CodeCoverageTrackingType - $CodeCoverageOutputPath"
    if($CodeCoverageTrackingType -ne 'Disabled')
    {
        if (-not (Test-Path -Path $CodeCoverageOutputPath))
        {
            $null = New-Item -Path $CodeCoverageOutputPath -ItemType Directory
        }
    }

    try
    {
        $clientContext = Open-ClientSessionWithWait -DisableSSLVerification:$DisableSSLVerification -AuthorizationType $AutorizationType -Credential $Credential -ServiceUrl $ServiceUrl 

        $form = Open-TestForm -TestPage $TestPage -DisableSSLVerification:$DisableSSLVerification -AuthorizationType $AutorizationType -ClientContext $clientContext
        Set-TestSuite -TestSuite $TestSuite -ClientContext $clientContext -Form $form
        Set-ExtensionId -ExtensionId $ExtensionId -Form $form -ClientContext $clientContext
        Set-TestCodeunits -TestCodeunitsFilter $TestCodeunitsRange -Form $form -ClientContext $clientContext
        Set-TestProcedures -Filter $TestProcedureRange -Form $form $ClientContext $clientContext
        Set-TestRunner -TestRunnerId $TestRunnerId -Form $form -ClientContext $clientContext
        Set-RunFalseOnDisabledTests -DisabledTests $DisabledTests -Form $form -ClientContext $clientContext
        Set-StabilityRun -StabilityRun $StabilityRun -Form $form -ClientContext $clientContext
        Clear-TestResults -Form $form -ClientContext $clientContext
        if($CodeCoverageTrackingType -ne 'Disabled'){
            Set-CCTrackingType -Value $CodeCoverageTrackingType -Form $form -ClientContext $clientContext
            Set-CCTrackAllSessions -Value:$CodeCoverageTrackAllSessions -Form $form -ClientContext $clientContext
            Set-CCExporterID -Value $CodeCoverageExporterId -Form $form -ClientContext $clientContext
            Clear-CCResults -Form $form -ClientContext $clientContext
        }
        $clientContext.CloseForm($form)
    }
    finally
    {
        if($clientContext)
        {
            $clientContext.Dispose()
        }
    }
}

function Run-NextTest
(
    [switch] $DisableSSLVerification,
    [ValidateSet('Windows','NavUserPassword','AAD')]
    [string] $AutorizationType = $script:DefaultAuthorizationType,
    [Parameter(Mandatory=$false)]
    [pscredential] $Credential,
    [Parameter(Mandatory=$true)]
    [string] $ServiceUrl,
    [string] $TestSuite = $script:DefaultTestSuite
)
{
    try
    {
        $clientContext = Open-ClientSessionWithWait -DisableSSLVerification:$DisableSSLVerification -AuthorizationType $AutorizationType -Credential $Credential -ServiceUrl $ServiceUrl
        $form = Open-TestForm -TestPage $TestPage -DisableSSLVerification:$DisableSSLVerification -AuthorizationType $AutorizationType -ClientContext $clientContext
        if($TestSuite -ne $script:DefaultTestSuite)
        {
            Set-TestSuite -TestSuite $TestSuite -ClientContext $clientContext -Form $form
        }

        $clientContext.InvokeAction($clientContext.GetActionByName($form, "RunNextTest"))
        
        $testResultControl = $clientContext.GetControlByName($form, "TestResultJson")
        $testResultJson = $testResultControl.StringValue
        $clientContext.CloseForm($form)
        return $testResultJson
    }
    finally
    {
        if($clientContext)
        {
            $clientContext.Dispose()
        }
    } 
}

function Open-ClientSessionWithWait
(
    [ValidateSet('Windows','NavUserPassword','AAD')]
    [string] $AuthorizationType = $script:DefaultAuthorizationType,
    [switch] $DisableSSLVerification,
    [string] $ServiceUrl,
    [pscredential] $Credential,
    [int] $ClientSessionTimeout = 20
)
{
        $lastErrorMessage = ""
        while(($ClientSessionTimeout -gt 0) -and (-not $clientSessionOpen))
        {
            try
            {
                $clientContext = Open-ClientSession -DisableSSLVerification:$DisableSSLVerification -AuthorizationType $AuthorizationType -Credential $Credential -ServiceUrl $ServiceUrl
                return $clientContext
            }
            catch
            {
                Start-Sleep -Seconds 1
                $ClientSessionTimeout--
                $lastErrorMessage = $_.Exception.Message
            }
        }

        throw "Could not open the client session. Check if the web server is running and you can log in. Last error: $lastErrorMessage"
}

function Set-TestCodeunits
(
    [string] $TestCodeunitsFilter,
    [ClientContext] $ClientContext,
    $Form
)
{
    if(!$TestCodeunitsFilter)
    {
        return
    }

    $testCodeunitRangeFilterControl = $ClientContext.GetControlByName($Form, "TestCodeunitRangeFilter")
    $ClientContext.SaveValue($testCodeunitRangeFilterControl, $TestCodeunitsFilter)
}

function Set-TestRunner
(
    [int] $TestRunnerId,
    [ClientContext] $ClientContext,
    $Form
)
{
    if(!$TestRunnerId)
    {
        return
    }

    $testRunnerCodeunitIdControl = $ClientContext.GetControlByName($Form, "TestRunnerCodeunitId")
    $ClientContext.SaveValue($testRunnerCodeunitIdControl, $TestRunnerId)
}

function Clear-TestResults
(
    [ClientContext] $ClientContext,
    $Form
)
{
    $ClientContext.InvokeAction($ClientContext.GetActionByName($Form, "ClearTestResults"))
}

function Set-ExtensionId
(
    [string] $ExtensionId,
    [ClientContext] $ClientContext,
    $Form
)
{
    if(!$ExtensionId)
    {
        return
    }

    $extensionIdControl = $ClientContext.GetControlByName($Form, "ExtensionId")
    $ClientContext.SaveValue($extensionIdControl, $ExtensionId)
}

function Set-TestSuite
(
    [string] $TestSuite = $script:DefaultTestSuite,
    [ClientContext] $ClientContext,
    $Form
)
{
    $suiteControl = $ClientContext.GetControlByName($Form, "CurrentSuiteName")
    $ClientContext.SaveValue($suiteControl, $TestSuite)
}

function Set-CCTrackingType
{
    param (
        [ValidateSet('Disabled', 'PerRun', 'PerCodeunit', 'PerTest')]
        [string] $Value,
        [ClientContext] $ClientContext,
        $Form
    )
    $TypeValues = @{
        Disabled = 0
        PerRun = 1
        PerCodeunit=2
        PerTest=3
    }
    $suiteControl = $ClientContext.GetControlByName($Form, "CCTrackingType")
    $ClientContext.SaveValue($suiteControl, $TypeValues[$Value])
}

function Set-CCTrackAllSessions
{
    param (
        [switch] $Value,
        [ClientContext] $ClientContext,
        $Form
    )
    if($Value){
        $suiteControl = $ClientContext.GetControlByName($Form, "CCTrackAllSessions");
        $ClientContext.SaveValue($suiteControl, $Value)
    }
}

function Set-CCExporterID
{
    param (
        [string] $Value,
        [ClientContext] $ClientContext,
        $Form
    )
    if($Value){
        $suiteControl = $ClientContext.GetControlByName($Form, "CCExporterID");
        $ClientContext.SaveValue($suiteControl, $Value)
    }
}

function Set-TestProcedures
{
    param (
        [string] $Filter,
        [ClientContext] $ClientContext,
        $Form
    )
    $Control = $ClientContext.GetControlByName($Form, "TestProcedureRangeFilter")
    $ClientContext.SaveValue($Control, $Filter)
}

function Clear-CCResults
{
    param (
        [ClientContext] $ClientContext,
        $Form
    )
    $ClientContext.InvokeAction($ClientContext.GetActionByName($Form, "ClearCodeCoverage"))
}
function Set-StabilityRun
(
    [bool] $StabilityRun,
    [ClientContext] $ClientContext,
    $Form
)
{
    $stabilityRunControl = $ClientContext.GetControlByName($Form, "StabilityRun")
    $ClientContext.SaveValue($stabilityRunControl, $StabilityRun)
}

function Set-RunFalseOnDisabledTests
(
    [ClientContext] $ClientContext,
    [array] $DisabledTests,
    $Form
)
{
    if(!$DisabledTests)
    {
        return
    }

    foreach($disabledTestMethod in $DisabledTests)
    {
        $testKey = $disabledTestMethod.codeunitName + "," + $disabledTestMethod.method
        $removeTestMethodControl = $ClientContext.GetControlByName($Form, "DisableTestMethod")
        $ClientContext.SaveValue($removeTestMethodControl, $testKey)
    }
}

function Open-TestForm(
    [int] $TestPage = $global:DefaultTestPage,
    [ClientContext] $ClientContext
)
{ 
    $form = $ClientContext.OpenForm($TestPage)
    if (!$form) 
    {
        throw "Cannot open page $TestPage. Verify if the test tool and test objects are imported and can be opened manually."
    }

    return $form;
}

function Open-ClientSession
(
    [switch] $DisableSSLVerification,
    [ValidateSet('Windows','NavUserPassword','AAD')]
    [string] $AuthorizationType,
    [Parameter(Mandatory=$false)]
    [pscredential] $Credential,
    [Parameter(Mandatory=$true)]
    [string] $ServiceUrl,
    [string] $Culture = $script:DefaultCulture,
    [timespan] $TransactionTimeout = $script:DefaultTransactionTimeout,
    [timespan] $TcpKeepActive = $script:DefaultTcpKeepActive
)
{
    [System.Net.ServicePointManager]::SetTcpKeepAlive($true, [int]$TcpKeepActive.TotalMilliseconds, [int]$TcpKeepActive.TotalMilliseconds)

    if($DisableSSLVerification)
    {
        Disable-SslVerification
    }

    switch ($AuthorizationType)
    {
        "Windows" 
        {
            $clientContext = [ClientContext]::new($ServiceUrl, $TransactionTimeout, $Culture)
            break;
        }
        "NavUserPassword" 
        {
            if ($Credential -eq $null -or $Credential -eq [System.Management.Automation.PSCredential]::Empty) 
            {
                throw "You need to specify credentials if using NavUserPassword authentication"
            }
        
            $clientContext = [ClientContext]::new($ServiceUrl, $Credential, $TransactionTimeout, $Culture)
            break;
        }
        "AAD"
        {
            $AadTokenProvider = $global:AadTokenProvider
            if ($AadTokenProvider -eq $null) 
            {
                throw "You need to specify the AadTokenProvider for obtaining the token if using AAD authentication"
            }

            $token = $AadTokenProvider.GetToken($Credential)
            $tokenCredential = [Microsoft.Dynamics.Framework.UI.Client.TokenCredential]::new($token)
            $clientContext = [ClientContext]::new($ServiceUrl, $tokenCredential, $TransactionTimeout, $Culture)
        }
    }

    return $clientContext;
}

function Disable-SslVerification
{
    if (-not ([System.Management.Automation.PSTypeName]"SslVerification").Type)
    {
        Add-Type -TypeDefinition  @"
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public static class SslVerification
{
    private static bool ValidationCallback(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; }
    public static void Disable() { System.Net.ServicePointManager.ServerCertificateValidationCallback = ValidationCallback; }
    public static void Enable()  { System.Net.ServicePointManager.ServerCertificateValidationCallback = null; }
}
"@
    }
    [SslVerification]::Disable()
}

function Enable-SslVerification
{
    if (([System.Management.Automation.PSTypeName]"SslVerification").Type)
    {
        [SslVerification]::Enable()
    }
}


function Convert-ResultStringToDateTimeSafe([string] $DateTimeString)
{
    [datetime]$parsedDateTime = New-Object DateTime
    
    try
    {
        [datetime]$parsedDateTime = [datetime]$DateTimeString
    }
    catch
    {
        Write-Host -ForegroundColor Red "Failed parsing DateTime: $DateTimeString"
    }

    return $parsedDateTime
}

if(!$script:TypesLoaded)
{
    Add-type -Path "$PSScriptRoot\Microsoft.Dynamics.Framework.UI.Client.dll"
    Add-type -Path "$PSScriptRoot\NewtonSoft.Json.dll"
    
    $clientContextScriptPath = Join-Path $PSScriptRoot "ClientContext.ps1"
    . "$clientContextScriptPath"
}

$script:TypesLoaded = $true;

$script:ActiveDirectoryDllsLoaded = $false;
$script:DateTimeFormat = 's';

# Console test tool
$global:DefaultTestPage = 130455;
$global:AadTokenProvider = $null

# Test Isolation Disabled
$global:TestRunnerIsolationCodeunit = 130450
$global:TestRunnerIsolationDisabled = 130451
$global:DefaultTestRunner = $global:TestRunnerIsolationCodeunit
$global:TestRunnerAppId = "23de40a6-dfe8-4f80-80db-d70f83ce8caf"

$script:CodeunitLineType = '0'
$script:FunctionLineType = '1'

$script:FailureTestResultType = '1';
$script:SuccessTestResultType = '2';
$script:SkippedTestResultType = '3';

$script:NumberOfUnexpectedFailuresBeforeAborting = 50;

$script:DefaultAuthorizationType = 'NavUserPassword'
$script:DefaultTestSuite = 'DEFAULT'
$script:DefaultErrorActionPreference = 'Stop'

$script:DefaultTcpKeepActive = [timespan]::FromMinutes(2);
$script:DefaultTransactionTimeout = [timespan]::FromMinutes(10);
$script:DefaultCulture = "en-US";

$script:AllTestsExecutedResult = "All tests executed."
$script:CCCollectedResult = "Done."
Export-ModuleMember -Function Run-AlTestsInternal,Open-ClientSessionWithWait, Open-TestForm, Open-ClientSession

# SIG # Begin signature block
# MIInxQYJKoZIhvcNAQcCoIIntjCCJ7ICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZyrQFtw1sPpia
# +24+QuL5XYVGpFZ/flno+oJAhbk8TqCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGaUwghmhAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCggbIwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEILvrrmN5jIVEGjaxxRzYSMSD
# Vn2AS55TzaS2uU+avirNMEYGCisGAQQBgjcCAQwxODA2oBiAFgBCAEMAUABUAEwA
# aQBiAHIAYQByAHmhGoAYaHR0cDovL3d3dy5taWNyb3NvZnQuY29tMA0GCSqGSIb3
# DQEBAQUABIIBAJNVn/59LTyxOXnpRzKUuMA90t+dnfhNjqCWNQHR7dGGTYGeHpNf
# 6pyKddLrWpFZB/E2aThe9uoPi83pHJLp+73zowjAwhwdo2EpyqcJxEUL5y2nam7Y
# MP0H+6kNtmPs6NwKnUBb8mSAOSzD24IOCszernWmrbcQTuSeHY2MqR7Djd7fUIge
# knfjqnXzhLihimQa1+fQJJCMaRqS3Pu+3cLR2Mf6rQs5j9n0xcd7q7t1Lo45t1T2
# KZZVWVTcwJRiVVD5qB4w4cpvcuN310jKWeR8g4XXgITZMbu0stmcuEiBZ6ITSi4e
# BICTYxO0xExc18sbe2NFehcCB7yWTQZoW2qhghcrMIIXJwYKKwYBBAGCNwMDATGC
# FxcwghcTBgkqhkiG9w0BBwKgghcEMIIXAAIBAzEPMA0GCWCGSAFlAwQCAQUAMIIB
# WAYLKoZIhvcNAQkQAQSgggFHBIIBQzCCAT8CAQEGCisGAQQBhFkKAwEwMTANBglg
# hkgBZQMEAgEFAAQgtkrMdIPI1TWLCOWBsJZbbKhzShqlKNSBymxdtaLWHoYCBmUv
# 3u/MGxgSMjAyMzEwMjAwODIxNTEuNjJaMASAAgH0oIHYpIHVMIHSMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQg
# SXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
# RVNOOjhENDEtNEJGNy1CM0I3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloIIRezCCBycwggUPoAMCAQICEzMAAAGz/iXOKRsbihwAAQAAAbMw
# DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcN
# MjIwOTIwMjAyMjAzWhcNMjMxMjE0MjAyMjAzWjCB0jELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQg
# T3BlcmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjo4RDQx
# LTRCRjctQjNCNzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALR8D7rmGICuLLBggrK9
# je3hJSpc9CTwbra/4Kb2eu5DZR6oCgFtCbigMuMcY31QlHr/3kuWhHJ05n4+t377
# PHondDDbz/dU+q/NfXSKr1pwU2OLylY0sw531VZ1sWAdyD2EQCEzTdLD4KJbC6wm
# AConiJBAqvhDyXxJ0Nuvlk74rdVEvribsDZxzClWEa4v62ENj/HyiCUX3MZGnY/A
# hDyazfpchDWoP6cJgNCSXmHV9XsJgXJ4l+AYAgaqAvN8N+EpN+0TErCgFOfwZV21
# cg7vgenOV48gmG/EMf0LvRAeirxPUu+jNB3JSFbW1WU8Z5xsLEoNle35icdET+G3
# wDNmcSXlQYs4t94IWR541+PsUTkq0kmdP4/1O4GD54ZsJ5eUnLaawXOxxT1fgbWb
# 9VRg1Z4aspWpuL5gFwHa8UNMRxsKffor6qrXVVQ1OdJOS1JlevhpZlssSCVDodMc
# 30I3fWezny6tNOofpfaPrtwJ0ukXcLD1yT+89u4uQB/rqUK6J7HpkNu0fR5M5xGt
# Och9nyncO9alorxDfiEdb6zeqtCfcbo46u+/rfsslcGSuJFzlwENnU+vQ+JJ6jJR
# UrB+mr51zWUMiWTLDVmhLd66//Da/YBjA0Bi0hcYuO/WctfWk/3x87ALbtqHAbk6
# i1cJ8a2coieuj+9BASSjuXkBAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQU0BpdwlFn
# UgwYizhIIf9eBdyfw40wHwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIw
# XwYDVR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9w
# cy9jcmwvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3Js
# MGwGCCsGAQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENB
# JTIwMjAxMCgxKS5jcnQwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEF
# BQcDCDAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAFqGuzfOsAm4
# wAJfERmJgWW0tNLLPk6VYj53+hBmUICsqGgj9oXNNatgCq+jHt03EiTzVhxteKWO
# LoTMx39cCcUJgDOQIH+GjuyjYVVdOCa9Fx6lI690/OBZFlz2DDuLpUBuo//v3e4K
# ns412mO3A6mDQkndxeJSsdBSbkKqccB7TC/muFOhzg39mfijGICc1kZziJE/6HdK
# CF8p9+vs1yGUR5uzkIo+68q/n5kNt33hdaQ234VEh0wPSE+dCgpKRqfxgYsBT/5t
# Xa3e8TXyJlVoG9jwXBrKnSQb4+k19jHVB3wVUflnuANJRI9azWwqYFKDbZWkfQ8t
# pNoFfKKFRHbWomcodP1bVn7kKWUCTA8YG2RlTBtvrs3CqY3mADTJUig4ckN/MG6A
# Ir8Q+ACmKBEm4OFpOcZMX0cxasopdgxM9aSdBusaJfZ3Itl3vC5C3RE97uURsVB2
# pvC+CnjFtt/PkY71l9UTHzUCO++M4hSGSzkfu+yBhXMGeBZqLXl9cffgYPcnRFjQ
# T97Gb/bg4ssLIFuNJNNAJub+IvxhomRrtWuB4SN935oMfvG5cEeZ7eyYpBZ4Dbkv
# N44ZvER0EHRakL2xb1rrsj7c8I+auEqYztUpDnuq6BxpBIUAlF3UDJ0SMG5xqW/9
# hLMWnaJCvIerEWTFm64jthAi0BDMwnCwMIIHcTCCBVmgAwIBAgITMwAAABXF52ue
# AptJmQAAAAAAFTANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMwMTgz
# MjI1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBAOThpkzntHIhC3miy9ckeb0O1YLT/e6cBwfSqWxO
# dcjKNVf2AX9sSuDivbk+F2Az/1xPx2b3lVNxWuJ+Slr+uDZnhUYjDLWNE893MsAQ
# GOhgfWpSg0S3po5GawcU88V29YZQ3MFEyHFcUTE3oAo4bo3t1w/YJlN8OWECesSq
# /XJprx2rrPY2vjUmZNqYO7oaezOtgFt+jBAcnVL+tuhiJdxqD89d9P6OU8/W7IVW
# Te/dvI2k45GPsjksUZzpcGkNyjYtcI4xyDUoveO0hyTD4MmPfrVUj9z6BVWYbWg7
# mka97aSueik3rMvrg0XnRm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75xqRdbZ2De
# +JKRHh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9fvzZnkXftnIv231fgLrbqn427DZM
# 9ituqBJR6L8FA6PRc6ZNN3SUHDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1DTsEz
# OUyOArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7XKHYC4jMYctenIPDC+hIK12NvDMk2
# ZItboKaDIV1fMHSRlJTYuVD5C4lh8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqv
# UAV6bMURHXLvjflSxIUXk8A8FdsaN8cIFRg/eKtFtvUeh17aj54WcmnGrnu3tz5q
# 4i6tAgMBAAGjggHdMIIB2TASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsGAQQBgjcV
# AgQWBBQqp1L+ZMSavoKRPEY1Kc8Q/y8E7jAdBgNVHQ4EFgQUn6cVXQBeYl2D9OXS
# ZacbUzUZ6XIwXAYDVR0gBFUwUzBRBgwrBgEEAYI3TIN9AQEwQTA/BggrBgEFBQcC
# ARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9zaXRv
# cnkuaHRtMBMGA1UdJQQMMAoGCCsGAQUFBwMIMBkGCSsGAQQBgjcUAgQMHgoAUwB1
# AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaA
# FNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8y
# MDEwLTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAt
# MDYtMjMuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCdVX38Kq3hLB9nATEkW+Geckv8
# qW/qXBS2Pk5HZHixBpOXPTEztTnXwnE2P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7p
# Zmc6U03dmLq2HnjYNi6cqYJWAAOwBb6J6Gngugnue99qb74py27YP0h1AdkY3m2C
# DPVtI1TkeFN1JFe53Z/zjj3G82jfZfakVqr3lbYoVSfQJL1AoL8ZthISEV09J+BA
# ljis9/kpicO8F7BUhUKz/AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNhcy4sa3tuPywJ
# eBTpkbKpW99Jo3QMvOyRgNI95ko+ZjtPu4b6MhrZlvSP9pEB9s7GdP32THJvEKt1
# MMU0sHrYUP4KWN1APMdUbZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMBV0lUZNlz
# 138eW0QBjloZkWsNn6Qo3GcZKCS6OEuabvshVGtqRRFHqfG3rsjoiV5PndLQTHa1
# V1QJsWkBRH58oWFsc/4Ku+xBZj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8vwLB
# gqJ7Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrpNPgkNWcr4A245oyZ1uEi6vAnQj0l
# lOZ0dFtq0Z4+7X6gMTN9vMvpe784cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFx
# BmoQtB1VM1izoXBm8qGCAtcwggJAAgEBMIIBAKGB2KSB1TCB0jELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IEly
# ZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVT
# Tjo4RDQxLTRCRjctQjNCNzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAg
# U2VydmljZaIjCgEBMAcGBSsOAwIaAxUAcYtE6JbdHhKlwkJeKoCV1JIkDmGggYMw
# gYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUF
# AAIFAOjbrogwIhgPMjAyMzEwMTkyMTMyNTZaGA8yMDIzMTAyMDIxMzI1NlowdzA9
# BgorBgEEAYRZCgQBMS8wLTAKAgUA6NuuiAIBADAKAgEAAgIQTgIB/zAHAgEAAgIT
# YjAKAgUA6N0ACAIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAow
# CAIBAAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBACsIqqucg/e2
# 5PGtaH33a0PF6eOKMaJ9OQ/X4QQmpJsc/wLEIcyKeYY4PQF4gORbEK6g8bq/iPhF
# EOraH5uttNgZPvOYJjYPgxc1N1nwXwQ0ehjyoD4bMdmbK/ywb+NVoEFHEWaAI4q0
# /Kx+y2JPIttnZBfc3lEGvkDNO2GTzJd6MYIEDTCCBAkCAQEwgZMwfDELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0
# IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAGz/iXOKRsbihwAAQAAAbMwDQYJYIZI
# AWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG
# 9w0BCQQxIgQg9r3j1wEJvlOJvJF/WQ9xpIBjEldBXqgfF6iL4+fe7T4wgfoGCyqG
# SIb3DQEJEAIvMYHqMIHnMIHkMIG9BCCGoTPVKhDSB7ZG0zJQZUM2jk/ll1zJGh6K
# Ohn76k+/QjCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMz
# AAABs/4lzikbG4ocAAEAAAGzMCIEICXOGYh1HrpJwBqiDYVK+XTEinSXOOU3JYKM
# eo/X04t0MA0GCSqGSIb3DQEBCwUABIICAEBJIdt9cuoarVLFbnsxD1rf1rns9EYz
# 847t7DggYrRyk0gkqs6bp/UOrp+3I7nqrMLsHvxzJCl9B0aEU8+6IxjgEDdBQqtJ
# hsNhFHpOjce3YfqFAimBNuhdike207LksEMsp8hNXyvFNE/BXAnIHpVRGW5gX2JO
# USOrsY6/OPpXm7OMSoeDFYdBViWlGpMWsrZUqP2hKy+fbqoS24nMRD8YV6gZurEX
# SXJquGTEiP/Q7PWmxzetRboHViMqcFIe4uXAvo0j1mbLJjIzUAA9cq7+jXBBRXiq
# P/1FfWPTt1zpK6k2dhKuye0kPbg3fK2nB0NZTJFWIiXVtpcSlkgetQO4Mkb7Wq0G
# p7A3CY7he9YY+7IQhDlouYbSdN039tJ/sdyckLtjy3jpw1nvtAbdH/bzXPUZ2zEH
# xmVMNZSn6mv+TCQ8e44U4+3iQ8RHE5MKyHf9Rb8kknvYAecOka4vfgnwj6g22F22
# 2Qisfc9JgZ9/KUeGeyv+bWivHwZppH9Vs3xqAq/n4DCBYijC1kUn7P9HMZxSTub+
# //upk9Sc5SkvXvsjkisDS0nHvWXaRdkZGPDJWjK7C+S9sh8C8ciWheV/4VTTSNHx
# 7wEnmnZYF75/RNc83OmCUVa4Cjksu77fn9QrgJqbMDrcszHJR5PsJeMDyIdF7YcK
# ByA8G+fD5VkC
# SIG # End signature block
