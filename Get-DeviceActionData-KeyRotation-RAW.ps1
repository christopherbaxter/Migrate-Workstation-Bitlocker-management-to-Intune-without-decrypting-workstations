<#
.SYNOPSIS
    Report on status of Bitlocker Recovery key rotation from Intune - via MS Graph API.
.DESCRIPTION
    This script will extract the results of the recovery key rotation in bulk from the MSGraph API
.PARAMETER TenantID
    Specify the Azure AD tenant ID.
.PARAMETER ClientID
    Specify the service principal, also known as app registration, Client ID (also known as Application ID).
.PARAMETER State
    Specify -TenantID and -ClientID, or edit the script and add hard code it
.EXAMPLE
    # Generate a report of the status of Bitlocker Key rotation requests, for all devices in estate: NOTE: Windows 10 Build 1909 and above required
    .\Get-DeviceActionData-KeyRotation.ps1 -TenantID "<tenant_id>" -ClientID "<client_id>"
.NOTES
    FileName:    Get-DeviceActionData-KeyRotation.ps1
    Author:      Christopher Baxter
    Contact:     GitHub - https://github.com/christopherbaxter
    Created:     2021-11-01
    Updated:     2022-02-03

    Depending on the size of your estate and the speed of your connection, this script may take a significant amount of time to run. Make sure that your elevated rights in AzureAD\Intune have an appropriate amount of time for this to function.

    This code will likely be able to get any device action data, if you are looking for something else, you will need to just know what you are looking for.

#>
#Requires -Modules "MSAL.PS","PoshRSJob","JoinModule"
[CmdletBinding(SupportsShouldProcess = $TRUE)]
param(
    #PLEASE make sure you have specified your details below, else edit this and use the switches\variables in command line.
    [parameter(Mandatory = $TRUE, HelpMessage = "Specify the Azure AD tenant ID.")]
    [ValidateNotNullOrEmpty()]
    [string]$TenantID,

    [parameter(Mandatory = $TRUE, HelpMessage = "Specify the service principal, also known as app registration, Client ID (also known as Application ID).")]
    [ValidateNotNullOrEmpty()]
    [string]$ClientID
)
Begin {}
Process {
    
    #############################################################################################################################################
    # Functions
    #############################################################################################################################################

    function Invoke-MSGraphOperation {
        <#
        .SYNOPSIS
            Perform a specific call to Graph API, either as GET, POST, PATCH or DELETE methods.
            
        .DESCRIPTION
            Perform a specific call to Graph API, either as GET, POST, PATCH or DELETE methods.
            This function handles nextLink objects including throttling based on retry-after value from Graph response.
            
        .PARAMETER Get
            Switch parameter used to specify the method operation as 'GET'.
            
        .PARAMETER Post
            Switch parameter used to specify the method operation as 'POST'.
            
        .PARAMETER Patch
            Switch parameter used to specify the method operation as 'PATCH'.
            
        .PARAMETER Put
            Switch parameter used to specify the method operation as 'PUT'.
            
        .PARAMETER Delete
            Switch parameter used to specify the method operation as 'DELETE'.
            
        .PARAMETER Resource
            Specify the full resource path, e.g. deviceManagement/auditEvents.
            
        .PARAMETER Headers
            Specify a hash-table as the header containing minimum the authentication token.
            
        .PARAMETER Body
            Specify the body construct.
            
        .PARAMETER APIVersion
            Specify to use either 'Beta' or 'v1.0' API version.
            
        .PARAMETER ContentType
            Specify the content type for the graph request.
            
        .NOTES
            Author:      Nickolaj Andersen & Jan Ketil Skanke & (very little) Christopher Baxter
            Contact:     @JankeSkanke @NickolajA
            Created:     2020-10-11
            Updated:     2020-11-11
    
            Version history:
            1.0.0 - (2020-10-11) Function created
            1.0.1 - (2020-11-11) Tested in larger environments with 100K+ resources, made small changes to nextLink handling
            1.0.2 - (2020-12-04) Added support for testing if authentication token has expired, call Get-MsalToken to refresh. This version and onwards now requires the MSAL.PS module
            1.0.3.Custom - (2020-12-20) Added aditional error handling. Not complete, but more will be added as needed. Christopher Baxter
        #>
        param(
            [parameter(Mandatory = $tRUE, ParameterSetName = "GET", HelpMessage = "Switch parameter used to specify the method operation as 'GET'.")]
            [switch]$Get,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST", HelpMessage = "Switch parameter used to specify the method operation as 'POST'.")]
            [switch]$Post,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH", HelpMessage = "Switch parameter used to specify the method operation as 'PATCH'.")]
            [switch]$Patch,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT", HelpMessage = "Switch parameter used to specify the method operation as 'PUT'.")]
            [switch]$Put,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "DELETE", HelpMessage = "Switch parameter used to specify the method operation as 'DELETE'.")]
            [switch]$Delete,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "GET", HelpMessage = "Specify the full resource path, e.g. deviceManagement/auditEvents.")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [string]$Resource,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "GET", HelpMessage = "Specify a hash-table as the header containing minimum the authentication token.")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [System.Collections.Hashtable]$Headers,
    
            [parameter(Mandatory = $tRUE, ParameterSetName = "POST", HelpMessage = "Specify the body construct.")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $tRUE, ParameterSetName = "PUT")]
            [ValidateNotNullOrEmpty()]
            [System.Object]$Body,
    
            [parameter(Mandatory = $fALSE, ParameterSetName = "GET", HelpMessage = "Specify to use either 'Beta' or 'v1.0' API version.")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "POST")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [ValidateSet("Beta", "v1.0")]
            [string]$APIVersion = "v1.0",
    
            [parameter(Mandatory = $fALSE, ParameterSetName = "GET", HelpMessage = "Specify the content type for the graph request.")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "POST")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "PATCH")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "PUT")]
            [parameter(Mandatory = $fALSE, ParameterSetName = "DELETE")]
            [ValidateNotNullOrEmpty()]
            [ValidateSet("application/json", "image/png")]
            [string]$ContentType = "application/json"
        )
        Begin {
            # Construct list as return value for handling both single and multiple instances in response from call
            $GraphResponseList = New-Object -TypeName "System.Collections.ArrayList"
            $Runcount = 1
            
            # Construct full URI
            $GraphURI = "https://graph.microsoft.com/$($APIVersion)/$($Resource)"
            #Write-Verbose -Message "$($PSCmdlet.ParameterSetName) $($GraphURI)"
        }
        Process {
            # Call Graph API and get JSON response
            do {
                try {
                    # Determine the current time in UTC
                    $UTCDateTime = (Get-Date).ToUniversalTime()
    
                    # Determine the token expiration count as minutes
                    $TokenExpireMins = ([datetime]$Headers["ExpiresOn"] - $UTCDateTime).Minutes
    
                    # Attempt to retrieve a refresh token when token expiration count is less than or equal to 10
                    if ($TokenExpireMins -le 10) {
                        #Write-Verbose -Message "Existing token found but has expired, requesting a new token"
                        #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
                        try { $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -ForceRefresh -Silent -ErrorAction Stop }
                        catch { $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -ErrorAction Stop }
                        if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
                        $Headers = New-AuthenticationHeader -AccessToken $AccessToken

                        #$AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -Silent -ForceRefresh
                        #$Headers = New-AuthenticationHeader -AccessToken $AccessToken
                    }
    
                    # Construct table of default request parameters
                    $RequestParams = @{
                        "Uri"         = $GraphURI
                        "Headers"     = $Headers
                        "Method"      = $PSCmdlet.ParameterSetName
                        "ErrorAction" = "Stop"
                        "Verbose"     = $VerbosePreference
                    }
    
                    switch ($PSCmdlet.ParameterSetName) {
                        "POST" {
                            $RequestParams.Add("Body", $Body)
                            $RequestParams.Add("ContentType", $ContentType)
                        }
                        "PATCH" {
                            $RequestParams.Add("Body", $Body)
                            $RequestParams.Add("ContentType", $ContentType)
                        }
                        "PUT" {
                            $RequestParams.Add("Body", $Body)
                            $RequestParams.Add("ContentType", $ContentType)
                        }
                    }
    
                    # Invoke Graph request
                    $GraphResponse = Invoke-RestMethod @RequestParams
                    
                    # Handle paging in response
                    if ($GraphResponse.'@odata.nextLink') {
                        $GraphResponseList.AddRange($GraphResponse.value) | Out-Null
                        $GraphURI = $GraphResponse.'@odata.nextLink'
                        #Write-Verbose -Message "NextLink: $($GraphURI)"
                    }
                    else {
                        # NextLink from response was null, assuming last page but also handle if a single instance is returned
                        if (-not([string]::IsNullOrEmpty($GraphResponse.value))) {
                            $GraphResponseList.AddRange($GraphResponse.value) | Out-Null
                        }
                        else {
                            $GraphResponseList.Add($GraphResponse) | Out-Null
                        }
                        
                        # Set graph response as handled and stop processing loop
                        $GraphResponseProcess = $fALSE
                    }
                }
                catch [System.Exception] {
                    $ExceptionItem = $PSItem
                    if ($ExceptionItem.Exception.Response.StatusCode -like "429") {
                        # Detected throttling based from response status code
                        $RetryInsecond = $ExceptionItem.Exception.Response.Headers["Retry-After"]
    
                        # Wait for given period of time specified in response headers
                        #Write-Verbose -Message "Graph is throttling the request, will retry in $($RetryInsecond) seconds"
                        Start-Sleep -second $RetryInsecond
                    }
                    elseif ($ExceptionItem.Exception.Response.StatusCode -like "Unauthorized") {
                        #Write-Verbose -Message "Your Account does not have the relevent privilege to read this data. Please Elevate your account or contact the administrator"
                        $Script:PIMExpired = $tRUE
                        $GraphResponseProcess = $fALSE
                    }
                    elseif ($ExceptionItem.Exception.Response.StatusCode -like "GatewayTimeout") {
                        # Detected Gateway Timeout
                        $RetryInsecond = 30
    
                        # Wait for given period of time specified in response headers
                        #Write-Verbose -Message "Gateway Timeout detected, will retry in $($RetryInsecond) seconds"
                        Start-Sleep -second $RetryInsecond
                    }
                    elseif ($ExceptionItem.Exception.Response.StatusCode -like "NotFound") {
                        #Write-Verbose -Message "The Device data could not be found"
                        $Script:StatusResult = $ExceptionItem.Exception.Response.StatusCode
                        $GraphResponseProcess = $fALSE
                    }
                    elseif ($PSItem.Exception.Message -like "*Invalid JSON primitive*") {
                        $Runcount++
                        if ($Runcount -eq 5) {
                            if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
                            try { $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -ForceRefresh -Silent -ErrorAction Stop }
                            catch { $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -ErrorAction Stop }
                            if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
                            $Headers = New-AuthenticationHeader -AccessToken $AccessToken
                        }
                        if ($Runcount -ge 10) {
                            #Write-Verbose -Message "An Unrecoverable Error occured - Error: Invalid JSON primitive"
                            $GraphResponseProcess = $fALSE
                        }
                        $RetryInsecond = 5
                        #Write-Verbose -Message "Invalid JSON Primitive detected, Trying again in $($RetryInsecond) seconds"
                        Start-Sleep -second $RetryInsecond
                        
                    }
                    else {
                        try {
                            # Read the response stream
                            $StreamReader = New-Object -TypeName "System.IO.StreamReader" -ArgumentList @($ExceptionItem.Exception.Response.GetResponseStream())
                            $StreamReader.BaseStream.Position = 0
                            $StreamReader.DiscardBufferedData()
                            $ResponseBody = ($StreamReader.ReadToEnd() | ConvertFrom-Json)
                            
                            switch ($PSCmdlet.ParameterSetName) {
                                "GET" {
                                    # Output warning message that the request failed with error message description from response stream
                                    Write-Warning -Message "Graph request failed with status code $($ExceptionItem.Exception.Response.StatusCode). Error message: $($ResponseBody.error.message)"
    
                                    # Set graph response as handled and stop processing loop
                                    $GraphResponseProcess = $fALSE
                                }
                                default {
                                    # Construct new custom error record
                                    $SystemException = New-Object -TypeName "System.Management.Automation.RuntimeException" -ArgumentList ("{0}: {1}" -f $ResponseBody.error.code, $ResponseBody.error.message)
                                    $ErrorRecord = New-Object -TypeName "System.Management.Automation.ErrorRecord" -ArgumentList @($SystemException, $ErrorID, [System.Management.Automation.ErrorCategory]::NotImplemented, [string]::Empty)
    
                                    # Throw a terminating custom error record
                                    $PSCmdlet.ThrowTerminatingError($ErrorRecord)
                                }
                            }
    
                            # Set graph response as handled and stop processing loop
                            $GraphResponseProcess = $fALSE
                        }
                        catch [System.Exception] {
                            if ($PSItem.Exception.Message -like "*Invalid JSON primitive*") {
                                $Runcount++
                                if ($Runcount -ge 10) {
                                    #Write-Verbose -Message "An Unrecoverable Error occured - Error: Invalid JSON primitive"
                                    $GraphResponseProcess = $fALSE
                                }
                                $RetryInsecond = 5
                                #Write-Verbose -Message "Invalid JSON Primitive detected, Trying again in $($RetryInsecond) seconds"
                                Start-Sleep -second $RetryInsecond
                                
                            }
                            else {
                                Write-Warning -Message "Unhandled error occurred in function. Error message: $($PSItem.Exception.Message)"
    
                                # Set graph response as handled and stop processing loop
                                $GraphResponseProcess = $fALSE
                            }
                        }
                    }
                }
            }
            until ($GraphResponseProcess -eq $fALSE)
    
            # Handle return value
            return $GraphResponseList
            
        }
    }

    function New-AuthenticationHeader {
        <#
        .SYNOPSIS
            Construct a required header hash-table based on the access token from Get-MsalToken cmdlet.
        .DESCRIPTION
            Construct a required header hash-table based on the access token from Get-MsalToken cmdlet.
        .PARAMETER AccessToken
            Pass the AuthenticationResult object returned from Get-MsalToken cmdlet.
        .NOTES
            Author:      Nickolaj Andersen
            Contact:     @NickolajA
            Created:     2020-12-04
            Updated:     2020-12-04
            Version history:
            1.0.0 - (2020-12-04) Script created
        #>
        param(
            [parameter(Mandatory = $tRUE, HelpMessage = "Pass the AuthenticationResult object returned from Get-MsalToken cmdlet.")]
            [ValidateNotNullOrEmpty()]
            [Microsoft.Identity.Client.AuthenticationResult]$AccessToken
        )
        Process {
            # Construct default header parameters
            $AuthenticationHeader = @{
                "Content-Type"  = "application/json"
                "Authorization" = $AccessToken.CreateAuthorizationHeader()
                "ExpiresOn"     = $AccessToken.ExpiresOn.LocalDateTime
            }
    
            # Amend header with additional required parameters for bitLocker/recoveryKeys resource query
            $AuthenticationHeader.Add("ocp-client-name", "My App")
            $AuthenticationHeader.Add("ocp-client-version", "1.2")
    
            # Handle return value
            return $AuthenticationHeader
        }
    }

    function Clear-ResourceEnvironment {
        # Clear any PowerShell sessions created
        Get-PSSession | Remove-PSSession

        # Release an COM object created
        #$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Shell)

        # Perform garbage collection on session resources 
        [System.GC]::Collect()         
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        # Remove any custom varialbes created
        #Get-Variable -Name MyShell -ErrorAction Silently$VerbosePreference | Remove-Variable
    
    }
    
    #############################################################################################################################################
    # Variables
    #############################################################################################################################################

    $Script:PIMExpired = $null
    $FilePath = "C:\Temp\BitlockerKeyEscrow\"
    $CSVFileName = "DevActionArray"
    $FileDate = Get-Date -Format 'yyyy_MM_dd'
    $RotationDataExportFile = "$($FilePath)InterimFiles\$($CSVFileName) - $($FileDate).csv"
    $ScriptStartTime = Get-Date -Format 'yyyy-MM-dd HH:mm'
    [string]$Resource = "deviceManagement/managedDevices"
    
    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

    #############################################################################################################################################
    # Get Authentication token and Authentication Header
    #############################################################################################################################################

    #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
    try { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop }
    catch { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ErrorAction Stop }
    if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
    $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken

    #############################################################################################################################################
    # Extract list of Intune Device IDs from MSGraph - Or supply your own list of IntuneDeviceIDs
    #############################################################################################################################################

    Write-Host "Extracting the data from MS Graph Intune. Expected runtime is 4 minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $DevActionDevList = [System.Collections.ArrayList]::new()
    $DevActionDevList = [System.Collections.ArrayList]@(Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'" -Headers $AuthenticationHeader -Verbose | Where-Object { $_.azureADDeviceId -ne "00000000-0000-0000-0000-000000000000" } | Select-Object id | Sort-Object id)
    Write-Host "Collected Data for $($DevActionDevList.count) objects from MS Graph Intune - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    #$RotationDataExportFile = "$($FilePath)InterimFiles\Targeted - $($CSVFileName) - $($FileDate).csv"
    #$InputFile = ""$($FilePath)InputFiles\TargetedIntuneDeviceIDs.csv"
    #$DevActionDevList = @(Import-Csv -Path $InputFile)
    #$DevActionDevList.count

    #############################################################################################################################################
    # Split the IntuneDeviceID array into smaller chunks (parts) and set Throttle limit for parallel processing
    #############################################################################################################################################

    Write-Host "Extracting list of Devices from Intune Device extraction list and splitting array into 30 parts for processing - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    $RawExtract = [System.Collections.ArrayList]::new()
    [int]$parts = 30
    $PartSize = [Math]::Ceiling($DevActionDevList.count / $parts)
    $SplitDevicelist = @()
    for ($i = 1; $i -le $parts; $i++) {
        $start = (($i - 1) * $PartSize)
        $end = (($i) * $PartSize) - 1
        if ($end -ge $DevActionDevList.count) {
            $end = $DevActionDevList.count
        }
        $SplitDevicelist += , @($DevActionDevList[$start..$end])
    }

    Write-Host "Completed extracting list of Devices from Intune Device extraction list - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $ThrottleLimit = 30
    
    #############################################################################################################################################
    # Specify the ScriptBlock for the PoshRSJob
    #############################################################################################################################################

    $ScriptBlock = {
        param ([System.Collections.Hashtable]$AuthenticationHeader, [string]$Resource, [string]$cID, [string]$tID, [string]$APIVersion)
        if (-not($Runcount)) {
            $Runcount = 0
        }
        $Runcount++
        if ($Runcount -ge 1000) {

            #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
            try { $AccessToken = Get-MsalToken -TenantId $tID -ClientId $cID -ForceRefresh -Silent -ErrorAction Stop }
            catch { $AccessToken = Get-MsalToken -TenantId $tID -ClientId $cID -ErrorAction Stop }
            if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
            $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken

            #$AccessToken = Get-MsalToken -TenantId $tID -ClientId $cID -ForceRefresh -Silent -ErrorAction Stop
            #$AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
            $Runcount = 0 
        }
        $GraphURI = "https://graph.microsoft.com/$($APIVersion)/$($Resource)/$($_)?`$select=id,deviceActionResults,aadRegistered,autopilotEnrolled"
        $RequestParams = @{
            "Uri"         = $GraphURI
            "Headers"     = $AuthenticationHeader
            "Method"      = "Get"
            "ErrorAction" = "Stop"
            "Verbose"     = $VerbosePreference
        }

        # Invoke Graph request
        try {
            Invoke-RestMethod @RequestParams
        }
        catch {
            #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
            try { $AccessToken = Get-MsalToken -TenantId $tID -ClientId $cID -ForceRefresh -Silent -ErrorAction Stop }
            catch { $AccessToken = Get-MsalToken -TenantId $tID -ClientId $cID -ErrorAction Stop }
            if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
            $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
            Invoke-RestMethod @RequestParams
        }
    }

    #############################################################################################################################################
    # Foreach loop to run through all items in each 'split' array
    #############################################################################################################################################

    Write-Host "Extracting Key rotation status information from MS Graph. This is done in 30 rounds - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $Counter = 0
    $RawExtract = [System.Collections.ArrayList]@(Foreach ($i in $SplitDevicelist) {
            # Get authentication token
            #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
            try { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop }
            catch { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ErrorAction Stop }
            if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
            $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
            $Counter++

            Write-Host "Extraction Round number $($Counter) of $($ThrottleLimit)"

            $i.id | Start-RSJob -ScriptBlock $ScriptBlock -Throttle $ThrottleLimit -ArgumentList $AuthenticationHeader, $Resource, $ClientID, $TenantID, "Beta" | Wait-RSJob -ShowProgress | Receive-RSJob
            Get-RSJob | Remove-RSJob -Force
        }
    )

    #############################################################################################################################################
    # Process the results and create an array with the failed extractions for retry
    #############################################################################################################################################

    Write-Host "Completed extracting Key rotation status information from MS Graph - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    Write-Host "Joining Intune Device List with Key Rotation Extraction and selecting devices that appear to have failed extraction - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $ExtractionCheck = [System.Collections.ArrayList]@($DevActionDevList | LeftJoin-Object $RawExtract -On id | Sort-Object id)
    Remove-Variable -Name DevActionDevList -Force
    Remove-Variable -Name SplitDevicelist -Force
    Clear-ResourceEnvironment
        
    # Processing Time is 10 Seconds
    $FailedExtractList = [System.Collections.ArrayList]@($ExtractionCheck | Where-Object { ($_.aadRegistered -like $null) -or ($_.autopilotEnrolled -like $null) } | Select-Object id )
    Remove-Variable -Name ExtractionCheck -Force
    Clear-ResourceEnvironment
        
    Write-Host "Completed joining Intune Device List with Key Rotation Extraction Array - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    #############################################################################################################################################
    # Split the failed extraction array into smaller chunks (parts) and set Throttle limit for parallel processing
    #############################################################################################################################################

    Write-Host "Splitting remaining devices for Key Rotation Extraction into 10 Parts - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    [int]$parts = 10
    $PartSize = [Math]::Ceiling($FailedExtractList.count / $parts)
    $FailedSplitList = @()
    for ($i = 1; $i -le $parts; $i++) {
        $start = (($i - 1) * $PartSize)
        $end = (($i) * $PartSize) - 1
        if ($end -ge $FailedExtractList.count) {
            $end = $FailedExtractList.count
        }
        $FailedSplitList += , @($FailedExtractList[$start..$end])
    }
    Write-Host "Completed splitting remaining devices for Key Rotation Extraction - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $ThrottleLimit = 30

    #############################################################################################################################################
    # Foreach loop to run through all items in each 'split' array - more of the same above
    #############################################################################################################################################

    Write-Host "Extracting remaining devices' Key Rotation data in 10 Rounds - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    # Processing Time is 11 Minutes
    $Counter = 0
    $FailedRawExtract = [System.Collections.ArrayList]@(Foreach ($item in $FailedSplitList) {
            # Get authentication token
            #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
            try { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop }
            catch { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ErrorAction Stop }
            if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
            $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken

            # Construct authentication header
            if ($AuthenticationHeader) {
                Remove-Variable -Name AuthenticationHeader -Force
            }
            $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
            $Counter++
                
            Write-Host "Extraction Round number $($Counter)"

            $item.id | Start-RSJob -ScriptBlock $ScriptBlock -Throttle $ThrottleLimit -ArgumentList $AuthenticationHeader, $Resource, $ClientID, $TenantID, "Beta" | Wait-RSJob -ShowProgress | Receive-RSJob
            Get-RSJob | Remove-RSJob -Force
        }
    )

    Remove-Variable -Name FailedSplitList -Force
    Clear-ResourceEnvironment

    Write-Host "Completed extracting remaining devices' Key Rotation data - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    #############################################################################################################################################
    # Generate the report and export it
    #############################################################################################################################################

    $ConsolidatedRawExtract = [System.Collections.ArrayList]@($RawExtract + $FailedRawExtract | Sort-Object id)
    Remove-Variable -Name FailedRawExtract -Force
    Remove-Variable -Name RawExtract -Force
    Clear-ResourceEnvironment

    # I used this little section for logic testing.

    #$ConsolidatedRawExtract | Export-csv -path "$($FilePath)InterimFiles\DevActionRAWArray.csv" -Delimiter ";" -NoTypeInformation

    #if ($ConsolidatedRawExtract.count -lt 1){
    #    $ConsolidatedRawExtract = [System.Collections.ArrayList]@(Import-Csv -Path "$($FilePath)InterimFiles\DevActionRAWArray.csv" -Delimiter ";")
    #}

    Write-Host "Processing Key rotation extraction Data for reporting - Processing Time is 43 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $DevActionArray = [System.Collections.ArrayList]@($ConsolidatedRawExtract | Select-Object @{Name = "IntuneDeviceID"; Expression = { $_.id } }, @{Name = "KeyRotationResult"; Expression = { if ($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object actionState ) { ($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object actionState).actionState } else { "NoRotationRequest" } } }, @{Name = "KeyRotationRequestDate"; Expression = { if ($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object startDateTime ) { (Get-Date -Date ($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object startDateTime).startDateTime -Format "yyyy/MM/dd HH:mm") } else { "NoRotationRequest" } } }, @{Name = "KeyRotationDate"; Expression = { if ($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object lastUpdatedDateTime ) { (Get-Date -Date ($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object lastUpdatedDateTime).lastUpdatedDateTime -Format "yyyy/MM/dd HH:mm") } else { "NoRotationRequest" } } }, @{Name = "KeyRotationError"; Expression = { if (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorcode -like $null) { "NoRotationRequest" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2147023728") { "0x80070490" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272166") { "0x803100DA" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2147467259") { "0x80004005" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272310") { "0x8031004A" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2147024809") { "0x80070057" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272159") { "0x803100E1" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272165") { "0x803100DB" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272295") { "0x80310059" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272384") { "0x80310000" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272376") { "0x80310008" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2147418113") { "0x8000FFFF" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272366") { "0x80310012" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272339") { "0x8031002D" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2146893783") { "0x80090029" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "0") { "0" } else { ($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode } } }, @{Name = "KeyRotationErrorDescription"; Expression = { if (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object actionState).actionState -like "Done" ) { "None" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2147023728") { "Element not Found" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272166") { "BitLocker recovery password rotation cannot be performed because backup policy for BitLocker recovery information is not set to required for the OS drive." }elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2147467259") { "General Failure." }elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272310") { "BitLocker Drive Encryption cannot be used because critical BitLocker system files are missing or corrupted. Use Windows Startup Repair to restore these files to your computer." }elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2147024809") { "One or more arguments are invalid - The parameter is incorrect." }elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272159") { "BitLocker recovery key backup endpoint is busy and cannot perform requested operation. Please retry after sometime." }elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272165") { "BitLocker recovery password rotation cannot be performed because backup policy for BitLocker recovery information is not set to required for fixed data drives." }elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272295") { "BitLocker Drive Encryption is already performing an operation on this drive. Please complete all operations before continuing." }elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272384") { "This drive is locked by BitLocker Drive Encryption. You must unlock this drive from Control Panel." }elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272376") { "BitLocker Drive Encryption is not enabled on this drive. Turn on BitLocker." }elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2147418113") { "Catastrophic failure" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272366") { "The drive cannot be encrypted because it contains system boot information. Create a separate partition for use as the system drive that contains the boot information and a second partition for use as the operating system drive and then encrypt the operating system drive." } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2144272339") { "The drive encryption algorithm and key cannot be set on a previously encrypted drive. To encrypt this drive with BitLocker Drive Encryption, remove the previous encryption and then turn on BitLocker." } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "-2146893783") { "The requested operation is not supported." } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "0") { if (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object actionState).actionState -match "Pending" ) { "Pending" } elseif (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object actionState).actionState -match "done" ) { "Key Rotation Successful." } elseif ({ (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object actionState).actionState -match "Failed") -and (($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object errorCode).errorCode -match "0") }) { ($_.deviceActionResults | Where-Object { $_.actionName -like "*BitLocker*" } | Select-Object actionState).actionState } } else { "Not Yet Determined." } } } | Sort-Object IntuneDeviceID)
    Remove-Variable -Name ConsolidatedRawExtract -Force
    Clear-ResourceEnvironment
    
    Write-Host "Completed Extracting Key Rotation Status Data from MS Graph - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
        
    $DevActionArray | Export-Csv -Path $RotationDataExportFile -Delimiter ";" -NoTypeInformation
        
    Remove-Variable -Name DevActionArray -Force
    Clear-ResourceEnvironment
        
    Write-Host "ALL DONE!!! - Script Start Time: $($ScriptStartTime) and Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green -BackgroundColor Red
}