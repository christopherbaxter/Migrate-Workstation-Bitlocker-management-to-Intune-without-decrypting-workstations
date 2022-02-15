<#
.SYNOPSIS
    Get the BitLocker recovery key presence for Intune managed devices.
.DESCRIPTION
    This script retrieves the BitLocker recovery key presence for Intune managed devices.
.PARAMETER TenantID
    Specify the Azure AD tenant ID.
.PARAMETER ClientID
    Specify the service principal, also known as app registration, Client ID (also known as Application ID).
.PARAMETER State
    Specify either 'Present' or 'NotPresent'. (no longer needed). Just use -Verbose
.EXAMPLE
    # Retrieve a list of Intune managed devices that have a BitLocker recovery key associated on the Azure AD device object:
    .\Get-IntuneManagedDeviceBitLockerKeyPresence.ps1 -TenantID "<tenant_id>" -ClientID "<client_id>"
    # Retrieve a list of Intune managed devices that doesn't have a BitLocker recovery key associated on the Azure AD device object:
    .\Get-IntuneManagedDeviceBitLockerKeyPresence.ps1 -TenantID "<tenant_id>" -ClientID "<client_id>"
.NOTES
    FileName:    Get-IntuneManagedDeviceBitLockerKeyPresence.ps1
    Author:      Nickolaj Andersen
    Contact:     @NickolajA
    Created:     2020-12-04
    Updated:     2020-12-04

    Hacked up and manipulated to work in massive organisations. RAM util is high, I have tried to manage this.
    I work in a large enterprise. I have found that scripting for a normal organisation is completely different to scripting for a large enterprise with over 100000 users.
    The biggest issue is that the scripts seldom scale out to handle more than 1mil objects. Suddenly the scrips start to fail randomly, timeouts become a problem. RAM shortages start to creep in
    and the built-in security mechanisms start to be a problem (schedules daily password changes, time limits on permissions (Time based permissions)).
    I have spent a lot of time on getting around these limitations, which you will see in the script below. I have made notes to explain what I have done, and why.
    I wish to thank Anders Ahl and Nickolaj Andersen as well as the authors of the functions I have used below, without these guys, this would not exist.
    
    You would still be left explaining to your management that you cannot get the report out. Currently, the script takes about 7hrs to complete, make sure you have a window long enough to cope with this. - Christopher Baxter.
    The line above is not really a concern any longer. All the MS Graph and AzureAD extracts have been moved to run first. The script then processes that data, reducing the time required for the permissions 'window'.

    Version history:
    1.0.0 - (2020-12-04) Script created by Nickolaj Andersen
    2.0.0 - (2021-01-30) Script tweaked, expanded, manipulated, debugged, scaled and skookumized to be able to handle massive organisations by Christopher Baxter.
    3.0.0 - (2021-05-07) Script completed. This now functions. Christopher Baxter.
    4.0.0 - (2021-07-01) Script reworked and optimised. Christopher Baxter.
    4.0.1 - (Yup, I lost track) Many performance improvements realised. This script will extract the data it needs, then export the data to a file, clear array variable and clear the memory.
            This was done in an effort to limit the amount of resources needed to run the script. Depending on the size of the environment, without this, RAM utilisation can easily climb beyond what is available. The script was written on a machine with 32Gb RAM, with no other tasks being performed.
            The script suffered constant crashes and severe performance degradation prior to this process being implemented. Your Mileage may vary.

#>
#Requires -Modules "MSAL.PS","ActiveDirectory","AzureAD","ImportExcel","JoinModule","PSReadline"
[CmdletBinding(SupportsShouldProcess = $TRUE)]
param(
    #PLEASE make sure you have specified your details below, else edit this and use the switches\variables in command line.
    [parameter(Mandatory = $TRUE, HelpMessage = "Specify the Azure AD tenant ID.")]
    [ValidateNotNullOrEmpty()]
    #[string]$TenantID = "", # Populate this with your TenantID, this will then allow the script to run without asking for the details
    [string]$TenantID,

    [parameter(Mandatory = $TRUE, HelpMessage = "Specify the service principal, also known as app registration, Client ID (also known as Application ID).")]
    [ValidateNotNullOrEmpty()]
    #[string]$ClientID = "" # Populate this with your ClientID\ApplicationID of your Service Principal, this will then allow the script to run without asking for the details
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
                        $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -Silent -ForceRefresh
                        $Headers = New-AuthenticationHeader -AccessToken $AccessToken
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
                            $AccessToken = Get-MsalToken -TenantId $Script:TenantID -ClientId $Script:ClientID -Silent -ForceRefresh
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
    $BitlockerKeyEscrowReport = $null
    $FileDate = Get-Date -Format 'yyyy_MM_dd'
    $ExcelFileName = "BitlockerBlendedReport"
    $FilePath = "C:\Temp\BitlockerKeyEscrow\"
    $ConsolidatedReportFileName = "$($ExcelFileName)_$($FileDate).xlsx"
    $ConsolidatedReportExport = "$($FilePath)$($ConsolidatedReportFileName)"
    $SCCMInputFiles = "$($FilePath)InputFiles"
    $InterimFileLocation = "$($FilePath)InterimFiles"

    $ADForest = (Get-ADForest).RootDomain
    $DomainTargets = (Get-ADForest -Identity $ADForest).Domains
    
    $ScriptStartTime = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $StaleDate = (Get-Date).AddDays(-90)
    [string]$Resource = "deviceManagement/managedDevices"
    
    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

    #############################################################################################################################################
    # Get Authentication Token and Authentication Header
    #############################################################################################################################################

    Clear-ResourceEnvironment

    #if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
    try { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop }
    catch { $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ErrorAction Stop }
    if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
    $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken
        
    #############################################################################################################################################
    # AzureAD Device Data Extraction
    #############################################################################################################################################

    Connect-AzureAD
    Write-Host "Extracting Data from AzureAD. Expected runtime is 7 minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $AzureADDevices = [System.Collections.ArrayList]@(Get-AzureADDevice -All:$TRUE | Where-Object { $_.DeviceOSType -like "*Windows*" } | Select-Object @{Name = "AzureADDeviceID"; Expression = { $_.DeviceId.toString() } }, @{Name = "ObjectID"; Expression = { $_.ObjectID.toString() } }, AccountEnabled, @{Name = "AADApproximateLastLogonTimeStamp"; Expression = { (Get-Date -Date $_.ApproximateLastLogonTimeStamp -Format 'yyyy/MM/dd HH:mm') } }, @{Name = "AADDisplayName"; Expression = { $_.DisplayName } }, @{Name = "AADLastDirSyncTime"; Expression = { (Get-Date -Date $_.LastDirSyncTime -Format 'yyyy/MM/dd HH:mm') } }, ProfileType, @{Name = "AADSTALE"; Expression = { if ($_.ApproximateLastLogonTimeStamp -le $StaleDate) { "TRUE" } elseif ($_.ApproximateLastLogonTimeStamp -gt $StaleDate) { "FALSE" } else { "NoLoginDateFound" } } } | Sort-Object azureADDeviceId )
    Write-Host "Collected Data for $($AzureADDevices.count) objects from AzureAD - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    Disconnect-AzureAD

    $AzureADDevices | Export-Csv -Path "$($InterimFileLocation)\AzureADExtract.csv" -Delimiter ";" -NoTypeInformation
        
    Remove-Variable -Name AzureADDevices -Force
    Clear-ResourceEnvironment

    #############################################################################################################################################
    # Intune Managed Device Data Extraction
    #############################################################################################################################################

    if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
    $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -Silent -ErrorAction Stop
    if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
    $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken

    Write-Host "Extracting the data from MS Graph Intune. Expected runtime is 4 minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $IntuneInterimArray = [System.Collections.ArrayList]::new()
    $IntuneInterimArray = [System.Collections.ArrayList]@(Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'" -Headers $AuthenticationHeader -Verbose | Where-Object { $_.azureADDeviceId -ne "00000000-0000-0000-0000-000000000000" } | Select-Object @{Name = "azureADDeviceId"; Expression = { $_.azureADDeviceId.toString() } }, @{Name = "IntuneDeviceID"; Expression = { $_.id.ToString() } }, @{Name = "MSGraphDeviceName"; Expression = { $_.deviceName } }, deviceEnrollmentType, azureADRegistered, @{Name = "enrolledDateTime"; Expression = { (Get-Date -Date $_.enrolledDateTime -Format "yyyy/MM/dd HH:mm") } }, @{Name = "MSGraphlastSyncDateTime"; Expression = { (Get-Date -Date $_.lastSyncDateTime -Format "yyyy/MM/dd HH:mm") } }, operatingSystem, osVersion, managementAgent, deviceRegistrationState, complianceState, @{Name = "UserUPN"; Expression = { $_.userPrincipalName } }, @{Name = "DeviceManufacturer"; Expression = { $_.manufacturer } }, @{Name = "DeviceModel"; Expression = { $_.model } }, @{Name = "DeviceSN"; Expression = { $_.serialNumber.ToString() } }, managedDeviceName, @{Name = "MSGraphEncryptionState"; Expression = { $_.isEncrypted } }, aadRegistered, autopilotEnrolled, joinType | Sort-Object IntuneDeviceID)

    Write-Host "Collected Data for $($IntuneInterimArray.count) objects from MS Graph Intune - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    $IntuneInterimArray | Export-Csv -Path "$($InterimFileLocation)\IntuneInterimArray.csv" -Delimiter ";" -NoTypeInformation
        
    Remove-Variable -Name IntuneInterimArray -Force
    Clear-ResourceEnvironment

    #############################################################################################################################################
    # Intune Encryption Reporting Data Extraction
    #############################################################################################################################################

    if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
    $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -ErrorAction Stop
    if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
    $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken

    Write-Host "Extracting the data from MS Graph Encryption Status. Expected runtime is 10 minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    $RAWIntuneEncryptionReportData = [System.Collections.ArrayList]::new()
    $RAWIntuneEncryptionReportData = [System.Collections.ArrayList]@(Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "deviceManagement/managedDeviceEncryptionStates" -Headers $AuthenticationHeader -Verbose | Where-Object { $_.azureADDeviceId -ne "00000000-0000-0000-0000-000000000000" } | Select-Object @{Name = "IntuneDeviceID"; Expression = { ($_.id).ToString() } }, @{Name = "EncryptionReadiness"; Expression = { $_.encryptionReadinessState } }, @{Name = "ReportEncryptionState"; Expression = { $_.encryptionState } }, @{Name = "Profiles"; Expression = { if ($_.policyDetails.policyName ) { $_.policyDetails.policyName } else { "No Profile Assigned" } } }, @{Name = "BitlockerState"; Expression = { if ($_.advancedBitLockerStates ) { $_.advancedBitLockerStates } else { "Unknown" } } }, @{Name = "ProfileState"; Expression = { if ($_.encryptionPolicySettingState ) { $_.encryptionPolicySettingState } else { "Unknown" } } } )

    Write-Host "Collected Data for $($RAWIntuneEncryptionReportData.count) objects from MS Graph Encryption Status - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    $RAWIntuneEncryptionReportData | Export-Csv -Path "$($InterimFileLocation)\RAWIntuneEncryptionReportData.csv" -Delimiter ";" -NoTypeInformation
        
    Remove-Variable -Name RAWIntuneEncryptionReportData -Force
    Clear-ResourceEnvironment

    #############################################################################################################################################
    # Bitlocker Key Escrow Data Extraction
    #############################################################################################################################################

    if ($AccessToken) { Remove-Variable -Name AccessToken -Force }
    $AccessToken = Get-MsalToken -TenantId $TenantID -ClientId $ClientID -ForceRefresh -ErrorAction Stop
    if ($AuthenticationHeader) { Remove-Variable -Name AuthenticationHeader -Force }
    $AuthenticationHeader = New-AuthenticationHeader -AccessToken $AccessToken

    Write-Host "Extracting the data from MS Graph Information Protection (Recovery Keys). Expected runtime is 37 minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $RawBitLockerRecoveryKeys = [System.Collections.ArrayList]::new()
    $RawBitLockerRecoveryKeys = [System.Collections.ArrayList]@(Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "informationProtection/bitlocker/recoveryKeys?`$select=id,createdDateTime,deviceId,volumeType" -Headers $AuthenticationHeader -Verbose)#:$VerbosePreference) # 27 Minutes Runtime
    Write-Host "Collected Data for $($RawBitLockerRecoveryKeys.count) objects from MS Graph Encryption Status - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    Write-Host "Processing Recovery Key Data for OS\Data disk Recovery Key Data - Expected Runtime is 15 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
        
    $RawOSBitlockerKey = [System.Collections.ArrayList]@($RawBitLockerRecoveryKeys | Where-Object { $_.volumeType -eq '1' }) # Processing Time = 1 Minute
    $RawDataBitlockerKey = [System.Collections.ArrayList]@($RawBitLockerRecoveryKeys | Where-Object { $_.volumeType -eq '2' }) # Processing Time = 1 Minute

    Remove-Variable -Name RawBitLockerRecoveryKeys -Force
    Clear-ResourceEnvironment

    #############################################################################################################################################
    # Deduplicate recovery key data and selecting only the latest object
    #############################################################################################################################################
    
    $DDRawOSBitlockerKey = [System.Collections.ArrayList]@($RawOSBitlockerKey | Sort-Object deviceId, createdDateTime -Descending | Group-Object -Property deviceId | Select-Object @{Name = 'GroupedList'; Expression = { $_.group | Select-Object -First 1 } } | Select-Object -ExpandProperty GroupedList) # Processing Time = 5 Minutes

    Remove-Variable -Name RawOSBitlockerKey -Force
    Clear-ResourceEnvironment

    $DDRawDataBitlockerKey = [System.Collections.ArrayList]@($RawDataBitlockerKey | Sort-Object deviceId, createdDateTime -Descending | Group-Object -Property deviceId | Select-Object @{Name = 'GroupedList'; Expression = { $_.group | Select-Object -First 1 } } | Select-Object -ExpandProperty GroupedList) # Processing Time = 1 Minute

    Remove-Variable -Name RawDataBitlockerKey -Force
    Clear-ResourceEnvironment

    $OSBitlockerKey = [System.Collections.ArrayList]@($DDRawOSBitlockerKey | Select-Object @{Name = "azureADDeviceId"; Expression = { $_.deviceId.toString() } }, @{Name = "OSBitlockerKeyKnown"; Expression = { if ($_.Id -gt 0) { "TRUE" } } }, @{Name = "OSKeyUploadDate"; Expression = { (Get-Date -Date $_.createdDateTime -Format "yyyy/MM/dd HH:mm") } } | Sort-Object AzureADDeviceID ) # 30 Seconds
        
    Remove-Variable -Name DDRawOSBitlockerKey -Force
    Clear-ResourceEnvironment

    $DataBitlockerKey = [System.Collections.ArrayList]@($DDRawDataBitlockerKey | Select-Object @{Name = "azureADDeviceId"; Expression = { $_.deviceId.toString() } }, @{Name = "DataBitlockerKeyKnown"; Expression = { if ($_.Id -gt 0) { "TRUE" } } }, @{Name = "DataKeyUploadDate"; Expression = { (Get-Date -Date $_.createdDateTime -Format "yyyy/MM/dd HH:mm") } } | Sort-Object AzureADDeviceID ) # 1 Second
        
    Remove-Variable -Name DDRawDataBitlockerKey -Force
    Clear-ResourceEnvironment
        
    Write-Host "Completed Processing Recovery Key Data for OS\Data disk Recovery Key Data - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
        
    $OSBitlockerKey | Export-Csv -Path "$($InterimFileLocation)\OSBitlockerKeys.csv" -Delimiter ";" -NoTypeInformation
    $DataBitlockerKey | Export-Csv -Path "$($InterimFileLocation)\DataBitlockerKeys.csv" -Delimiter ";" -NoTypeInformation
        
    Remove-Variable -Name OSBitlockerKey -Force
    Remove-Variable -Name DataBitlockerKey -Force
    Clear-ResourceEnvironment

    #############################################################################################################################################
    # OnPrem AD Data Extraction
    #############################################################################################################################################

    $AllOPCompsArray = [System.Collections.ArrayList]::new()
    $RAWAllComps = [System.Collections.ArrayList]::new()
    $OPADProcessed = 0
    $OPADCount = $DomainTargets.Count

    Write-Host "Extracting AD OnPrem computer account Data - Expected Runtime is 3 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    foreach ( $DomainTarget in $DomainTargets ) {
    
        [string]$ServerTarget = (Get-ADDomainController -Discover -DomainName $DomainTarget).HostName
       
        $OPDisplay = ( $OPADProcessed / $OPADCount ).tostring("P")
        Write-Progress -Activity "Extracting Data" -Status "Collecting Data from OnPrem AD - $($OPADProcessed) of $($OPADCount) - $($OPDisplay) Completed" -CurrentOperation "Extracting from $($DomainTarget) on $($ServerTarget)" -PercentComplete (( $OPADProcessed / $OPADCount ) * 100 )
        $Comps = [System.Collections.ArrayList]@(Get-ADComputer -Server $ServerTarget -Filter 'operatingsystem -like "Windows 10*"' -Properties CN, CanonicalName, objectGUID, LastLogonDate, Enabled -ErrorAction Stop)
        
        $RAWAllComps += $Comps
        Remove-Variable -Name Comps -Force
    }
        
    Write-Host "Completed AD OnPrem computer account extraction - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    Write-Host "Standardising OnPrem AD Data - Expected Runtime is 2 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    $AllOPCompsArray = [System.Collections.ArrayList]@($RAWAllComps | Select-Object @{Name = "azureADDeviceId"; Expression = { $_.objectGUID.toString() } }, @{Name = "OPDeviceName"; Expression = { $_.CN } }, @{Name = "OPDeviceFQDN"; Expression = { "$($_.CN).$($_.CanonicalName.Split('/')[0])" } }, @{Name = "SourceDomain"; Expression = { "$($_.CanonicalName.Split('/')[0])" } }, @{Name = "OPLastLogonTS"; Expression = { (Get-Date -Date $_.LastLogonDate -Format "yyyy/MM/dd HH:mm") } }, @{Name = "OPSTALE"; Expression = { if ($_.LastLogonDate -le $StaleDate) { "TRUE" } elseif ($_.LastLogonDate -gt $StaleDate) { "FALSE" } else { "NoLoginDateFound" } } } | Sort-Object azureADDeviceId )
    Remove-Variable -Name RAWAllComps -Force
    Clear-ResourceEnvironment

    Write-Host "Completed AD OnPrem Data Standardisation - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    $AllOPCompsArray | Export-Csv -Path "$($InterimFileLocation)\AllOPCompsArray.csv" -Delimiter ";" -NoTypeInformation
                
    #############################################################################################################################################
    # Blending OnPrem AD data with MSGraph Intune Data
    #############################################################################################################################################

    Write-Host "Blending OnPrem AD Data Array with MS Graph Intune Data - Expected Runtime is 11 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    if ($AllOPCompsArray.count -lt 1) {
        $AllOPCompsArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\AllOPCompsArray.csv" -Delimiter ";")
    }
    $IntuneInterimArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\IntuneInterimArray.csv" -Delimiter ";")
    $IntuneInterimArray = [System.Collections.ArrayList]@($IntuneInterimArray | Sort-Object azureADDeviceId)
    $RAWAllDevPreProcArray = [System.Collections.ArrayList]@($IntuneInterimArray | LeftJoin-Object $AllOPCompsArray -On azureADDeviceId)
    $RAWAllPreDevNoIntuneDeviceID = [System.Collections.ArrayList]@($AllOPCompsArray | LeftJoin-Object $IntuneInterimArray -On azureADDeviceId)
    $RAWAllDevNoIntuneDeviceID = [System.Collections.ArrayList]@($RAWAllPreDevNoIntuneDeviceID | Where-Object { $_.IntuneDeviceID -like $null })
    Remove-Variable -Name IntuneInterimArray -Force
    Remove-Variable -Name RAWAllPreDevNoIntuneDeviceID -Force
    Remove-Variable -Name AllOPCompsArray -Force
    Clear-ResourceEnvironment
        
    Write-Host "Completed blending OnPrem AD Data Array with MS Graph Intune Data - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
        
    #############################################################################################################################################
    # Deduplicating the Blended Data
    #############################################################################################################################################

    Write-Host "Deduplicating blended data (OnPrem AD and MS Graph Intune Data) - Expected Runtime is 33 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
        
    $RAWAllDevProcArray = [System.Collections.ArrayList]@($RAWAllDevPreProcArray + $RAWAllDevNoIntuneDeviceID | Sort-Object AzureADDeviceID)
    $DDAllDevProcArray = [System.Collections.ArrayList]@($RAWAllDevProcArray | Group-Object -Property AzureADDeviceID | Select-Object @{Name = 'GroupedList'; Expression = { $_.group | Select-Object -First 1 } } | Select-Object -ExpandProperty GroupedList)
    $DDAllDevProcArray = [System.Collections.ArrayList]@($DDAllDevProcArray | Sort-Object IntuneDeviceID)
    Remove-Variable -Name RAWAllDevPreProcArray -Force
    Remove-Variable -Name RAWAllDevNoIntuneDeviceID -Force
    Clear-ResourceEnvironment

    $DDAllDevProcArray | Export-Csv -Path "$($InterimFileLocation)\DDAllDevProcArray.csv" -Delimiter ";" -NoTypeInformation

    Write-Host "Completed deduplicating blended data (OnPrem AD and MS Graph Intune Data) - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    #############################################################################################################################################
    # Blending KeyRotation Data
    #############################################################################################################################################

    Write-Host "Blending KeyRotation data with previously blended data - Expected Runtime is 7 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    if ($DDAllDevProcArray.count -lt 1) {
        $DDAllDevProcArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\DDAllDevProcArray.csv" -Delimiter ";")
    }

    # Make sure that the script named "Get-DeviceActionData-KeyRotation.ps1" has been run at least once, else this will fail.

    $RotationDataImportFileName = @(Get-ChildItem -Path "$($InterimFileLocation)\" | Where-Object { $_.Name -like "DevActionArray*" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Select-Object name)
    $RotationDataImportFile = "$($InterimFileLocation)\$($RotationDataImportFileName.Name)"
    $DevActionArray = [System.Collections.ArrayList]@(Import-Csv -Path $RotationDataImportFile -Delimiter ";")
        
    $AllDevPreProcArray = [System.Collections.ArrayList]@($DDAllDevProcArray | LeftJoin-Object $DevActionArray -On IntuneDeviceID)
    Remove-Variable -Name RAWAllDevProcArray -Force
    Remove-Variable -Name DevActionArray -Force
    Clear-ResourceEnvironment
    
    Write-Host "Completed blending KeyRotation data with previously blended data - Expected Runtime is 7 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
        
    $AllDevPreProcArray | Export-Csv -Path "$($InterimFileLocation)\AllDevPreProcArray.csv" -Delimiter ";" -NoTypeInformation
    
    #############################################################################################################################################
    # Blending Intune Encryption Report Data
    #############################################################################################################################################

    # Import for Testing
    # $AllDevPreProcArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\AllDevPreProcArray.csv" -Delimiter ";")

    Write-Host "Blending OnPrem AD/Intune Data Array with MS Graph Intune Encryption Data - Expected Runtime is 42 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    if ($AllDevPreProcArray.count -lt 1) {
        $AllDevPreProcArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\AllDevPreProcArray.csv" -Delimiter ";")
    }
    $IntuneEncryptionReportData = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\RAWIntuneEncryptionReportData.csv" -Delimiter ";" | Sort-Object IntuneDeviceID)
        
    $RAWAllDevProcArray = [System.Collections.ArrayList]@($AllDevPreProcArray | LeftJoin-Object $IntuneEncryptionReportData -On IntuneDeviceID | Sort-Object AzureADDeviceID)
    Remove-Variable -Name AllDevPreProcArray -Force
    Remove-Variable -Name IntuneEncryptionReportData -Force
    Clear-ResourceEnvironment

    $AllDevProcArray = [System.Collections.ArrayList]@($RAWAllDevProcArray | Group-Object -Property AzureADDeviceID | Select-Object @{Name = 'GroupedList'; Expression = { $_.group | Select-Object -First 1 } } | Select-Object -ExpandProperty GroupedList)
        
    $AllDevProcArray | Export-Csv -Path "$($InterimFileLocation)\AllDevProcArray.csv" -Delimiter ";" -NoTypeInformation
        
    # Import for Testing
    # $AllDevProcArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\AllDevProcArray.csv" -Delimiter ";")

    Remove-Variable -Name RAWAllDevProcArray -Force
    Clear-ResourceEnvironment

    Write-Host "Completed blending OnPrem AD/Intune Data Array with MS Graph Intune Encryption Data - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    
    #############################################################################################################################################
    # Blending Recovery Key Data
    #############################################################################################################################################

    Write-Host "Blending Recovery Key Data into Report - Expected Runtime is 18 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    if ($AllDevProcArray.count -lt 1) {
        $AllDevProcArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\AllDevProcArray.csv" -Delimiter ";")
    }
    $OSBitlockerKey = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\OSBitlockerKeys.csv" -Delimiter ";")
    $AllDevOSKeyData = [System.Collections.ArrayList]@($AllDevProcArray | LeftJoin-Object $OSBitlockerKey -On azureADDeviceId)

    Remove-Variable -Name AllDevProcArray -Force
    Remove-Variable -Name OSBitlockerKey -Force
    Clear-ResourceEnvironment

    $DataBitlockerKey = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\DataBitlockerKeys.csv" -Delimiter ";")
    $AllDevArray = [System.Collections.ArrayList]@($AllDevOSKeyData | LeftJoin-Object $DataBitlockerKey -On azureADDeviceId)
        
    $AllDevArray | Export-Csv -Path "$($InterimFileLocation)\AllDevArray.csv" -Delimiter ";" -NoTypeInformation
        
    Write-Host "Completed blending Recovery Key Data into Report - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
        
    #############################################################################################################################################
    # Blending AzureAD Data
    #############################################################################################################################################

    Write-Host "Blending AzureAD Data into Report - Expected Runtime is 13 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    if ($AllDevArray.count -lt 1) {
        $AllDevArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\AllDevArray.csv" -Delimiter ";")
    }
    $AzureADDevices = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\AzureADExtract.csv" -Delimiter ";")
    $AzureADBlendedArray = [System.Collections.ArrayList]@($AzureADDevices | LeftJoin-Object $AllDevArray -On azureADDeviceId)
        
    $AzureADBlendedArray | Export-Csv -Path "$($InterimFileLocation)\AzureADBlendedArray.csv" -Delimiter ";" -NoTypeInformation

    Remove-Variable -Name DataBitlockerKey -Force
    Remove-Variable -Name AllDevOSKeyData -Force
    Clear-ResourceEnvironment

    Write-Host "Completed blending AzureAD Data into Report - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    #############################################################################################################################################
    # Processing Report Data
    #############################################################################################################################################

    Write-Host "Processing Report Data - Expected Runtime is 2 Minutes - Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    if ($AzureADBlendedArray.count -lt 1) {
        $AzureADBlendedArray = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\AzureADBlendedArray.csv" -Delimiter ";")
    }
    $BitlockerKeyEscrowReport = [System.Collections.ArrayList]::new()
    $BitlockerKeyEscrowReport = [System.Collections.ArrayList]@($AzureADBlendedArray | Select-Object azureADDeviceId, IntuneDeviceID, ObjectID, MSGraphlastSyncDateTime, AADApproximateLastLogonTimeStamp, AADLastDirSyncTime, AccountEnabled, AADSTALE, OPLastLogonTS, OPSTALE, operatingSystem, osVersion, @{Name = "OSBuild"; Expression = { if (-not($_.osVersion)) { "None" } elseif ($_.osVersion -like "*10240*") { "1507" } elseif ($_.osVersion -like "*10586*") { "1511" } elseif ($_.osVersion -like "*14393*") { "1607" } elseif ($_.osVersion -like "*15063*") { "1703" } elseif ($_.osVersion -like "*15254*" -or $_.osVersion -like "*16299*") { "1709" } elseif ($_.osVersion -like "*17133*" -or $_.osVersion -like "*17134*" -or $_.osVersion -like "*17692*") { "1803" } elseif ($_.osVersion -like "*17763*") { "1809" } elseif ($_.osVersion -like "*18362*" -or $_.osVersion -like "*18990*") { "1903" } elseif ($_.osVersion -like "*18363*") { "1909" } elseif ($_.osVersion -like "*19041*") { "2004" } elseif ($_.osVersion -like "*19042*") { "20H2" } elseif ($_.osVersion -like "*19043*") { "21H1" } elseif ($_.osVersion -like "*19044*") { "21H2" } elseif ($_.osVersion -like "*20161*" -or $_.osVersion -like "*20197*" -or $_.osVersion -like "*20262*") { "2004 (Insider)" } elseif ($_.osVersion -like "*21327*" -or $_.osVersion -like "*21327*") { "21H2 (Dev)" } elseif ($_.osVersion -like "*21354*" -or $_.osVersion -like "*21376*" -or $_.osVersion -like "*21996*") { "21H2 (Insider)" } elseif ($_.osVersion -like "*22000*" -or $_.osVersion -like "*22454*" -or $_.osVersion -like "*22471*" -or $_.osVersion -like "*22483*") { "Windows 11" } else { "Unknown" } } }, DeviceManufacturer, DeviceModel, DeviceSN, managedDeviceName, AADDisplayName, MSGraphDeviceName, OPDeviceName, OPDeviceFQDN, SourceDomain, UserUPN, managementAgent, complianceState, deviceRegistrationState, deviceEnrollmentType, azureADRegistered, enrolledDateTime, ProfileType, MSGraphEncryptionState, @{Name = "OSBitlockerKeyKnown"; Expression = { if ($_.OSBitlockerKeyKnown -like "TRUE") { "TRUE" } else { "FALSE" } } }, OSKeyUploadDate, @{Name = "DataBitlockerKeyKnown"; Expression = { if ($_.DataBitlockerKeyKnown -like "TRUE") { "TRUE" } else { "FALSE" } } }, DataKeyUploadDate, @{Name = "SwitchCompleted"; Expression = { if (($_.MSGraphEncryptionState -match "TRUE") -and ($_.OSBitlockerKeyKnown -match "TRUE")) { "TRUE" } else { "FALSE" } } }, EncryptionReadiness, ReportEncryptionState, Profiles, BitlockerState, ProfileState, aadRegistered, autopilotEnrolled, joinType, KeyRotationResult, KeyRotationRequestDate, KeyRotationDate, KeyRotationError, KeyRotationErrorDescription)
    Remove-Variable -Name AzureADBlendedArray -Force
    Clear-ResourceEnvironment

    $BitlockerKeyEscrowReport | Export-Csv -Path "$($InterimFileLocation)\BitlockerKeyEscrowReport.csv" -Delimiter ";" -NoTypeInformation
    Write-Host "Completed processing Report Data - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
    Remove-Variable -Name BitlockerKeyEscrowReport -Force
    Clear-ResourceEnvironment

    #############################################################################################################################################
    # Processing SCCM Hardware Data and SCCM DDPE Data
    #############################################################################################################################################

    $SCCMHWFiles = @( Get-ChildItem -Path "$($SCCMInputFiles)\" | Where-Object { $_.Name -like "*Hardware Inventory*" })
    $DDPEFiles = @( Get-ChildItem -Path "$($SCCMInputFiles)\" | Where-Object { $_.Name -like "*Dell Encryption*" })

    $SCCMProcData = [System.Collections.ArrayList]::New()
    $DDPEProcData = [System.Collections.ArrayList]::New()

    Write-Host "Processing SCCM HW Data into Array - Expected Runtime is 2 Minutes - Start Time: $(Get-Date -Format "yyyy/MM/dd HH:mm")" -ForegroundColor Green
    $SCCMHWData = [System.Collections.ArrayList]@(foreach ($file in $SCCMHWFiles) { Import-Csv -Path "$($SCCMInputFiles)\$($file.Name)" })
    $SCCMProcDataRAW = [System.Collections.ArrayList]@($SCCMHWData | Select-Object @{Name = "DeviceSN"; Expression = { ($_.BIOS_Serial_Number).ToString() } }, @{Name = "TPMConfig"; Expression = { ($_.TPM_Version2.split(",")[0]) } }, @{Name = "TPMActive"; Expression = { $_.TPM_Active2 } }, @{Name = "TPMEnabled"; Expression = { $_.TPM_Enabled2 } }, @{Name = "TPMOwned"; Expression = { $_.TPM_Owned2 } }, @{Name = "TPMReady"; Expression = { $_.TPM_Ready2 } }, @{Name = "TPMError"; Expression = { $_.TPM_Error } }, @{Name = "SecureBootEnabled"; Expression = { $_.SecureBoot_State2 } }, @{Name = "UEFIEnabled"; Expression = { $_.UEFI_State2 } }, @{Name = "EncryptionSupport"; Expression = { if (($_.TPM_Active2 -match "0") -or ($_.TPM_Enabled2 -match "0") -or ($_.TPM_Error -notlike "0")) { "NotSupported" } elseif (($_.TPM_Version2 -like "2.0*" -and $_.UEFI_State2 -match "1" -and $_.TPM_Active2 -match "1" -and $_.TPM_Enabled2 -match "1" -and $_.TPM_Error -match "0") -and ($_.SecureBoot_State2 -match "1")) { "TPM2.0 Encryption - SecureBoot" } elseif (($_.TPM_Version2 -like "2.0*" -and $_.UEFI_State2 -match "1" -and $_.TPM_Active2 -match "1" -and $_.TPM_Enabled2 -match "1" -and $_.TPM_Error -match "0") -and ($_.SecureBoot_State2 -match "0")) { "TPM2.0 Encryption - NoSecureBoot" } elseif ($_.TPM_Version2 -like "2.0*" -and $_.UEFI_State2 -match "0" -and $_.TPM_Active2 -match "1" -and $_.TPM_Enabled2 -match "1" -and $_.TPM_Error -match "0") { "NotSupported" } elseif ($_.TPM_Version2 -like "1.2*" -and $_.UEFI_State2 -match "1" -and $_.TPM_Active2 -match "1" -and $_.TPM_Enabled2 -match "1" -and $_.TPM_Error -match "0") { "NotSupported" } elseif (($_.TPM_Version2 -like "1.2*" -and $_.UEFI_State2 -match "0" -and $_.TPM_Active2 -match "1" -and $_.TPM_Enabled2 -match "1" -and $_.TPM_Error -match "0")) { "TPM1.2 Encryption" } else { "Unknown" } } }, @{Name = "SCCMHWScan"; Expression = { (Get-Date -Date $_.Last_HW_Scan -Format "yyyy/MM/dd") } }) #3 Minutes
    $SCCMProcData = [System.Collections.ArrayList]@($SCCMProcDataRAW | Where-Object { ($_.DeviceSN -notlike "None" -and $_.DeviceSN -notlike "" -and $_.DeviceSN -notlike "0" -and $_.DeviceSN -notlike "*INVALID*" -and $_.DeviceSN -notlike "*------*") } | Sort-Object DeviceSN )
    Write-Host "SCCM HW Data into Array Processing Completed - Completion Time: $(Get-Date -Format "yyyy/MM/dd HH:mm")" -ForegroundColor Green
    Remove-Variable -Name SCCMHWData -Force
    Remove-Variable -Name SCCMProcDataRAW -Force
    Clear-ResourceEnvironment

    Write-Host "Processing SCCM Dell DDPE Data into Array - Expected Runtime is 20 Seconds - Start Time: $(Get-Date -Format "yyyy/MM/dd HH:mm")" -ForegroundColor Green
    $DDPEData = @( foreach ($file in $DDPEFiles) { Import-Excel -Path "$($SCCMInputFiles)\$($file.Name)" -StartRow 5 -StartColumn 2 })
    $DDPEProcData = [System.Collections.ArrayList]@( $DDPEData | Select-Object @{Name = "OPDeviceName"; Expression = { ($_."Computer Name").ToString() } }, @{Name = "ApplicationName"; Expression = { $_."Application Name" } }, @{Name = "ApplicationVersion"; Expression = { $_."App Version" } } | Sort-Object OPDeviceName )
    Write-Host "SCCM Dell DDPE Data into Array Processing Completed - Completion Time: $(Get-Date -Format "yyyy/MM/dd HH:mm")" -ForegroundColor Green
    Remove-Variable -Name DDPEData -Force
    Clear-ResourceEnvironment
        
    #############################################################################################################################################
    # Blending SCCM Hardware Data
    #############################################################################################################################################
    
    Write-Host "Blending Bitlocker\Intune Data with SCCM HW Data - Expected Runtime is 16 Minutes - Start Time: $(Get-Date -Format "yyyy/MM/dd HH:mm")" -ForegroundColor Green
    if ($BitlockerKeyEscrowReport.count -lt 1) {
        $BitlockerKeyEscrowReport = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\BitlockerKeyEscrowReport.csv" -Delimiter ";" | Sort-Object DeviceSN)
    }
    else {
        $BitlockerKeyEscrowReport = [System.Collections.ArrayList]@($BitlockerKeyEscrowReport | Sort-Object DeviceSN )
    }
    $TempArray2 = [System.Collections.ArrayList]@($BitlockerKeyEscrowReport | LeftJoin-Object $SCCMProcData -On DeviceSN )
    Write-Host "Completed blending Bitlocker\Intune Data with SCCM HW Data - Completion Time: $(Get-Date -Format "yyyy/MM/dd HH:mm")" -ForegroundColor Green

    $TempArray2 | Export-Csv -Path "$($InterimFileLocation)\TempArray2.csv" -Delimiter ";" -NoTypeInformation

    #############################################################################################################################################
    # Blending SCCM DDPE Data
    #############################################################################################################################################

    Write-Host "Blending Bitlocker\Intune Data with SCCM Dell DDPE Data - Expected Runtime is 16 Minutes - Start Time: $(Get-Date -Format "yyyy/MM/dd HH:mm")" -ForegroundColor Green
    if ($TempArray2.count -lt 1) {
        $TempArray2 = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\TempArray2.csv" -Delimiter ";" | Sort-Object DeviceSN)
    }
    else {
        $TempArray2 = [System.Collections.ArrayList]@($TempArray2 | Sort-Object OPDeviceName )
    }
    $TempArray3 = [System.Collections.ArrayList]@($TempArray2 | LeftJoin-Object $DDPEProcData -On OPDeviceName )
    Write-Host "Completed blending Bitlocker\Intune Data with SCCM Dell DDPE Data - Completion Time: $(Get-Date -Format "yyyy/MM/dd HH:mm")" -ForegroundColor Green

    $TempArray3 | Export-Csv -Path "$($InterimFileLocation)\TempArray3.csv" -Delimiter ";" -NoTypeInformation
    
    #############################################################################################################################################
    # Deduplicating and Sorting Data for Export
    #############################################################################################################################################

    Write-Host "Deduplicating Data and sorting for Export - Expected Runtime is 1 Hour 14 Minutes - Start Time: $(Get-Date -Format "yyyy/MM/dd HH:mm")" -ForegroundColor Green
    if ($TempArray3.count -lt 1) {
        $TempArray3 = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\TempArray3.csv" -Delimiter ";" | Sort-Object DeviceSN)
    }
    $DDTempArray3 = [System.Collections.ArrayList]@($TempArray3 | Sort-Object AzureADDeviceID | Group-Object -Property AzureADDeviceID | Select-Object @{Name = 'GroupedList'; Expression = { $_.group | Select-Object -First 1 } } | Select-Object -ExpandProperty GroupedList)

    $DDTempArray3 | Export-Csv -Path "$($InterimFileLocation)\DDTempArray3.csv" -Delimiter ";" -NoTypeInformation
    
    if ($DDTempArray3.count -lt 1) {
        $DDTempArray3 = [System.Collections.ArrayList]@(Import-Csv -Path "$($InterimFileLocation)\DDTempArray3.csv" -Delimiter ";" | Sort-Object DeviceSN)
    }
    $ReportingArray = [System.Collections.ArrayList]@($DDTempArray3 | Where-Object { ($_.AzureADDeviceID -like "*") } | Select-Object AzureADDeviceID, IntuneDeviceID, ObjectID, AADDisplayName, OPDeviceName, OPDeviceFQDN, SourceDomain, managedDeviceName, @{Name = "AADAccountEnabled"; Expression = { $_.AccountEnabled } }, MSGraphlastSyncDateTime, AADApproximateLastLogonTimeStamp, AADLastDirSyncTime, AADSTALE, OPLastLogonTS, OPSTALE, @{Name = "DeviceStale"; Expression = { if (($_.AADSTALE -match "TRUE" -and $_.OPSTALE -match "TRUE") -or ($_.AADSTALE -like $null -and $_.OPSTALE -match "TRUE") -or ($_.AADSTALE -match "TRUE" -and $_.OPSTALE -like $null)) { "TRUE" } else { "FALSE" } } }, operatingSystem, osVersion, OSBuild, DeviceManufacturer, DeviceModel, DeviceSN, UserUPN, managementAgent, EncryptionReadiness, ReportEncryptionState, MSGraphEncryptionState, EncryptionSupport, Profiles, ProfileState, BitLockerState, complianceState, deviceRegistrationState, RecoveryKeyExists, OSBitlockerKeyKnown, OSKeyUploadDate, DataBitlockerKeyKnown, DataKeyUploadDate, SwitchCompleted, TPMConfig, TPMActive, TPMEnabled, TPMOwned, TPMReady, TPMError, SecureBootEnabled, UEFIEnabled, KeyRotationResult, KeyRotationRequestDate, KeyRotationDate, KeyRotationError, KeyRotationErrorDescription, SCCMHWScan, ApplicationName, ApplicationVersion | Sort-Object AzureADDeviceID )
    Write-Host "Data Deduplication and sorting for Export Completed - Completion Time: $(Get-Date -Format "yyyy/MM/dd HH:mm")" -ForegroundColor Green

    #############################################################################################################################################
    # Exporting Report
    #############################################################################################################################################

    $ReportingArray | Export-Excel -Path $ConsolidatedReportExport -ClearSheet -WorksheetName $FileDate -TableName "BitlockerBlendedReport" -AutoSize -AutoFilter -Verbose:$VerbosePreference

    Write-Host "Completed Processing Report Data and Exporting report Files - Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green

    Write-Host "ALL DONE!!! - Script Start Time: $($ScriptStartTime) and Completion Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Red -BackgroundColor Green

    #############################################################################################################################################
    # ALL DONE!!!
    #############################################################################################################################################
    
}