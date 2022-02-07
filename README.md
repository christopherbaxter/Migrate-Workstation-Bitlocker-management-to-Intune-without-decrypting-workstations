# How do you Migrate from a 3rd party Bitlocker Management solution to Intune, without decrypting the workstations
When I was asked to take on this challenge, I never realised that what I was asked to do was work with our Mission Critical representative and work for over a year on this behemouth of a project. I also was blindsided with something that caught me way off guard, The love for PowerShell and these crazy, deep focused projects.

This project has produced 3 Design Change Requests for Intune, 2 around reporting, and 1 to allow the setting for compliance checking on workstations that if they do not have a bitlocker recovery key, they cannot be compliant with the policy, if the encryption policy is set to have the recovery key escrowed to AzureAD.

I am sharing with you the working PowerShell code that allowed us to uncover some flaws within the Intune around reporting. I have more than 40 non-working scripts that I created (and abandoned), I will not share these though.

## What is needed for this script to function?

You will need a Service Principal in AzureAD with sufficient rights. I have a Service Principal that I use for multiple processes. I suggest following the guide from <https://msendpointmgr.com/2021/01/18/get-intune-managed-devices-without-an-escrowed-bitlocker-recovery-key-using-powershell/>. My permissions are set as in the image below. Please do not copy my permissions, this Service Principal is used for numerous tasks. I really should correct this, unfortunately, time has not been on my side, so I just work with what work for now. The reporting script "Get-IntuneManagedBitlockerKeyPresence-RAW.ps1" was built from the script referenced in the guide above.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/ServicePrincipal%20-%20API%20Permissions.jpg)

I also elevate my AzureAD account to 'Intune Administrator', 'Cloud Device Administrator' and 'Security Reader'. These permissions also feel more than needed. Understand that I work in a very large environment, that is very fast paced, so I elevate these as I need them for other tasks as well.

You will need to make sure that you have the following PowerShell modules installed. There is a lot to consider with these modules as some cannot run with others. This was a bit of a learning curve. 

ActiveDirectory\
AzureAD\
ImportExcel\
JoinModule\
MSAL.PS\
PSReadline (May not be needed, not tested without this)

Ultimately, I built a VM on-prem in one of our data centres to run this script, including others. My machine has 4 procs and 16Gb RAM, the reason for an on-prem VM is because most of our workforce is working from home (me included), and running this script is a little slow through the VPN. Our ExpressRoute also makes this data collection significantly more efficient. In a small environment, you will not need this VM.

# Disclaimer

Ok, so my code may not be very pretty, or efficient in terms of coding. I have only been scripting with PowerShell since September 2020, have had very little (if any), formal PowerShell training and have no previous scripting experience to speak of, apart from the '1 liners' that AD engineers normally create, so please, go easy. I have found that I LOVE PowerShell and finding strange solutions like this have become a passion for me.

My company has a 'Mission Critical' contract with Microsoft for AD and also an E5 licence. We also make extensive use of Azure Cloud, AzureAD, Intune and pretty much whatever Microsoft has to offer from Azure. 

I logged a request with Microsoft to assist us with this process (a few in fact, but the E5 and Mission Critical contracts unlocked some 'Power-Ups' for us, allowing some pretty cool access to some pretty cool specialists - I have asked for permission to share their details and thank them here).

The code in the scripts may not all be my own, and I will thank those whose code I have used, when I explain it below, but the processes and the logic is my own.

## Christopher, enough ramble, How does this thing work?

Before we start, I have expected 'runtimes' for each section. This is for my environment, and will not be accurate for your environment. Use the Measure-Command cmdlet to measure for your specific environment. I added this in because the script could run for hours and appear to be doing nothing.

Also, You will notice a LOT of shared code between my scripts. The reason for this is because a lot of the code is either the same, or similar. The real magic is in the logic in the code after the data extractions occur. Each script has been written to run pretty much separate from the others, as long as the 'input' files exist. The scripts should run.

### Parameters

The first section is where we supply the TenantID (of the AzureAD tenant) and the ClientID of the Service Principal you have created. If you populate these (hard code), then the script will not ask for these and will immediately go to the Authentication process.

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/Parameters.jpg)

### Functions

The functions needed by the script are included in the script. I have modified the 'Invoke-MSGraphOperation' function significantly. I was running into issues with the token and renewing it. I also noted some of the errors went away with a retry or 2, so I built this into the function. Sorry @JankeSkanke @NickolajA for hacking at your work. :-)

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/Functions.jpg)

### Asking for the TenantID and the ClientID if not specified

This section will ask for the TenantID and the ClientID if not specified in 'params'

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/TenantID%20-%20ClientID%20request.jpg)

### The Variables

The variables get set here. I have a need to upload the report for another team to use for another report. Enable these and you will be able to do the same.

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/Variables.jpg)

The variable section also has a section to use the system proxy. I was having trouble with the proxy, intermittently. Adding these lines solved the problem

### The initial Authentication and Token Acquisition

Ok, so now the 'fun' starts.

The authentication and token acquisition will allow for auth with MFA. You will notice in the script that I have these commands running a few times in the script. This allows for token renewal without requiring MFA again. I also ran into some strange issues with different MS Graph API resources, where a token used for one resource, could not be used on the next resource, this corrects this issue, no idea why, never dug too deep into it because I needed it to work, not be pretty. :-)

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/InitAuthToken.jpg)

### AzureAD Device Data Extraction

This section also requires an authentication process and will allow for MFA. The reason why I added this in here is that the script takes a long time to run in my environment, and so, if I perform this extraction first, without the initial auth\token process, the script will complete this process, then sit waiting for auth and MFA, and in essence, not run. Same if this was moved to after the MS Graph extractions. 

Having the 'authy' bits in this order, the script will ask for auth and MFA for MS Graph, then auth and MFA for AzureAD, one after the other with no delay, allowing the script to run without manual intervention. 

You will notice that I sort the data. This is needed to speed up the 'join' processes later on

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/ExtractAzureAD%20Device%20details.jpg)

You will notice a hashed out export line, as well as a resource cleanup (Remove-Variable and Clear-ResourceEnvironment). This is to serve 2 purposes, and is included in most of the sections. 

1. Allow for faster troubleshooting of the code (In my environment, the data extraction can take hours, and with a failure, this will mean that I will be waiting, a lot). Enable this to dump the file to the directory of your choice. It is not wise to leave them here.
2. Depending on the amount of data being extracted, you may run short of RAM. This will free that RAM. Also, PowerShell seems to take a beating in terms of performance if there is a LOT of data, this prevents this performance degradation. Use this at your own discretion.

### Intune Device Data Extraction

You will notice here that I refresh the token for MS Graph extraction. This should not ask for auth or MFA again as we are simply renewing a current token. Without this section, data extraction fails. Nothing really fancy here, apart from the data transformation perhaps. I also sort the data for use later.

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/ExtractIntune%20Device%20details.jpg)

### On-Prem AD Data Extraction

This script has been written to extract all the details for all the Windows 10 devices in an AD forest. If you are wanting to specify only a specific domain, you will need to edit the 'Variables' section.

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/On-Prem%20AD%20extract%20-%201.jpg)

I had a number of issues with the extraction with timeouts for some reason. I assume this is some strange network latency or something similar. One needs to remember to pick ones battles.

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/On-Prem%20AD%20extract%20-%202%20-%20Retries.jpg)

This has proven to be pretty reliable, thankfully.

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/On-Prem%20AD%20extract%20-%203%20-%20Data%20Export.jpg)

Nothing really to see here

### Blending On-Prem AD Data with Intune Data

Here you will see that there is a section that if enabled, will import the exported data from the previous extractions, if the relevent export is enabled above. If you are testing, this section will test for the existance of the data in memory, if not present, will import from the 'interim' file\s

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/On-prem%20AD%20with%20Intune%20Data%20Blending%20Process.jpg)

Now things start to get interesting. The script in this section 'blends' the previously extracted data. The data is matched using the 'objectGUID' from the on-prem data extraction with the 'AzureADDeviceID' from the Intune extract. Interestingly, the on-prem AD 'objectGUID' and the 'AzureADDeviceID' is the same. At least if the devices are Hybrid joined. I am unable to comment on other environments though. Your mileage may vary.

I 'blend' the data in both 'directions'. I noted that I got different numbers so, for completeness, this process was born. This also creates a number of duplicate records.

### Data Deduplication and AzureAD Data Blending

In this section, I deduplicate the data, then 'blend' the AzureAD device data with the previously 'blended' data. I deduplicate the data first before blending the next lot of data. 

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/On-Prem%20-%20Intune%20Data%20blend%20deduplication.jpg)

### Report Export - This is where the Magic happens!!!

This little section, is where the magic happens. This section will do the calculation on the OPStale, AADStale and MSGraphLastSyncStale fields. These are calculated fields higher up in the script. If a device is stale on-prem (likely if working remotely), but not in AzureAD, then the device is **NOT** stale\dormant. If the device is not matched to an AzureAD object, then the device **IS** classified as stale\dormant. In the same way, if the device is classified as stale\dormant in AzureAD, and not in on-prem AD, the device is **NOT** stale\dormant. If the AzureAD device is stale in in AzureAD but the device is not matched to an on-prem object, the device **IS** stale.

The export will export all devices in the report, both stale and active. This is easily switched. The code is in the script. There is also the 'remote' export if you would like to send the extract to another server\share.

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/Report%20Export.jpg)

### Whats Next?

I will be incorporating the payload into another script to be able to actually disable the devices that are stale, both in AzureAD and on-prem AD. This will be in the next few weeks.