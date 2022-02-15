# How do you Migrate from a 3rd party Bitlocker Management solution to Intune, without decrypting the workstations
When I was asked to take on this challenge, I never realised that what I was asked to do was work with our Mission Critical representative and work for over a year on this behemoth of a project. I also was blindsided with something that caught me way off guard, The love for PowerShell and these crazy, deep focused projects.

This project has produced 3 Design Change Requests for Intune, 2 around reporting, and 1 to allow the setting for compliance checking on workstations that if they do not have a Bitlocker recovery key, they cannot be compliant with the policy, if the encryption policy is set to have the recovery key escrowed to AzureAD.

I am sharing with you the working PowerShell code that allowed us to uncover some flaws within the Intune around reporting. I have more than 40 non-working scripts that I created (and abandoned), I will not share these though.

## What is needed for this script to function?

You will need a Service Principal in AzureAD with sufficient rights. I have a Service Principal that I use for multiple processes. I suggest following the guide from <https://msendpointmgr.com/2021/01/18/get-intune-managed-devices-without-an-escrowed-bitlocker-recovery-key-using-powershell/>. My permissions are set as in the image below. Please do not copy my permissions, this Service Principal is used for numerous tasks. I really should correct this, unfortunately, time has not been on my side, so I just work with what work for now. The reporting script "Get-IntuneManagedBitlockerKeyPresence-RAW.ps1" was built from the script referenced in the guide above.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/ServicePrincipalAPIPermissions.jpg)

I also elevate my AzureAD account to 'Intune Administrator', 'Cloud Device Administrator' and 'Security Reader'. These permissions also feel more than needed. Understand that I work in a very large environment, that is very fast paced, so I elevate these as I need them for other tasks as well.

You will need to make sure that you have the following PowerShell modules installed. There is a lot to consider with these modules as some cannot run with others. This was a bit of a learning curve. 

ActiveDirectory\
AzureAD\
ImportExcel\
JoinModule\
MSAL.PS\
PSReadline (May not be needed, not tested without this)

Ultimately, I built a VM on-prem in one of our data centres to run this script, including others. My machine has 4 procs and 16Gb RAM, the reason for an on-prem VM is because most of our workforce is working from home (me included), and running this script is a little slow through the VPN. Our ExpressRoute also makes this data collection significantly more efficient. In a small environment, you will not need this VM.

# Disclaimer.

Ok, so my code may not be very pretty, or efficient in terms of coding. I have only been scripting with PowerShell since September 2020, have had very little (if any), formal PowerShell training and have no previous scripting experience to speak of, apart from the '1 liners' that AD engineers normally create, so please, go easy. I have found that I LOVE PowerShell and finding strange solutions like this have become a passion for me.

My company has a 'Mission Critical' contract with Microsoft for AD and also an E5 licence. We also make extensive use of Azure Cloud, AzureAD, Intune and pretty much whatever Microsoft has to offer from Azure. 

I logged a request with Microsoft to assist us with this process (a few in fact, but the E5 and Mission Critical contracts unlocked some 'Power-Ups' for us, allowing some pretty cool access to some pretty cool specialists - I have asked for permission to share their details and thank them here).

The code in the scripts may not all be my own, and I will thank those whose code I have used, when I explain it below, but the processes and the logic is my own.

## Christopher, enough ramble, how does this thing work?

Before we start, I have expected 'runtimes' for each section. This is for my environment, and will not be accurate for your environment. Use the Measure-Command cmdlet to measure for your specific environment. I added this in because the script could run for hours and appear to be doing nothing.

Also, you will notice a LOT of shared code between my scripts. The reason for this is because a lot of the code is either the same, or similar. The real magic is in the logic in the code after the data extractions occur. Each script has been written to run pretty much separate from the others, as long as the 'input' files exist. The scripts should run.

### Parameters.

The first section is where we supply the TenantID (of the AzureAD tenant) and the ClientID of the Service Principal you have created. If you populate these (hard code), then the script will not ask for these and will immediately go to the Authentication process.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/01-Parameters.jpg)

### Functions.

The functions needed by the script are included in the script. I have modified the 'Invoke-MSGraphOperation' function significantly. I was running into issues with the token and renewing it. I also noted some of the errors went away with a retry or 2, so I built this into the function. Sorry @JankeSkanke @NickolajA for hacking at your work. Check out their work here: https://github.com/MSEndpointMgr/Intune

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/02-Functions.jpg)

### The Variables.

The variables get set here. I have a need to upload the report for another team to use for another report. Enable these and you will be able to do the same.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/03-Variables.jpg)

The variable section also has a section to use the system proxy. I was having trouble with the proxy, intermittently. Adding these lines solved the problem

### The initial Authentication and Token Acquisition.

Ok, so now the 'fun' starts.

The authentication and token acquisition will allow for auth with MFA. You will notice in the script that I have these commands running a few times in the script. This allows for token renewal without requiring MFA again. I also ran into some strange issues with different MS Graph API resources, where a token used for one resource, could not be used on the next resource, this corrects this issue, no idea why, never dug too deep into it because I needed it to work, not be pretty. :-)

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/04-AuthTokenCreation.jpg)

### AzureAD Device Data Extraction.

This section also requires an authentication process and will allow for MFA. The reason why I added this in here is that the script takes a long time to run in my environment, and so, if I perform this extraction first, without the initial auth\token process, the script will complete this process, then sit waiting for auth and MFA, and in essence, not run. Same if this was moved to after the MS Graph extractions. 

Having the 'authy' bits in this order, the script will ask for auth and MFA for MS Graph, then auth and MFA for AzureAD, one after the other with no delay, allowing the script to run without manual intervention. 

You will notice that I sort the data. This is needed to speed up the 'join' processes later on

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/05-AzureADDeviceExtract.jpg)

I export this data to a file, remove the array, then run the Clear-ResourceEnvironment function to free up memory on my system

### Intune Device Data Extraction.

You will notice here that I refresh the token for MS Graph extraction. This should not ask for auth or MFA again as we are simply renewing a current token. Without this section, data extraction fails. Nothing really fancy here, apart from the data transformation perhaps. I also sort the data for use later.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/06-MSGraphIntuneDeviceDataExtract.jpg)

### Intune Encryption Report data Extract.

I switched from manually downloading the Encryption report from Intune, to extracting the report within the script itself. The reason for this was that I would intermittently have errors when trying to download the report. This report also would take about 20 minutes to generate, the failure would also only surface after the 20 minutes. If I received a failure, this would be the case for the day, and would only work again the following day. Not an option for me, so, this now exists.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/07-MSGraphDeviceEncryptionReportExtract.jpg)

### Extract Bitlocker Recovery keys and process this data - I do not keep these keys at all, as this is a security risk.

The section will extract all the recovery keys from the MS Graph API. This data is not kept other than in memory. 

The extracted data is split into 2 arrays, one containing only the recovery key data for the OS disk, and the second containing the recovery key data for the data disk (many of our workstations have 2 disks).

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/08-MSGraphBitlockerRecoveryKeyExtract-1.jpg)

### Deduplicating the recovery key data, process this data and export the data to a file.

Here the script deduplicates the recovery keys and selects the latest recovery key (We often have duplicates here), based on the date uploaded.

The data is then processed and sorted for use later. The 'clean' data is then exported to a file, the arrays removed and the memory cleaned using the Clear-ResourceEnvironment function

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/09-MSGraphBitlockerRecoveryKeyExtract-2.jpg)

### OnPrem AD data extract.

This section exists to be able to map the device back to the source domain, so we know which support team would be responsible for the hardware

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/10-OnPremADDeviceDataExtract.jpg)

### Blending On-Prem AD Data with Intune Data.

Here we first import the Intune data that was extracted earlier. 

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/11-BlendOnPremADwithIntuneData.jpg)

Now things start to get interesting. The script in this section 'blends' the previously extracted data. The data is matched using the 'objectGUID' from the on-prem data extraction with the 'AzureADDeviceID' from the Intune extract. Interestingly, the on-prem AD 'objectGUID' and the 'AzureADDeviceID' is the same. At least if the devices are Hybrid joined. I am unable to comment on other environments though. Your mileage may vary.

I 'blend' the data in both 'directions'. I noted that I got different numbers (object counts) so, for completeness, this process was born. This also creates a number of duplicate records.

### Data Deduplication and AzureAD Data Blending.

In this section, I deduplicate the data, then 'blend' the AzureAD device data with the previously 'blended' data. I deduplicate the data first before blending the next lot of data. 

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/12-DeduplicatingData.jpg)

### Blending the Key Rotation results report - This is where things get fun.

The script will require the output from another script called "Get-DeviceActionData-KeyRotation-RAW.ps1", but this output will be useless till you have run the script called "RotateBitlockerKeys-Parallel-RAW.ps1" first. The reporting script is looking for the output to include the errors of the key rotation command, for each device. NOTE: Windows 10 1909 is the minimum OS build supporting key rotation.

Both scripts are very similar, in that they both have the same structure, and could very easily be included into the same script with switches. I have simply not had the time to do this work as yet.

# Running "RotateBitlockerKeys-Parallel-RAW.ps1" script or "Get-DeviceActionData-KeyRotation-RAW.ps1" script. -> Come, SIDEBAR, lets see these scripts.

Before you can run the script in it's entirety, you need to first run the 'RotateBitlockerKeys-Parallel-RAW.ps1' script. Detailed [here](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk).

Leave the environment for a few days, then run the 'Get-DeviceActionData-KeyRotation-RAW.ps1' script to collect the report file needed for the key rotation data. Detailed [here](https://github.com/christopherbaxter/Intune-DeviceActionReporting-Bulk).

# Ok, back from the Sidebar... Where were we?
### Oh yes, Blending the KeyRotation results from the above scripts, in order to include the Bitlocker key rotation result messages into the main report.

The already blended device data is then blended with the Bitlocker recovery key rotation data extracted with "Get-DeviceActionData-KeyRotation-RAW.ps1"

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/13-BlendingKeyRotationData.jpg)

### The data array in the pipeline is then 'blended' with the Intune Encryption reporting data.

The script will now 'blend' the data in the pipeline with the Intune Encryption data extracted earlier. The data is matched on IntuneDeviceID, then deduplicated

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/14-BlendingEncryptionReportData.jpg)

### Blending the Bitlocker Recovery Key data with the data in the pipeline.

This data is blended matching on AzureDeviceID (I do not match data on names, but only on IDs that cannot be changed)

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/15-BlendingRecoveryKeyData.jpg)

### Blending the AzureAD data into the data in the pipeline.

I am blending the data here using the AzureADDeviceID. I am also using the AzureAD data as the 'left' array and the pipeline data as the 'right' array. This was done because the AzureAD data is the basis of all the data, and will allow for the most accurate and also the most complete amount of data.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/16-BlendingAzureADData.jpg)

### Processing the report data.

This section of code is actually massive. I am doing a ton of data transformation here, including mapping the OS build code to a more human readable format (like 20H2), amongst other stuff. Read the code in the file, this line is 2643 columns long (MASSIVE).

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/17-ProcessingReportData.jpg)

### Processing the SCCM report files received from our SCCM team.

The SCCM team provide me with hardware reports for all devices in SCCM in .csv format. This data is also split between 2 environments, so I get 2 reports. These reports are emailed to me, so I copy them into the 'InputFiles' folder. The SCCM team are doing a migration into Azure as well. The extracts from this report are not yet available to me. I don’t expect a problem with the import, as long as the extracts from them are in the same format. I am confident I will be able to import any number of reports for processing, as long as there is no duplication of objects based on the serial number field. (I will provide a snippet of what the data looks like in these files).

The SCCM hardware reports are relevant in order to be able to get an accurate view of the TPM and BIOS type configuration. The only supported configurations for TPM backed encryption using Bitlocker are either TPM2.0 with native UEFI foot, or TPM1.2 and Legacy boot (possibly UEFI with CSM). Secure boot is also a consideration but may only affect silent encryption.

The SCCM team also provide a report on the machines still running Dell Data Protection agent (we are migrating away from Dell's DDPE). This data is relevant because if the agent exists on a machine, the machine will not escrow the recovery key to Azure. The SCCM team provide these reports in .xlsx format, with the usual SCCM merged cells and the data being in a location that is not expected for the script. I coded around this (so I didn’t need to edit the .xlsx files anymore) with the Import-Excel command. I specify the start row, and start column. A simple fix really.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/18-ProcessSCCMHardwareReports.jpg)

### Now blend this data with the pipeline data.

In this section, I blend the SCCM hardware data based on the device serial number.

The SCCM Dell DDPE report data is the only data blended on a changeable attribute, the on prem device name. No ideal, but not much I can do there on this one. The sheer speed at which the organisation needs this report does not allow me much time to correct this stuff. I need the script functional. And considering that we are talking about less than 1% of the device estate, this section was most certainly not a priority.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/19-BlendingSCCMReportingData.jpg)

### Another round of data deduplication...

In this section, I deduplicate the reporting data again. There is also a little data transformation occurring here. Nothing serious though, most of the heavy lifting has already been done. This is mostly just to make sure that the data fields are in the order we expect them to be in, as well as correcting some column names to be easier to read (this report is epically massive, so every little bit helps). 

The line managing this is 1364 columns wide. So yes, it is a massive report.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/20-DataDedupForExport.jpg)

### Exporting the report

Here I am simply exporting the .xlsx file. This thing is massive, but mainly because of the size of our estate and the numerous bits of data needed in order to properly understand what is happening in the estate.

![](https://github.com/christopherbaxter/Workstation-Bitlocker-management-using-Intune/blob/main/Images/BitlockerReportScript/21-ReportExport.jpg)

The big report is then uploaded to MS for analysis.

# So what’s next here Chris?

Well, I will still upload a sample of the report, as well as a sample of the SCCM files. 

I will also document and upload the scripts I created in order to get workstation logs from machines that are not accessible from the network. These scripts make use of Azure Blob storage, a script used to remediate a number of the errors we have been seeing, like 'Element not Found' and 'General Failure' in the Bitlocker key rotation logs. This script also replaces itself on the workstations every 10 days, the devices collect the new script from Azure Blob storage.

There is much still to see here.