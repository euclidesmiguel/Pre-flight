# Pre-flight - Checking objects before an Exchange hybrid migration

Customers working with FastTrack Center for mail migration can use a tool called Integrated Preflight. It checks for basic migration pre-requisites, such as:  

- Does the target object exist in Office 365?
- Does it have the required e-mail addresses?
- Are all the e-mail addresses valid?
- Does the admin account have the required permissions?
 

The Preflight tool does not check everything, but it is a good start and it saves a lot of later time. Here are some things the proposed pre-flight will not check:

- Is the user properly licensed?
- Are there corrupt items in the mailbox?
- Does the network have enough bandwidth?
 

The tool help customers identify potential errors early and fix them before attempting the mailbox migration. Imagine you tell you user “next week you will have a 100 GB mailbox” and that next week you find out that a missing e-mail address prevented the migration. That is what we want to prevent.


The Preflight tool, however, is only available to customers using our migration services. But that does not mean you cannot apply the same principles to achieve the same results by yourself: You can use PowerShell to connect to Exchange online and perform the same tests we do.

# How to use these scripts  

Clone or download as a zip this repository.  
later in a Windows Powershell 5.1 run the "Preflight.ps1" file.  
The script requires the "ExchangeOnlineManagement" powershell module. if it is not installed, it will attempt to install it.  
If the installation fails, please open Windows Powershell with "Run as Administrator".  

# related links  
https://docs.microsoft.com/en-us/archive/blogs/fasttracktips/pre-flight-and-migration-tool-update