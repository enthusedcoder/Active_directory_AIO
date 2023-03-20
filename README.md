# Active_directory_AIO
All in one Active Directory tool to simplify administration


This is a tool that I created in order to simplify the Active Directory administration of one of our clients.  Apart from the usual 
password reset and unlock user capabilities, this tool also incorporates the creation and disabling of both Active Directory accounts
and the associated office 365 mailbox for new hires as well as terminated users.  It automates the whole process, which would normally
require at least 10 minutes to do manually, with just the click of a few buttons.  Created to provide user-friendly interface so that
non-technical individuals could help in the resolution of the most reocurring tickets.  Used in a live environment and works very well.

## Important considerations

Please note that when you launch the tool for the first time, it may take a while for the main window to appear, as the tool has to check for prerequisites as well as prepare the local machine so that it can be run properly.  In order to verify that the tool is still working, right click an empty space in your task bar, select "Task Manager", select "Processes" (Windows 7) or "Details" (Windows 8 - 10) tab, sort the list of processes by CPU consumption by clicking the header indicated by the number 1 in the image below, and you should see "powershell.exe" consuming a good percentage of your CPU resources.  This means it is working.


![Picture](https://i.imgur.com/WJaHUF3.png)

## Prerequisites
This application requires Powershell 5 to be installed.  You can verify the version of powershell you have installed by opening an elevated Command prompt and typing the below command:
powershell -executionpolicy bypass -command "$PSVersiontable"
The indicated picture below shows where the relevant information is:
![Picture](https://i.imgur.com/1oXi3by.png)
If the indicated number is less than 5, you need to upgrade.  Please note that for those running Windows Server 2008 and Windows 7 machines, this means that you also need to verify that you have at least Dotnet 4.5.2 installed on your machine, as that is a prerequisite for powershell 5.  If you need to install/update the .NET framework on your computer, [download the latest .Net Framework](https://www.microsoft.com/en-us/download/details.aspx?id=55170) and install it.
**IMPORTANT** If you are using a Windows 7 or Server 2008 machine, there is an additional update that needs to be installed before the .NET framework will install successfully.  Please see [this page](https://support.microsoft.com/en-us/help/4020302) for more information as well as the link to download the needed update.  If, however, you have the prerequisites for powershell 5 already installed on your pc, the tool will automatically install powershell 5 for you so that you don't have to figure it out yourself.

![Picture](https://i.imgur.com/gDlPrc3.png)
