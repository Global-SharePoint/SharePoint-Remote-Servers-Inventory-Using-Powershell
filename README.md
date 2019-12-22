# SharePoint-Remote-Servers-Inventory-Using-Powershell
This script will give complete details about the SharePoint servers with the various color model and will start the mandatory services if any of the services are stopped automatically!!!

Please follow the below steps to execute the script.

-------------------------Ho to deploy or use this script:-------------------------

1. Download the "SharePoint Remote Servers Inventory" zip file.

2. Unzip the Server Reports zip file.

3. Copy the Server Reports folder to SharePoint Test server.

4. Open the ServerReports.ps1 file in PowerShell editor. 

5. Look for the below lines and change the parameter value with your actual value.
Input Parameters:

 1. $servers = @("Server1","Server2","Server3")  #In line number 92..pass the server name.

 2. Configure the "To","CC", "From" and "SMTP" as below:

$To =  "tesemail1@global-sharepoint.com"

$Cc = "tesemail2@global-sharepoint.com"

$From = "SharePoint2016@global-sharepoint.com"

$SMTP = "sharepoint.test.com" 
6. save the script.

7. Run the script - if SMPT server is configured in your sharepoint farm, the configured user will get email. 

8. If you find everything is working as expected you can configure this in the windows task scheduler.

9. Done!!!
---------------------------Deployment done------------------------------------------------------------------

For now I have coded this to handle three servers, but certainly it can be extended to handle multiple servers dynamically...need to extend the code. This ServerReports.ps1 can be configured as a windows task scheduler job, so that everyday morning we can get automatic email to the configured email id with colorful status in html table format. If server has serious issue and need your attention then the particular row/service will be displayed in red color with the proper description and if just warning color will be yellow, if everything fine then color will be lime green. 

Not only color coding - if any of the mandatory services from any of the configured server stopped automatically due to some issue, the code will restart those services and if some issue related to timer job, the code also will restart the timer services to fix that issue - these actions are well described in the email report. More or less I can say we can get the complete details (what went well or wrong) about all servers in last 24 hours. I am sure like me you will be benefitted with this report.

Some sample output(for complete screeenshot please download the attached files).

https://i0.wp.com/global-sharepoint.com/wp-content/uploads/2019/12/Get-SharePoint-server-inventory-using-PowerShell-HTML-Table.png?resize=1024%2C354&ssl=1

https://i1.wp.com/global-sharepoint.com/wp-content/uploads/2019/12/Get-SharePoint-server-inventory-using-PowerShell-HTML-Table2.png?resize=1024%2C215&ssl=1

Reference URL:
https://global-sharepoint.com/powershell/sharepoint-server-monitoring-automation/
