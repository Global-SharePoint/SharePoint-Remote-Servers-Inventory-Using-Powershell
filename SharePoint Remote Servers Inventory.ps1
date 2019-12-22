################################################################################
#This script will get all details about the server and will send email to concerned person.
#Servers name will be the input to this script.
#Author: Habibur Rahaman
#Date: 22.12.2019
################################################################################

 #=============Variables Declaration====================#
$To =  "tesemail1@global-sharepoint.com"
$Cc = "tesemail2@global-sharepoint.com"
$From = "SharePoint2016@global-sharepoint.com"
$SMTP = "sharepoint.test.com"
$Subject = "[SharePoint 2016] Server and Application Report(Auto email) - $mailDate"
#===========Variables Declaration ends here=============#



#Variable for Site collection
cls
$PSshell = Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorVariable err -ErrorAction SilentlyContinue  
if($PSshell -eq $null){ Add-PSSnapin "Microsoft.SharePoint.PowerShell" }  
  
$fileName = "ServerReportsLog"
#'yyyyMMddhhmm   yyyyMMdd
$enddate = (Get-Date).tostring("yyyyMMddhhmmss")
#$filename =  $enddate + '_VMReport.doc'  
$logFileName = $fileName +"_"+ $enddate+"_Log.txt"   
$invocation = (Get-Variable MyInvocation).Value  
$directoryPath = Split-Path $invocation.MyCommand.Path 


$directoryPathForLog=$directoryPath+"\"+"LogFiles"
if(!(Test-Path -path $directoryPathForLog))  
        {  
            New-Item -ItemType directory -Path $directoryPathForLog            
        }   
 
$logPath = $directoryPathForLog + "\" + $logFileName   
$isLogFileCreated = $False 

$outFileLocation=$directoryPath+"\"+"ServerAndApplicationReport.html"

#Color variable
 $colorWhileMandatoryServicStopped="#ff0000"   #Red
 $colorWhileOptionalServicStopped="#ffbf00"    #Orange
 $colorWhileMandatoryServicesRunning="#00ffff" #blue

$mandatoryServiceNamesInAs01Server=@("World Wide Web Publishing Service","Claims to Windows Token Service","Net.Pipe Listener Adapter","SharePoint Administration","SharePoint Search Host Controller","SharePoint Timer Service","SharePoint Tracing Service","SharePoint User Code Host"); #AS01
$mandatoryServiceMsg="This service was in stopped mode,script started it automatically."
$optionalServiceNamesInAs01Server=@("AppFabric Caching Service","SharePoint VSS Writer") #AS01
$optionalServiceMsg="As expected, nothing wrong."

$mandatoryServiceNamesInOOSServer=@("World Wide Web Publishing Service","Net.Pipe Listener Adapter");
$optionalServiceNamesInOOSServer=@("Claims to Windows Token Service");


$mandatoryServiceNamesInWs01Server=@("AppFabric Caching Service","Claims to Windows Token Service","Windows Fabric Host Service","Net.Pipe Listener Adapter","Service Bus Gateway","Service Bus Message Broker","SharePoint Administration","SharePoint Timer Service","SharePoint Tracing Service","SharePoint User Code Host","World Wide Web Publishing Service","Workflow Manager Backend"); #WS01
$optionalServiceNamesInWs01Server=@("SharePoint Search Host Controller","SharePoint VSS Writer")

#End color variable

#Log file creation function 
function Write-Log([string]$logMsg)  
{   
    if(!$isLogFileCreated){   
        #Write-Host "Creating Log File..."   
        if(!(Test-Path -path $directoryPath))  
        {  
            #Write-Host "Please Provide Proper Log Path" -ForegroundColor Red   
        }   
        else   
        {   
            $script:isLogFileCreated = $True   
            #Write-Host "Log File ($logFileName) Created..."   
            [string]$logMessage = [System.String]::Format("[$(Get-Date)] - {0}", $logMsg)   
            Add-Content -Path $logPath -Value $logMessage   
        }   
    }   
    else   
    {   
        [string]$logMessage = [System.String]::Format("[$(Get-Date)] - {0}", $logMsg)   
        Add-Content -Path $logPath -Value $logMessage   
    }   
} 

#Log file creation function ends here.

#$head='<style>body{font-family:Calibri;font-size:10pt;}th{background-color:black;color:white;}td{background-color:#19fff0;color:black;}h4{margin-right: 0px; margin-bottom: 0px; margin-left: 0px;}</style>'

#Add SharePoint PowerShell Snap-In
Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue
#SharePoint Servers, use comma to separate.......first server=Application(AS01)...second server=WFE(WS01)..
$servers = @("Server1","Server2","Server3")
#===============#
# Server Report # 
#===============#
#Memory Ustilization
#$Memory = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $servers| Select CsName , FreePhysicalMemory  , TotalVisibleMemorySize , Status | ConvertTo-Html -Fragment



Try
  {

[xml]$Memory = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $servers| Select CsName , @{Name="FreePhysicalMemory"; Expression = {[int]($_.FreePhysicalMemory/1mb)}} , @{Name="TotalVisibleMemorySize"; Expression = {[int]($_.TotalVisibleMemorySize/1mb)}},Status | ConvertTo-Html -Fragment

for($i=1;$i -le $Memory.table.tr.count-1;$i++)
 {
  

     $Memory.table.tr[$i].ChildNodes[($Memory.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $Memory.table.tr[$i].ChildNodes[($Memory.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $Memory.table.tr[$i].ChildNodes[($Memory.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $Memory.table.tr[$i].ChildNodes[($Memory.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     
 }

 $body = @"
<H2> </H2>
$($Memory.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $MemoryModified=$htmlFile;
  
   $logMessage="Memory report test completed"
   #Write-Host $logMessage -BackgroundColor DarkGreen
   Write-Log $logMessage
   }
 catch
    {
    $ErrorMessage = "Error in memory test: "+$_.Exception.Message
    #Write-Host $ErrorMessage -BackgroundColor Red
    Write-Log $ErrorMessage

    }

#Disk Report
$serverFreeDiskInCdriveShouldBe="80"   #Should be more than double of RAM Inslalled.
$serverFreeDiskInCdrive=""
$serverDiskActionMsg="immediately inform to ADVANIA"
try
{
[xml]$diskReportInServer1 = Get-WmiObject -Class Win32_LogicalDisk -Filter DriveType=3 -ComputerName $servers[0] | 
Select DeviceID , @{Name="Size(GB)";Expression={"{0:N1}" -f($_.size/1gb)}}, @{Name="Free space(GB)";Expression={"{0:N1}" -f($_.freespace/1gb)}}| ConvertTo-Html -Fragment
   

 for($i=1;$i -le $diskReportInServer1.table.tr.count-1;$i++)
 {
          
     $diskReportInServer1.table.tr[$i].ChildNodes[($diskReportInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $diskReportInServer1.table.tr[$i].ChildNodes[($diskReportInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $diskReportInServer1.table.tr[$i].ChildNodes[($diskReportInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
    
 }


 $body = @"
<H2> </H2>
$($diskReportInServer1.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $diskReportInServer1Modified=$htmlFile;

[xml]$diskReportInServer2 = Get-WmiObject -Class Win32_LogicalDisk -Filter DriveType=3 -ComputerName $servers[1] | 
Select DeviceID , @{Name="Size(GB)";Expression={"{0:N1}" -f($_.size/1gb)}}, @{Name="Free space(GB)";Expression={"{0:N1}" -f($_.freespace/1gb)}} | ConvertTo-Html -Fragment

for($i=1;$i -le $diskReportInServer2.table.tr.count-1;$i++)
 {
  
     $diskReportInServer2.table.tr[$i].ChildNodes[($diskReportInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $diskReportInServer2.table.tr[$i].ChildNodes[($diskReportInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $diskReportInServer2.table.tr[$i].ChildNodes[($diskReportInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
          
 }

 $body = @"
<H2></H2>
$($diskReportInServer2.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $diskReportInServer2Modified=$htmlFile;

$logMessage="Disk report test completed"
   #Write-Host $logMessage -BackgroundColor DarkGreen
   Write-Log $logMessage
}
catch
{
    $ErrorMessage = "Error in disk report test: "+$_.Exception.Message
    #Write-Host $ErrorMessage -BackgroundColor Red
    Write-Log $ErrorMessage
}

#Server UpTime
try
{
[xml]$FarmUpTime = Get-WmiObject -class Win32_OperatingSystem -ComputerName $servers | 
Select-Object __SERVER,@{label='LastRestart';expression={$_.ConvertToDateTime($_.LastBootUpTime)}} | ConvertTo-Html -Fragment


for($i=1;$i -le $FarmUpTime.table.tr.count-1;$i++)
 {
     $FarmUpTime.table.tr[$i].ChildNodes[($FarmUpTime.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $FarmUpTime.table.tr[$i].ChildNodes[($FarmUpTime.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
             
 }

 $body = @"
<H2></H2>
$($FarmUpTime.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $FarmUpTimeModified=$htmlFile;
	  
  $logMessage="Server Uptime report test completed."
   #Write-Host $logMessage -BackgroundColor DarkGreen
   Write-Log $logMessage
}
catch
{
    $ErrorMessage = "Error in server uptime report test: "+$_.Exception.Message
    #Write-Host $ErrorMessage -BackgroundColor Red
    Write-Log $ErrorMessage

}

#Core SharePoint services

try
{

[xml]$coreservicesInServer1 = Get-WmiObject -Class Win32_Service -ComputerName $servers[0] | ? {($_.Name -eq "AppFabricCachingService") -or ($_.Name -eq "c2wts") -or ($_.Name -eq "FIMService") `
-or ($_.Name -eq "FIMSynchronizationService") -or ($_.Name -eq "Service Bus Gateway") -or ($_.Name -eq "Service Bus Message Broker") -or ($_.Name -eq "SPAdminV4") `
-or ($_.Name -eq "SPSearchHostController") -or ($_.Name -eq "OSearch15") -or ($_.Name -eq "SPTimerV4") -or ($_.Name -eq "SPTraceV4") -or ($_.Name -eq "SPUserCodeV4") `
-or ($_.Name -eq "SPWriterV4") -or ($_.Name -eq "FabricHostSvc") -or ($_.Name -eq "WorkflowServiceBackend") -or  ($_.Name -eq "W3SVC") -or ($_.Name -eq "NetPipeActivator")} `
| Select-Object DisplayName, StartName, StartMode, State,@{Name = 'Expected Status'; Expression = {"Running"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} | ConvertTo-Html -Fragment

 for($i=1;$i -le $coreservicesInServer1.table.tr.count-1;$i++)
 {
   $test=$coreservicesInServer1.table.tr;

   if(($coreservicesInServer1.table.tr[$i].td[-3] -eq "Stopped")) 
   { 
        
        $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].InnerText="Not OK"
        foreach($oneMandatorySrvc in $mandatoryServiceNamesInAs01Server)
        {
            $serviceStatusColumn=$coreservicesInServer1.table.tr[$i].td[-3];
            $serviceNameColumn=$coreservicesInServer1.table.tr[$i].td[-6];

            if($serviceStatusColumn -contains "Stopped" -and $mandatoryServiceNamesInAs01Server -contains $serviceNameColumn)
            {
               $service = get-service -ComputerName $servers[0] -Name "W3SVC" #1. World wide service(IIS)
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)

               }

               $service = get-service -ComputerName $servers[0] -Name "c2wts" #2. Claims to Windows Token Service
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

               $service = get-service -ComputerName $servers[0] -Name "NetPipeActivator" #3. Net.Pipe Listener Adapter
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

               $service = get-service -ComputerName $servers[0] -Name "SPAdminV4" #4. SharePoint Administration
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

               $service = get-service -ComputerName $servers[0] -Name "SPSearchHostController" #5. SharePoint Search Host Controller
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

               $service = get-service -ComputerName $servers[0] -Name "SPTimerV4" #6. SharePoint Timer Service
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

               $service = get-service -ComputerName $servers[0] -Name "SPTraceV4" #7. SharePoint Tracing Service
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

               $service = get-service -ComputerName $servers[0] -Name "SPUserCodeV4" #8. SharePoint User Code Host
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

            }
        }

    }
    elseif(($coreservicesInServer1.table.tr[$i].td[-3] -eq "Running"))
    {
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)

    }
    
    if(($optionalServiceNamesInAs01Server -contains $coreservicesInServer1.table.tr[$i].td[-6])) 
    {
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].InnerText="Stopped"
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].InnerText=$optionalServiceMsg;

     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer1.table.tr[$i].ChildNodes[($coreservicesInServer1.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)

    }
 }

 $body = @"
<H2></H2>
$($coreservicesInServer1.innerxml)
"@

# Convert to HTML and save the file
#ConvertTo-Html -Head $head -Body $body |
      #Out-File $filename
      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $coreservicesInServer1Modifed=$htmlFile;

[xml]$coreservicesInServer2 = Get-WmiObject -Class Win32_Service -ComputerName $servers[1] | ? {($_.Name -eq "AppFabricCachingService") -or ($_.Name -eq "c2wts") -or ($_.Name -eq "FIMService") `
-or ($_.Name -eq "FIMSynchronizationService") -or ($_.Name -eq "Service Bus Gateway") -or ($_.Name -eq "Service Bus Message Broker") -or ($_.Name -eq "SPAdminV4") `
-or ($_.Name -eq "SPSearchHostController") -or ($_.Name -eq "OSearch15") -or ($_.Name -eq "SPTimerV4") -or ($_.Name -eq "SPTraceV4") -or ($_.Name -eq "SPUserCodeV4") `
-or ($_.Name -eq "SPWriterV4") -or ($_.Name -eq "FabricHostSvc") -or ($_.Name -eq "WorkflowServiceBackend") -or  ($_.Name -eq "W3SVC") -or ($_.Name -eq "NetPipeActivator")} `
| Select-Object DisplayName, StartName, StartMode, State,@{Name = 'Expected Status'; Expression = {"Running"}},@{Name = 'Action Taken'; Expression = {"Not Required"}}| ConvertTo-Html -Fragment

for($i=1;$i -le $coreservicesInServer2.table.tr.count-1;$i++)
 {
   $test=$coreservicesInServer2.table.tr;

   if(($coreservicesInServer2.table.tr[$i].td[-3] -eq "Stopped")) 
   { 
        
        $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText="Not OK"
        
        foreach($oneMandatorySrvc in $mandatoryServiceNamesInWs01Server)
        {
            $serviceStatusColumn=$coreservicesInServer2.table.tr[$i].td[-3];
            $serviceNameColumn=$coreservicesInServer2.table.tr[$i].td[-6];

            if($serviceStatusColumn -contains "Stopped" -and $mandatoryServiceNamesInWs01Server -contains $serviceNameColumn)
            {
               $service = get-service -ComputerName $servers[1] -Name "W3SVC" #1. World Wide Web Publishing Service(IIS)
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)

               }

               $service = get-service -ComputerName $servers[1] -Name "c2wts" #2. Claims to Windows Token Service
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

               $service = get-service -ComputerName $servers[1] -Name "NetPipeActivator" #3. Net.Pipe Listener Adapter
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

               $service = get-service -ComputerName $servers[1] -Name "SPAdminV4" #4. SharePoint Administration
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }
               			   
               $service = get-service -ComputerName $servers[1] -Name "WorkflowServiceBackend" #5. Workflow Manager Backend
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }
			   

               $service = get-service -ComputerName $servers[1] -Name "SPTimerV4" #6. SharePoint Timer Service
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

               $service = get-service -ComputerName $servers[1] -Name "SPTraceV4" #7. SharePoint Tracing Service
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

               $service = get-service -ComputerName $servers[1] -Name "SPUserCodeV4" #8. SharePoint User Code Host
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }
			   
			   $service = get-service -ComputerName $servers[1] -Name "AppFabricCachingService" #9. AppFabric Caching Service
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }
			   
			   $service = get-service -ComputerName $servers[1] -Name "FabricHostSvc" #10. Windows Fabric Host Service
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }
			   
			   $service = get-service -ComputerName $servers[1] -Name "Service Bus Gateway" #11. Service Bus Gateway
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }
			   
			   $service = get-service -ComputerName $servers[1] -Name "Service Bus Message Broker" #12. Service Bus Message Broker
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;

               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               }

            }
        }

    }
    elseif(($coreservicesInServer2.table.tr[$i].td[-3] -eq "Running"))
    {
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)

    }
    
    if(($optionalServiceNamesInWs01Server -contains $coreservicesInServer2.table.tr[$i].td[-6])) 
    {
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].InnerText="Stopped"
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].InnerText=$optionalServiceMsg;

     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer2.table.tr[$i].ChildNodes[($coreservicesInServer2.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)

    }
 }

 $body = @"
<H2></H2>
$($coreservicesInServer2.innerxml)
"@

# Convert to HTML and save the file
#ConvertTo-Html -Head $head -Body $body |
      #Out-File $filename

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $coreservicesInServer2Modifed=$htmlFile;
    
   $logMessage="SharePoint core service report test completed."
   #Write-Host $logMessage -BackgroundColor DarkGreen
   Write-Log $logMessage

}
catch
{
    $ErrorMessage = "Error in sharepoint core service report test: "+$_.Exception.Message
    #Write-Host $ErrorMessage -BackgroundColor Red
    Write-Log $ErrorMessage
}
###Office Online Server...
[xml]$coreservicesInServer3 = Get-WmiObject -Class Win32_Service -ComputerName $servers[2] | ? {($_.Name -eq "c2wts") -or ($_.Name -eq "W3SVC") -or ($_.Name -eq "NetPipeActivator")}| Select-Object DisplayName, StartName, StartMode, State,@{Name = 'Expected Status'; Expression = {"Running"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} | ConvertTo-Html -Fragment

for($i=1;$i -le $coreservicesInServer3.table.tr.count-1;$i++)
 {
           $test=$coreservicesInServer3.table.tr;
           $serviceStatusColumn=$coreservicesInServer3.table.tr[$i].td[-3];
           $serviceNameColumn=$coreservicesInServer3.table.tr[$i].td[-6];

            if($serviceStatusColumn -contains "Stopped" -and $mandatoryServiceNamesInOOSServer -contains $serviceNameColumn)
            {
               $service = get-service -ComputerName $servers[2] -Name "W3SVC" #1. World Wide Web Publishing Service(IIS)
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)

               }

               $service = get-service -ComputerName $servers[2] -Name "NetPipeActivator" #2.Net.Pipe Listener Adapter
               if($service.Status -eq "Stopped")
               {
               $service.Start()
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-1)].InnerText=$mandatoryServiceMsg;
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)

               }

          }

          elseif(($coreservicesInServer3.table.tr[$i].td[-3] -eq "Running"))
            {
         $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
         $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
         $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
         $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
         $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
         $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
            }
    if(($optionalServiceNamesInOOSServer -contains $coreservicesInServer3.table.tr[$i].td[-6])) 
    {
     $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-2)].InnerText="Stopped"
     $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-1)].InnerText=$optionalServiceMsg;

     $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)
     $coreservicesInServer3.table.tr[$i].ChildNodes[($coreservicesInServer3.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileOptionalServicStopped)

    }
}

$body = @"
<H2></H2>
$($coreservicesInServer3.innerxml)
"@

# Convert to HTML and save the file
#ConvertTo-Html -Head $head -Body $body |
      #Out-File $filename

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $coreservicesInServer3Modifed=$htmlFile;

###Ends Office online server

#====================#
# Application Report #
#====================#

$applicationStatusMsg="Login to SP server and troubleshoot it."
#SharePoint Farm Status
try
{
    [xml]$SPFarm = Get-SPFarm | select Name , NeedsUpgrade , Status , BuildVersion,@{Name = 'Expected Status'; Expression = {"Online"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} |ConvertTo-Html –Fragment


 for($i=1;$i -le $SPFarm.table.tr.count-1;$i++)
 {
    
    if($SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-4)].InnerText -eq "Online")
    {

     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     }
     else
     {
      
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)

     $SPFarm.table.tr[$i].ChildNodes[($SPFarm.table.tr[$i].ChildNodes.Count-1)].InnerText=$applicationStatusMsg
     
     }         
 }

 $body = @"
<H2> </H2>
$($SPFarm.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $SPFarmModified=$htmlFile;

    $logMessage="SharePoint farm report test completed."
   #Write-Host $logMessage -BackgroundColor DarkGreen
   Write-Log $logMessage
}
catch
{
   $ErrorMessage = "Error in sharepoint farm report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}

#Web Application Pool Status

$applicationStatusMsgWAppPool="Login to server and troubleshoot it."
try
{
   [xml]$WAppPool = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.ApplicationPools | select Name, Username,Status,@{Name = 'Expected Status'; Expression = {"Online"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} | ConvertTo-Html -Fragment

  for($i=1;$i -le $WAppPool.table.tr.count-1;$i++)
 {
    if($WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-3)].InnerText -eq "Online")
    {

     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     }
     else
     {
     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     
     $WAppPool.table.tr[$i].ChildNodes[($WAppPool.table.tr[$i].ChildNodes.Count-1)].InnerText=$applicationStatusMsgWAppPool

     }        
     
 }

 $body = @"
<H2> </H2>
$($WAppPool.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $WAppPoolModified=$htmlFile;


   $logMessage="Web application pool report test completed."
   #Write-Host $logMessage -BackgroundColor DarkGreen
   Write-Log $logMessage
}
catch
{
  
   $ErrorMessage = "Error in Web application pool report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage

}

#Web Application Status Check
$webApplicationStatusCheckMsg="Login to the server and troubleshoot it."
try
{
#$WebApplication = Get-SPWebApplication | Select Name , Url , ContentDatabases , NeedsUpgrade , Status | ConvertTo-Html -Fragment
[xml]$WebApplication = Get-SPWebApplication | Select Name , Url, NeedsUpgrade , Status,@{Name = 'Expected Status'; Expression = {"Online"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} | ConvertTo-Html -Fragment


  for($i=1;$i -le $WebApplication.table.tr.count-1;$i++)
 {
    if($WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-3)].InnerText -eq "Online")
    {

     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
      
    }
    else
    {
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
    
     $WebApplication.table.tr[$i].ChildNodes[($WebApplication.table.tr[$i].ChildNodes.Count-1)].InnerText=$webApplicationStatusCheckMsg;
    }      
     
 }

 $body = @"
<H2> </H2>
$($WebApplication.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $WebApplicationModified=$htmlFile;

$logMessage="Web application status report test completed."
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage

}
catch
{
   $ErrorMessage = "Error in web application status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage

}
#IIS Total number of current connections
try
{
[xml]$IIS1 = Get-Counter -ComputerName $servers -Counter "\web service(_total)\Current Connections" | 
Select Timestamp , Readings | ConvertTo-Html -Fragment

 for($i=1;$i -le $IIS1.table.tr.count-1;$i++)
 {
  

     $IIS1.table.tr[$i].ChildNodes[($IIS1.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $IIS1.table.tr[$i].ChildNodes[($IIS1.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)   
             
     
 }

 $body = @"
<H2> </H2>
$($IIS1.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $IIS1Modified=$htmlFile;

$logMessage="IIS connection report test completed."
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage

}
catch
{
   $ErrorMessage = "Error in total number of iis connection report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage 
}
#Service Application Pool Status
#Returns the specified Internet Information Services (IIS) application pool. NOTE, these are the SERVICE app pools, NOT WEB app pools.
try
{
[xml]$SAppPool = Get-SPServiceApplicationPool | Select Name , ProcessAccountName , status,@{Name = 'Expected Status'; Expression = {"Online"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} | ConvertTo-Html -Fragment

for($i=1;$i -le $SAppPool.table.tr.count-1;$i++)
 {
    
    if($SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-3)].InnerText -eq "Online")

    {

     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
      
    }
    else
    {

     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     
     $SAppPool.table.tr[$i].ChildNodes[($SAppPool.table.tr[$i].ChildNodes.Count-1)].InnerText="Login to SharePoint server to troubleshoot it.";
    }         
     
 }
 $body = @"
<H2> </H2>
$($SAppPool.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $SAppPoolModified=$htmlFile;


$logMessage="Service application pool status test completed."
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage
}
catch
{
  $ErrorMessage = "Error in service application pool status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}

#Service Application Status
try
{
[xml]$ServiceAppplication = Get-SPServiceApplication | Select DisplayName , ApplicationVersion , Status , NeedsUpgrade,@{Name = 'Expected Status'; Expression = {"Online"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} | Convertto-Html -Fragment

for($i=1;$i -le $ServiceAppplication.table.tr.count-1;$i++)
 {
  
     if($ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-4)].InnerText -eq "Online")
     {
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     }
     else
     {
      
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     
     $ServiceAppplication.table.tr[$i].ChildNodes[($ServiceAppplication.table.tr[$i].ChildNodes.Count-1)].InnerText="Login to the server to troubleshoot it."

     }        
     
 }

 $body = @"
<H2> </H2>
$($ServiceAppplication.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $ServiceAppplicationModified=$htmlFile;


$logMessage="Service application status test completed."
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage

}
catch
{
  $ErrorMessage = "Error in service application status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}
#Service Application Proxy Status
try
{
[xml]$ApplicationProxy = Get-SPServiceApplicationProxy | Select TypeName , Status , NeedsUpgrade,@{Name = 'Expected Status'; Expression = {"Online"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} | ConvertTo-Html -Fragment

for($i=1;$i -le $ApplicationProxy.table.tr.count-1;$i++)
 {
  
    if($ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-4)].InnerText -eq "Online")
    {
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     #$ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     }
     else
     {
      
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     
     $ApplicationProxy.table.tr[$i].ChildNodes[($ApplicationProxy.table.tr[$i].ChildNodes.Count-1)].InnerText="Login to the server to troubleshoot it."

     }          
     
 }

 $body = @"
<H2> </H2>
$($ApplicationProxy.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $ApplicationProxyModified=$htmlFile;

$logMessage="Service application proxy status test completed."
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage
}
catch
{
   $ErrorMessage = "Error in service application proxy status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}
#Search Administration Component Status
try
{
[xml]$SPSearchAdminComponent = Get-SPEnterpriseSearchAdministrationComponent -SearchApplication "Search Service Application" | Select Servername , Initialized | ConvertTo-Html -Fragment

for($i=1;$i -le $SPSearchAdminComponent.table.tr.count-1;$i++)
 {
  

     $SPSearchAdminComponent.table.tr[$i].ChildNodes[($SPSearchAdminComponent.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPSearchAdminComponent.table.tr[$i].ChildNodes[($SPSearchAdminComponent.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
                   
     
 }

 $body = @"
<H2> </H2>
$($SPSearchAdminComponent.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $SPSearchAdminComponentModified=$htmlFile;

$logMessage="Search admin component status report completed."
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage

}
catch
{
 $ErrorMessage = "Error in search admin component status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}
#Search Scope Status 

#Search Query Topology Status
#Using Get-SPEnterpriseSearchTopology instead of Get-SPEnterpriseSearchQueryTopology
try
{
[xml]$SPQueryTopology = Get-SPEnterpriseSearchTopology -SearchApplication "Search Service Application" | select TopologyId , State | ConvertTo-Html -Fragment

for($i=1;$i -le $SPQueryTopology.table.tr.count-1;$i++)
 {
  

     $SPQueryTopology.table.tr[$i].ChildNodes[($SPQueryTopology.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPQueryTopology.table.tr[$i].ChildNodes[($SPQueryTopology.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
                 
     
 }

 $body = @"
<H2> </H2>
$($SPQueryTopology.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $SPQueryTopologyModified=$htmlFile;

$logMessage="Search query topology status report completed."
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage
}
catch
{
   $ErrorMessage = "Error in search query topology status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}

#Content Sources Status with Crawl Log Counts
try
{
[xml]$SPContentSource = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication "Search Service Application" |
Select Name , SuccessCount , CrawlStatus, LevelHighErrorCount, ErrorCount, DeleteCount, WarningCount | ConvertTo-Html -Fragment

for($i=1;$i -le $SPContentSource.table.tr.count-1;$i++)
 {
  
     $SPContentSource.table.tr[$i].ChildNodes[($SPContentSource.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPContentSource.table.tr[$i].ChildNodes[($SPContentSource.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPContentSource.table.tr[$i].ChildNodes[($SPContentSource.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPContentSource.table.tr[$i].ChildNodes[($SPContentSource.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPContentSource.table.tr[$i].ChildNodes[($SPContentSource.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPContentSource.table.tr[$i].ChildNodes[($SPContentSource.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPContentSource.table.tr[$i].ChildNodes[($SPContentSource.table.tr[$i].ChildNodes.Count-7)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
                       
     
 }

 $body = @"
<H2> </H2>
$($SPContentSource.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $SPContentSourceModified=$htmlFile;

$logMessage="Search content source status report completed."
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage

}
catch
{
   $ErrorMessage = "Error in search content source status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}
#Crawl Component Status
#$SPCrawlComponent = Get-SPEnterpriseSearchCrawlComponent -SearchApplication "Search Service Application" -CrawlTopology "0e49b522-4b7b-44d5-bc77-99b53e1c9f03"| Select ServerName , State , DesiredState , IndexLocation | ConvertTo-Html -Fragment
try
{
[xml]$SPCrawlComponent = Get-SPEnterpriseSearchStatus -SearchApplication "Search Service Application" | Select Name , State,@{Name = 'Expected State'; Expression = {"Active"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} | ConvertTo-Html -Fragment

for($i=1;$i -le $SPCrawlComponent.table.tr.count-1;$i++)
 {
  
    if($SPCrawlComponent.table.tr[$i].ChildNodes[($SPCrawlComponent.table.tr[$i].ChildNodes.Count-3)].InnerText -eq "Active")
    {

     $SPCrawlComponent.table.tr[$i].ChildNodes[($SPCrawlComponent.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPCrawlComponent.table.tr[$i].ChildNodes[($SPCrawlComponent.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPCrawlComponent.table.tr[$i].ChildNodes[($SPCrawlComponent.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $SPCrawlComponent.table.tr[$i].ChildNodes[($SPCrawlComponent.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
    } 
    else
    {
     $SPCrawlComponent.table.tr[$i].ChildNodes[($SPCrawlComponent.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SPCrawlComponent.table.tr[$i].ChildNodes[($SPCrawlComponent.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SPCrawlComponent.table.tr[$i].ChildNodes[($SPCrawlComponent.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
     $SPCrawlComponent.table.tr[$i].ChildNodes[($SPCrawlComponent.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicStopped)
             
     
     $SPCrawlComponent.table.tr[$i].ChildNodes[($SPCrawlComponent.table.tr[$i].ChildNodes.Count-1)].InnerText="Login to the server to troubleshoot it."

    }  
 }

 $body = @"
<H2> </H2>
$($SPCrawlComponent.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $SPCrawlComponentModified=$htmlFile;


$logMessage="Search crawl component status report completed."
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage

}
catch
{
   $ErrorMessage = "Error in search crawl component status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}

#Content Database Status
try
{
[xml]$DBCheck = Get-SPDatabase | Select Name , Status , NeedsUpgrade , @{Name="size(GB)";Expression={"{0:N1}" -f($_.Disksizerequired/1gb)}},@{Name = 'Expected Status'; Expression = {"Online"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} | ConvertTo-Html -Fragment
for($i=1;$i -le $DBCheck.table.tr.count-1;$i++)
 {
    
    if($DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-5)].InnerText -eq "Online")
    {

     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
    } 
    else
    {
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor',$colorWhileMandatoryServicesRunning)
     
     $DBCheck.table.tr[$i].ChildNodes[($DBCheck.table.tr[$i].ChildNodes.Count-1)].InnerText="Login to the server to troubleshoot it."

    }
  
 }

 $body = @"
<H2> </H2>
$($DBCheck.innerxml)
"@

      $htmlFile=ConvertTo-Html -Head $head -Body $body
      $DBCheckModified=$htmlFile;

$logMessage="Database check status report completed."
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage

}
catch
{
   $ErrorMessage = "Error in database check status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}


#Failed Timer Jobs, Last 7 days

$f = Get-SPFarm 
$ts = $f.TimerService 
$jobs = $ts.JobHistoryEntries | ?{$_.Status -eq "Failed" -and $_.StartTime -gt ((get-date).AddDays(-7))}  
#$failedJobs=$jobs.Count      //.Count seems to do nothing. This one does something on the other hand:
try
{
[xml]$failedjobs = $jobs | select StartTime,JobDefinitionTitle,Status,ErrorMessage,@{Name = 'Expected Status'; Expression = {"Running"}},@{Name = 'Action Taken'; Expression = {"Not Required"}} | ConvertTo-Html -Fragment
$logMessage="Failed timer job status report completed."

$msgForWOPITimer="Low severity,Restart the timer service from both ther servers.";
$msgForUserProfileServiceFeedCache="Low severity,however need to take action if frequency of failure is increased.";
$msgForOtherTimer="Check the log and take appropirate action(Restart the timer service from both ther servers)."

for($i=1;$i -le $failedjobs.table.tr.count-1;$i++)
 {
   $test=$failedjobs.table.tr;
   #$failedTimerJobMsg=$failedjobs.table.tr[$i].td[-4];
   $restAllTimerService=$failedjobs.table.tr[$i].td[-5];

   if($failedjobs.table.tr[$i].td[-4] -eq "Failed" -and $failedjobs.table.tr[$i].td[-5] -eq "WOPI Discovery Synchronization") 
   { 
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-1)].InnerText=$msgForWOPITimer;
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)

               Update-SPWOPIProofKey

   }
   elseif(($failedjobs.table.tr[$i].td[-4] -eq "Failed") -and ($failedjobs.table.tr[$i].td[-5] -eq "SP2016 UserProfile Service - Feed Cache Full Repopulation Job")) 
   { 
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-1)].InnerText=$msgForUserProfileServiceFeedCache;
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)

               #$service = get-service -ComputerName $servers[0] -Name "SPTimerV4" #6. SharePoint Timer Service
               

   }
   elseif(($failedjobs.table.tr[$i].td[-4] -eq "Failed") -and ($failedjobs.table.tr[$i].td[-5] -eq "SP2016 UserProfile Service - Feed Cache Repopulation Job")) 
   { 
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-1)].InnerText=$msgForUserProfileServiceFeedCache;
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileOptionalServicStopped)

               #$service = get-service -ComputerName $servers[0] -Name "SPTimerV4" #6. SharePoint Timer Service
               

   }
   
   else 
      { 
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-1)].InnerText=$msgForOtherTimer;
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-1)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-2)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-3)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-4)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-5)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)
               $failedjobs.table.tr[$i].ChildNodes[($failedjobs.table.tr[$i].ChildNodes.Count-6)].SetAttribute('bgcolor', $colorWhileMandatoryServicStopped)

   }

}

#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage
}
catch
{
   $ErrorMessage = "Error in failed timer job status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}

$body = @"
<H2></H2>
$($failedjobs.innerxml)
"@

# Convert to HTML and save the file
#ConvertTo-Html -Head $head -Body $body |
      #Out-File $filename

      $htmlFileFailedJob=ConvertTo-Html -Head $head -Body $body
      $failedjobsModifed=$htmlFileFailedJob;



#SharePoint Solution Status
#$SPSolutions = Get-SPSolution | Select Name , Deployed , Status | ConvertTo-Html -Fragment
#=======================#
# Table for html output #
#=======================#
<#$legendStat="The checks are considered 'Green' in terms of RAG Status when status values are one of the following: <br>
<b>OK, Online, Ready, Idle, CrawlStarting, Active, True</b>"#>

$firstServer=$servers[0]
$secondServer=$servers[1]
$thirdServer=$servers[2]


#<font color = blue><H4><B>Note :</B></H4></font>$legendStat

$testFile=ConvertTo-Html -Body "

<font color = brown><H2><B><br>Server Report:</B></H2></font>

<font color = blue><H4><B>Failed Timer Jobs, Last 7 days</B></H4></font>$failedjobsModifed
<font color = blue><H4><B>SharePoint Server $firstServer Core Services Status</B></H4></font>$coreservicesInServer1Modifed
<font color = blue><H4><B>SharePoint Server $secondServer Core Services Status</B></H4></font>$coreservicesInServer2Modifed
<font color = blue><H4><B>Office Online Server $thirdServer Core Services Status</B></H4></font>$coreservicesInServer3Modifed

<font color = blue><H4><B>Server Memory Usage </B></H4></font>$MemoryModified

<font color = blue><H4><B>Disk Usage $firstServer</B></H4></font>$diskReportInServer1Modified
<font color = blue><H4><B>Disk Usage $secondServer</B></H4></font>$diskReportInServer2Modified
<font color = blue><H4><B>Uptime</B></H4></font>$FarmuptimeModified

<font color = brown><H2><B><br>Application Report:</B></H2></font>

<font color = blue><H4><B>Farm Status</B></H4></font>$SPFarmModified
<font color = blue><H4><B>Web Application POOL Status</B></H4>$WAppPoolModified 
<font color = blue><H4><B>Web Application Status</B></H4>$WebApplicationModified
<font color = blue><H4><B>IIS: Current number of active connections</B></H4></font>$IIS1Modified
<font color = blue><H4><B>Service Application POOL Status</B></H4>$SAppPoolModified
<font color = blue><H4><B>Service Application Status</B></H4>$ServiceAppplicationModified
<font color = blue><H4><B>Service Application Proxy Status</B></H4>$ApplicationProxyModified
<font color = blue><H4><B>Search Administration Component Status</B></H4></font>$SPSearchAdminComponentModified
<font color = blue><H4><B>Search Query Topology Status</B></H4></font>$SPQueryTopologyModified
<font color = blue><H4><B>Content Sources Status with Crawl Log Counts</B></H4></font>$SPContentSourceModified
<font color = blue><H4><B>Crawl Component Status</B></H4>$SPCrawlComponentModified
<font color = blue><H4><B>Content Database Status</B></H4></font>$DBCheckModified

" -Title "SharePoint Farm Health Check Report" -head $head | Out-File $outFileLocation

#-Title "SharePoint Farm Health Check Report" -head $head | Out-File $outFileLocation

#| Out-File $outFileLocation

#=================================#
# Send email to SharePoint Admins #
#=================================#
Function logstamp {
$now=get-Date
$yr=$now.Year.ToString()
$mo=$now.Month.ToString()
$dy=$now.Day.ToString()
$hr=$now.Hour.ToString()
$mi=$now.Minute.ToString()
if ($mo.length -lt 2) {
$mo="0"+$mo #pad single digit months with leading zero
}
if ($dy.length -lt 2) {
$dy="0"+$dy #pad single digit day with leading zero
}
if ($hr.length -lt 2) {
$hr="0"+$hr #pad single digit hour with leading zero
}
if ($mi.length -lt 2) {
$mi="0"+$mi #pad single digit minute with leading zero
}
echo $dy-$mo-$yr
}


try
{

#$output = preg_replace('/(<[^>]+) style=".*?"/i', '$1', $outFileLocation);

$Report = Get-Content $outFileLocation

#Send-MailMessage -To $To -SmtpServer $SMTP -From $From -Subject $Subject -BodyAsHtml "$Report" -Cc $Cc

Send-MailMessage -To $To -SmtpServer $SMTP -From $From -Subject $Subject -BodyAsHtml "$Report" -Cc $Cc

$logMessage="Server report email has been successfully sent to:  "+$To+ ", "+$Cc
#Write-Host $logMessage -BackgroundColor DarkGreen
Write-Log $logMessage 

}
catch
{
   $ErrorMessage = "Error in failed timer job status report test: "+$_.Exception.Message
   #Write-Host $ErrorMessage -BackgroundColor Red
   Write-Log $ErrorMessage
}

