# Virtual Machine Performance Graphs for CPU and Memory usage per Cluster

## Scope:
Script that exports perforamnce reports for all the powered on vms (CPU, Memory) and creates an excel file with performance tabs per cluster.

## Requirements:
* Windows Server 2012 and above // Windows 10
* Powershell 5.1 and above
* PowerCLI either standalone or import the module in Powershell (Preferred)
* Import-Excel Module 
* MS Excel 

## Configuration
In order to set the days for the monitor change the variable $nsamples . By default is 31 days/
```powershell
 $nsamples = 31
```

To save the excel file to some specific folder modify the following variable
```powershell
$XLS_Path = 'C:\ibm_apar\PerfGraph_'+$Date+'.xlsx'
```
For sending the result file via e-mail you will need to have the SMTP ip to be set instead of "mailServerIP", uncomment the line and set the "From", "To" and "Subject":
```
 send-mailmessage -from "Perf_Mon@Customer.com" -to "joedoe@ibm.com" -subject "Perfomance Charts $(get-date -f "dd-MM-yyyy")" -body "Below you can find the rvtools report. Please see attachment `n `n `n" -Attachments $destination -smtpServer MailServerIP
 ```


## Example
 Run the script
 ```powershell
 # make sure to change the directory in case you are not running the script from C:\
 PS> C:\PerfGraph.ps1 
 ```
The script will run and create 4 tabs for each Cluster. 

- Chart 
  Will give you graphs for Cpu and memory for all the vms that have performnce data for the given timeframe

![Alt text](/screenshot/chart.jpg?raw=true "Main Usage")
 

- CPU 
  Will give you data for Cpu average usage for all the vms that have performnce data , on the given timeframe, each day

![Alt text](/screenshot/cpu.jpg?raw=true "CPU Usage")

- Memory 
  Will give you data for memory average usage for all the vms that have performnce data, on the given timeframe, each day

![Alt text](/screenshot/memory.jpg?raw=true "Memory Usage")

- No Data 
  You will find all the VMs that have perfomance data less than given timeframe. Those Vms are skipped.

![Alt text](/screenshot/nodata.jpg?raw=true "VMs with No data during the timeframe")
## Frequetly Asked Questions:
* When I am executing the script it gives you an error "vCenter not found".
   > Before you execute the script you need first to be connected on a vCenter Server.
   ```powershell
   PS> Connect-VIServer <vCenter-IP-FQDN>
   ```
   
* When I run the script it gives me error on Excel commands
  > You are missing the Excel module. You need to import it prior of running the script.
  ```powershell 
  PS> Install-Module -Name ImportExcel
  ```
