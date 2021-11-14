########################################################################
# Script for collecting performance report of all powered on VMs       #
########################################################################

# $vCenter = Read-Host "Enter the vCenter name wherein the target cluster resides"
 
 $vms_pwrod = Get-VM | Where-Object {$_.PowerState -eq "PoweredOn"} | select Name
 $sdate = '08/01/2021'  
 $fdate = '09/01/2021'
 $nsamples = 31
 [int]$a = 1
 [int]$b = 1
 $vms_xls = 1 
 $Addtitle = $nsamples + 1
 $nodata='No data'
 $nodatavm = 0
 $VMs = $vms_pwrod.name
 $stats = @()
 $mem_stats = @()

#Connect-VIServer $vCenter

$clusters = Get-Cluster
$vm_name = Get-Cluster $clusters | Get-VM| Where-Object PowerState -eq PoweredOn

#======================================================================== 
# CPU and Mem average...
#========================================================================

 foreach ($cluster in $clusters) {
   foreach ($vm_name in $vm_name) {

     Write-Host "Collecting data for $vm_name …"
     $wkcluster = ($cluster | select Name).Name
      ## ===  Gather CPU average use =================== 
     $stats = Get-Stat –Entity $vm_name  -Stat 'cpu.usage.average' –Start $sdate –Finish $fdate 
     if (!$stats) {
         $stats+= $vm_name.Name+"`r"   
         $stats+= $stats| Export-Excel -Append –Path .\PerfResult.xlsx -WorksheetName $wkcluster" No Data"
         $nodatavm++
     }else {
       $stats +=""
       $stats = $stats | Select Entity,Timestamp,Value |Sort Timestamp -Verbose| Export-Excel -Append –Path .\PerfResult.xlsx -WorksheetName $wkcluster" Cpu"
     }
     ## ===  Gather Memory average use =================== 
     $stats_mem = Get-Stat –Entity $vm_name -Stat 'mem.usage.average' –Start $sdate –Finish $fdate 
     $stats_mem +=""
     $stats_mem = $stats_mem | Select Entity,Timestamp,Value |Sort Timestamp -Verbose | Export-Excel -Append –Path .\PerfResult.xlsx -WorksheetName $wkcluster" Memory"
  
     $number_vm++ 
     $stats =""
      
 }



    #========================================================================= 
    # Open Excel
    #=========================================================================
     $excel = New-Object -ComObject Excel.Application 
     $excel.Visible = $false
     $wb = $excel.Workbooks.Open("C:\ibm_apar\PerfGraph\PerfResult.xlsx")
     
     $clusterN = $wkcluster+" Cpu"
   #======= CPU Workbook fix =================================================== 
     $wsData = $wb.WorkSheets.item(1) 
     $wsData.Name = $clusterN
    
   #=======Delete the first row =============================================
     $wsData.Cells.Item(1,1).EntireRow.Delete()
   
 
   #=======Put VM and CPU on every line======================================
   
    Do {
       $wsData.Cells.Item($a,1) = 'VM name'
       $b++
       $wsData.Cells.Item($a,2) = 'Dated'
       $wsData.Cells.Item($a,3) = 'CPU %'
       $vms_xls++
       $a+=$Addtitle
       $b+=$Addtitle
   } While ($vms_xls -le $number_vm)

      
     #======= Memory Workbook fix =================================================== 
     $clusterM = $wkcluster+" Memory"
     $wsData = $wb.WorkSheets.item(2) 
     $wsData.Name = $clusterM
     
     # Reset values    
     $a = 1
     $b = 1
     $vms_xls = 1 
    

    #=======Delete the first row =============================================
     $wsData.Cells.Item(1,1).EntireRow.Delete()
    #=========================================================================
   
    Do {
       $wsData.Cells.Item($a,1) = 'VM name'
       $b++
       $wsData.Cells.Item($a,2) = 'Dated'
       $wsData.Cells.Item($a,3) = 'Memory %'
       $vms_xls++
       $a+=$Addtitle
       $b+=$Addtitle
   } While ($vms_xls -le $number_vm)



   #======= No data Workbook fix =================================================== 
     $clusterNoD = $wkcluster+" No Data"
     $wsData = $wb.WorkSheets.item(3) 
     $wsData.Name = $clusterNoD
     $wsData.Cells.Item(1,1) = 'No data for the following VMs'

   # ======== Save Excel file and quit ====================================
     $wb.Save();
     $excel.Quit()


##################################################################
# Create graph part 2       #
##################################################################


$xlChart=[Microsoft.Office.Interop.Excel.XLChartType]

$xl = new-object -ComObject Excel.Application   
$fileName = 'C:\ibm_apar\PerfGraph\PerfResult.xlsx'
$wb = $xl.Workbooks.Open($fileName)
$wsData = $wb.WorkSheets.item(1) 
$memData = $wb.WorkSheets.item(2)

$topmem = 0
$top = 0
$Left = 0
$total = $number_vm
$loops = 1
$a=1
$b=$nsamples+1
$chart_No=1
$vnm = 2




# Adding a new sheet where the chart would be created
$wsChart = $wb.Sheets.Add();
$wkcluster
$wsChart.Name = $wkcluster

#========================================================================= 
# Loop for CPU graphs
#=========================================================================
$total = $total - $nodatavm
Do{

#Activating the Data sheet
$wsData.activate()

#Selecting the source data - We cn select the first cell with Range and select CurrentRegion which selects theenire table
$DataforCPUChart = $wsData.Range("B"+$a,"C"+$b)

#Adding the Charts
$CPU_Chart = $wsChart.Shapes.AddChart().Chart

# Providing the chart types
$CPU_Chart.ChartType = 4

#Providing the source data
$CPU_Chart.SetSourceData($DataforCPUChart)


# Set it true if want to have chart Title
$CPU_Chart.HasTitle = $true

# Providing the Title for the chart
$vmname = $wsData.Cells.Item($vnm,1).Text
$CPU_Chart.ChartTitle.Text = "Graph for CPU on $vmname for period from  $sdate to $fdate "

# Setting up the position of chart
$wsChart.shapes.item("Chart "+$chart_No).top = $top
$wsChart.shapes.item("Chart "+$chart_No).left = 0
$wsChart.shapes.item("Chart "+$chart_No).Width = 700
$chart_No+=1

#Activating the Data sheet for Memory Graphs
$memData.activate()


#Selecting the source data - We cn select the first cell with Range and select CurrentRegion which selects theenire table
$DataforMemChart = $memData.Range("B"+$a,"C"+$b)

#Adding the Charts
$Mem_Chart = $wsChart.Shapes.AddChart().Chart

# Providing the chart types
$Mem_Chart.ChartType = 4

#Providing the source data
$Mem_Chart.SetSourceData($DataforMemChart)


# Set it true if want to have chart Title
$Mem_Chart.HasTitle = $true

# Providing the Title for the chart
$vmname = $wsData.Cells.Item($vnm,1).Text
$Mem_Chart.ChartTitle.Text = "Graph for Memory on $vmname for period from  $sdate to $fdate " 

# Setting up the position of chart
$wsChart.shapes.item("Chart "+$chart_No).top = $topmem
$wsChart.shapes.item("Chart "+$chart_No).left = 800
$wsChart.shapes.item("Chart "+$chart_No).Width = 700

# Increasing top size for the graphs"
$top+=300
$topmem+=300

#Increasing values for chart and loops
$loops+=1
$a+=$nsamples+1
$b+=$nsamples+1
$chart_No+=1
$vnm += $Addtitle

} While ($loops -le $total)
#========================================================================= 
# End of loops
#=========================================================================

}

# Saving the sheet
$wb.Save();

# Closing the work book and xl
$wb.close() 
$xl.Quit()

# Releasting the com object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
