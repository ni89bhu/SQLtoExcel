##############################################################################################

$db = "ETAdmin10047"
$instance = "ETPVMDFHINFOG01\SQLEXPRESS"
$duration =  "3"
##############################################################################################

$12 = $db -replace "ETAdmin",""
$13 = Invoke-Sqlcmd -ServerInstance $instance -Database EventTrackerData -Query "
SELECT        ReportTitle
FROM            tbl_RptQueue
WHERE        (ID = '$12')
"
$14 = $13.reportTitle
$15 = get-date -Format "dd.MMM hh.mm.tt"
$16 = $14 + "-" + $15
##############################################################################################

$16 > "D:\OtherFiles\Elapsed_Time.txt"
Measure-Command {
Send-SQLDataToExcel -Connection "Server=$instance;Database=$db;Trusted_Connection=True;" -MsSQLserver -SQL "SELECT LogTime, Computer, [User Name], [Source IP Address], [Source Port], [Destination IP Address], [Destination Port], [Source Interface], [Destination Interface], Direction, [Service Name], 
                         [Requested URL], [Destination Host Address], Action, Priority, [Category Description], [Message Details], [Critical Level], [Critical Score], [Bytes Sent], [Bytes Received]
FROM            Events
WHERE        (LogTime >= DATEADD(day, -$duration, GETDATE()))
ORDER BY LogTime DESC" -AutoSize -TableStyle "Light13" -Title "$14" -WorksheetName ReportData -IncludePivotTable -PivotRows Computer -PivotData Computer -IncludePivotChart -PivotChartType DoughnutExploded -Path "D:\OtherFiles\$16.xlsx" 
} >> "D:\OtherFiles\Elapsed_Time.txt"
##############################################################################################