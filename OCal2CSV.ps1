$DebugPreference = "Continue"
$VerbosePreference = "SilentlyContinue"


Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -comobject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")
$calendar =
  $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
$items = $calendar.items
write-debug ("#of Items : {0}" -f $items.count)

$start = Get-Date
$end = $start.AddDays(60)
# query string see also <https://msdn.microsoft.com/en-us/library/office/ff869597.aspx>
#single quoteを忘れたために、実行時エラーになっていた
$condition = "[Start] >= '{0:MM/dd/yyyy} 00:00 am' AND [End] < '{1:MM/dd/yyyy} 00:00 am'" -f $start, $end
write-debug $condition 

$items.IncludeRecurrences = $true
$items.Sort("[Start]")
$filteredItems = $items.Restrict($condition) | Select-Object -Property subject,location,Organizer,StartInStartTimeZone,EndInEndTimeZone
write-debug ("#of Items : {0}" -f $filteredItems.count)

Write-Verbose "export items..."
$filteredItems | 
export-csv -Encoding utf8 -path c:\tmp\ocal.csv -NoTypeInformation
Write-Verbose "done."