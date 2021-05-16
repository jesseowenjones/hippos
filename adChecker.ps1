#Convert to excel file to csv
Function ExcelToCsv ($File) {
    $myDir = pwd
    $myDir = $myDir.path
    $excelFile = "$myDir\" + $File + ".xlsx"
    $Excel = New-Object -ComObject Excel.Application
    $wb = $Excel.Workbooks.Open($excelFile)
	
    foreach ($ws in $wb.Worksheets) {
        $ws.SaveAs("$myDir\" + "report.csv", 6)
    }
    $Excel.Quit()
}

$FileName = "rawReport"
ExcelToCsv -File $FileName
Write-Output "Excel file converted to CSV successfully."

#get current folder location
$folder = pwd
$folder = $folder.path

#this assumes that the csv file is in the same folder and is called "report.csv"
#load csv into script - we also need to provide headers
$csv = Import-Csv $folder"\report.csv" -header saleNo,stockNo,desc,qty,price,1,2,3,4,5,6,7,8,9

#now we need to grab out just the stock numbers from the list
$stockNumbers = $csv.stockNo

#need a counter in the foreach
$count = 0;

#setup output file
$outFile = $folder + "\results.txt"
Set-Content -Path $outFile -Value "" #here we are resetting the file

#now we have to split the sale time out from the stock number
foreach ($stock in $stockNumbers)
{
    $stock = $stockNumbers[$count].Split(" ")[2]
    $count++
    if ($stock-contains("SWSN")) 
    {
        continue
    } 
    $url = "https://www.gumtree.com.au/s-caboolture-south-sunshine-coast/$stock/k0l3006282r5"
    $page = iwr $url
    $pos = $page.Content.IndexOf("zeroSearchResults") + 19
    $deleted = $page.Content.Substring($pos, 1)
    if ($deleted-eq("t")) 
    {
        Add-Content -Path $outFile -Value "$stock - removed"
        Add-Content -Path $outFile -Value " "
    }
    else 
    {
        Add-Content -Path $outFile -Value "$stock - NOT removed. Link:"
        Add-Content -Path $outFile -Value "$url"
        Add-Content -Path $outFile -Value " "
    }
    Write-Output "Stock number checked."
}
