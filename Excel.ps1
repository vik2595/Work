$objExcel=New-Object -ComObject Excel.Application
$objExcel.Visible=$True
$Mdir = Get-ChildItem 'C:\Users\acct\Desktop\Vikas Data\Sales\POS Generated Report\' -Directory | sort LastWriteTime | select -last 1
$dir = Get-ChildItem 'C:\Users\acct\Desktop\Vikas Data\Sales\POS Generated Report\' -Recurse -Directory | sort LastWriteTime | select -last 1
$file = Get-ChildItem 'C:\Users\acct\Desktop\Vikas Data\Sales\POS Generated Report\' -Recurse -File | sort LastWriteTime | select -last 1
$dir = Get-ChildItem 'C:\Users\acct\Desktop\Vikas Data\Sales\POS Generated Report\' -Recurse -Directory | sort LastWriteTime | select -last 1
$Mdir = Get-ChildItem 'C:\Users\acct\Desktop\Vikas Data\Sales\POS Generated Report\' -Directory | sort LastWriteTime | select -last 1
$workbook = $objExcel.Workbooks.Open("C:\Users\acct\Desktop\Vikas Data\Sales\POS Generated Report\$Mdir\$dir\$file")
$worksheet = $workbook.worksheets.item('Sheet1')
$max = $worksheet.UsedRange.Rows.Count
Function DelLastRow
{
for ($i = $max; $i -ge 0; $i--) {
    If ($worksheet.Cells.Item($i, 1).text -eq "") {
        $Range = $worksheet.Cells.Item($i, 1).EntireRow
        [void]$Range.Delete()
        echo $i
    } 
    Else {$i = 0}
}
}
