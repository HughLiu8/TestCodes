	try
	{
		$xl = New-Object -COM "Excel.Application"
		$xl.Visible = $true
		$wb = $xl.Workbooks.Open("C:/BibleGroup/Test.xlsx")
		$ws = $wb.Sheets.Item(3)
 
        
        $ws.Cells.Item(2, 2) = "567"

		$rows = $ws.UsedRange.Rows.Count
		$first = 1
		for ($i=2; $i -le $rows; $i++){
            $v1 = $ws.Cells.Item($i, 1).text 

		}
		
        $wb.Save()
		$wb.Close()
        
		$xl.Quit()
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
	}
	catch
	{
		
	}