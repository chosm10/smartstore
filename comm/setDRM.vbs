Set args = WScript.Arguments
arg1 = args.Item(0)
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = False 
Set objWorkbook = objExcel.Workbooks.Open(arg1)
objWorkbook.Save
objWorkbook.Close 
objExcel.Quit