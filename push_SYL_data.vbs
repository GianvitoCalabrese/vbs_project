Dim lookfor 
Dim table_array 'table_array
Dim varResult  'varResult
Dim table_array_col  'table_array_col
Dim lookFor_col  'lookFor_col
Dim Wbk 
Dim rng: 
Dim i 
Dim cell 
Dim Application

Set Application = CreateObject("Excel.Application")
Set rng = Application.Range("A2:A161")

For i = 1 To rng.Rows.Count
    Set lookfor = rng.Cells(RowIndex=i, ColumnIndex="A")
    Set Wbk = Workbooks.Open("//theconnection.onsemi.com/GSCO/LP/GFO/Shared Documents/Foundry Notes/LFOUNDRY/Other/SBL's and SYL's/SYL Limits Q2 2021.xlsx")
    Set table_array = Wbk.Sheets("Summary").Range("H3:J134")
    
    For Each cell In table_array.Columns(1).Cells
        cell.Value = Trim(cell.Value)
    Next
    
    table_array_col = 2  'pull data on this column
    varResult = Application.VLookup(lookfor.Value, table_array, table_array_col, 0)
    lookFor_col = 3  'lookFor.Value starting from 0/1/2/3 where to start from column
    lookfor.Offset(0, lookFor_col) = varResult
Next
     
Workbooks("SYL Limits Q2 2021.xlsx").Close SaveChanges=False