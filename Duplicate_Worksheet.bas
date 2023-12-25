'First page will be draft table (or anything)
'Second page is used to specify sheetnames 

Sub Create_UPB()
Dim sh1 As Worksheet, sh2 As Worksheet, c As Range
Set sh1 = Sheets("Sample")
Set sh2 = Sheets("WorkSheet")
    For Each c In sh2.Range("A2:A4")
        sh1.Copy after:=Sheets(Sheets.Count)
        ActiveSheet.Name = c.Value
        

          Next
        
 End Sub
