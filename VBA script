Option Explicit
Sub Combine1Folder()
'Copy one wb to the other in folder, assuming header in both
'automatically path
Dim mypath As String, mypath1 As String, file As String, Name As String, period As String
Dim fcell As String 'This is the first cell with content, i.e. for file with header, this would be A2, without header would be A1
Dim fsize As Double, x As Long, i As Long
Dim del As Range, rng As Range, empt As Range
Dim firstAddress As String
Dim co As String, per As String, per_name As String
Dim y As Integer, l As Integer
'y is the position of underscore in file name, l is the length of company name co in file name,
'l = y - 20 (20 is the starting postiion of company name co)


per_name = Application.InputBox("Please enter quarter (e.g. 2015Q2)")
per = Application.InputBox("Please enter date (e.g. 2015-06-30)")

mypath1 = Application.InputBox("Please enter path to files (e.g. C:\Files).")
mypath = mypath1 & "\"

file = Dir(mypath & "*.xls")

'first cell always A2
fcell = "A2"

Dim main As Workbook, wb As Workbook
Dim lrow As Double, lcol As Double, mrow As Double, mcol As Double

Set wb = Nothing
Set main = Nothing

On Error GoTo ErrMsg

'First, open all files in the folder
While file <> ""
    Workbooks.Open (mypath & file)
    file = Dir
Wend

'Combine once all wbs are open
For Each wb In Workbooks
    If wb.Name <> "PERSONAL.XLSB" Then
    
        'all files names consistent with underscore
        y = InStr(20, wb.Name, "_")
        l = y - 20
        co = Mid(wb.Name, 20, l)


        If main Is Nothing And IsEmpty(wb.Worksheets(1).Range(fcell).Value) = False Then
            Set main = wb

'Add header
            main.Worksheets(1).rows(2).Insert
            main.Worksheets(1).Range("A2").Value = co
            main.Worksheets(1).Range("B2").NumberFormat = "yyyy-mm-dd"
            main.Worksheets(1).Range("B2").Value = per

'Find last row
            mrow = main.Worksheets(1).Range("A:A").Find("*", [A1], , , xlByRows, xlPrevious).row
            'mcol = main.Worksheets(1).Cells.Find("*", [a1], , , xlByColumns, xlPrevious).Column

        ElseIf Not main Is Nothing And IsEmpty(wb.Worksheets(1).Range(fcell).Value) = False Then
            lrow = wb.Worksheets(1).Range("A:A").Find("*", [A1], , , xlByRows, xlPrevious).row
            lcol = wb.Worksheets(1).Cells.Find("*", [A1], , , xlByColumns, xlPrevious).Column
            
            main.Worksheets(1).Range("A" & mrow + 1).Value = co
            main.Worksheets(1).Range("B" & mrow + 1).NumberFormat = "yyyy-mm-dd"
            main.Worksheets(1).Range("B" & mrow + 1).Value = per

'Copy and paste to main without touching source files
            wb.Worksheets(1).Range(fcell, wb.Worksheets(1).Cells(lrow, lcol)).Copy _
            Destination:=main.Worksheets(1).Range("A" & mrow + 2) '+ 2 to paste to the next empty line
                'update mrow, lrow was last row with data in wb, hence minus 1 to not count header from wb as it was
                'not pasted. However, since we added header of name+period between files....
            mrow = mrow + lrow

            wb.Close savechanges:=False
            
        ElseIf IsEmpty(wb.Worksheets(1).Range(fcell).Value) = True Then
            wb.Close savechanges:=False
        
        End If
    End If
    
Next wb
        
main.Worksheets(1).Range("E:E, O:O, Y:Y, AE:AE, AG:AG").NumberFormat = "dd-mmm-yy"

'Save with period as file name
main.Worksheets(1).SaveAs mypath & "csv\" & per_name, xlCSV
main.Close savechanges:=False

Exit Sub

ErrMsg:

MsgBox ("Please verify that the files are consistent with the standard format."), , "File Error"

End Sub

