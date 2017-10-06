Attribute VB_Name = "niin"
Option Explicit

Sub init()
Dim desc As String, cat As Byte, macroName As String
desc = "Adds leading zeroes to NIINs due to incorrect formatting and extracts NIINS from NSNs " & vbCrLf & _
       "NIIN(<Cell Range>)"
cat = 7
macroName = "NIIN"
control.RegisterUDF macroName, desc, cat
End Sub

Function NIIN(CELL As Range) As String
Attribute NIIN.VB_Description = "Adds leading zeroes to NIINs due to incorrect formatting and extracts NIINS from NSNs \r\nNIIN(<Cell Range>)"
Attribute NIIN.VB_ProcData.VB_Invoke_Func = " \n7"
'if length of range is <= 9 (Badly formatted NIIN), then reformat with leading zero
'if legnth of range is > 9 (NSN) then extract NIIN or last 9
  NIIN = formatNIIN(CELL)
End Function

Function NIIN_Help() As String
NIIN_Help = "Adds leading zeroes to NIINs due to incorrect formatting and extracts NIINS from NSNs " & vbCrLf & _
       "NIIN(<Cell Range>)"
End Function


Sub formatColumnAsNIIN()
'get single column as input, insert new column (to right), reformat as text then get NIINS
'same logic as NIIN function
Dim wb As Workbook, ws As Worksheet
Dim firstRow As Byte, lastRow As Long, i As Long
Dim targetColumn As Byte, destinationColumn As Byte
Dim header As Byte, overwriteNewColumn As Byte
Dim rng As Range, newRng As Range
Dim columnLetter As String, newHeader As String, currentHeader As String

Set wb = ActiveWorkbook
'if !calledFromFormula then get input; else use target column
On Error Resume Next
Set rng = Application.InputBox("Select (click on ) the column that you'd like to reformat as NIINs" & _
                               vbCrLf & vbCrLf & "      (e.g. Click on Column A or Cell B4, etc...)    ", _
                               "Column Selection", Type:=8)
If rng Is Nothing Then control.cleanExit


Set ws = rng.Parent

header = MsgBox("Does your spreadsheet contain headers?", vbYesNo, "Header Indicator")
columnLetter = Mid(rng.Address, 2, 1)
overwriteNewColumn = MsgBox("Would you like to overwrite your values in Column " & columnLetter & _
                            "?" & vbCrLf & vbCrLf & "    Selecting 'No' will insert a new Column.", _
                            vbYesNo, "Override Selection")
control.prepAndCleanup

'get destination column
targetColumn = rng.column
If overwriteNewColumn = 6 Then
    destinationColumn = targetColumn
ElseIf overwriteNewColumn = 7 Then
    ws.Columns(targetColumn + 1).Insert (xlShiftToRight)
    destinationColumn = targetColumn + 1
End If

'assumes no header
If header = 7 Then
    firstRow = 1
ElseIf header = 6 Then
    firstRow = 2
End If

'update column header
If header = 6 And overwriteNewColumn = 7 Then
    currentHeader = ws.Cells(1, targetColumn).Value
    If Left(currentHeader, 4) <> "ORIG" Then ws.Cells(1, targetColumn).Value = "ORIG_" & currentHeader
    If Left(ws.Cells(1, destinationColumn).Value, 5) <> "FIXED" Then
        newHeader = "FIXED_" & currentHeader
        ws.Cells(1, destinationColumn).Value = newHeader
    End If
End If

ws.Columns(destinationColumn).NumberFormat = "@"
lastRow = ws.Cells(Rows.Count, targetColumn).End(-4162).Row

For i = firstRow To lastRow
    ws.Cells(i, destinationColumn) = formatNIIN(ws.Cells(i, targetColumn))
Next i

ws.Cells.EntireColumn.AutoFit

control.prepAndCleanup True

End Sub


Function formatNIIN(rng As Range) As String
'if length of range is <= 9 (Badly formatted NIIN), then reformat with leading zero
'if legnth of range is > 9 (NSN) then extract NIIN or last 9
  If Len(rng.Value) < 10 Then
      formatNIIN = Application.WorksheetFunction.Text(rng, "000000000")
  Else: formatNIIN = Right(rng.Value, 9)
  End If
End Function

