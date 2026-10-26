# Excel VBA Automation Showcase

A collection of advanced Excel and VBA solutions I have built to automate business processes, primarily in the financial and retail sectors. 

## Project 1: Cash Sales Billing System (VBA)

### The Business Problem
A retail business was manually typing cash sale bills into Excel. This led to human errors, duplicate bill numbers, and incorrect running totals at the end of the month. They needed a secure, automated way to generate, search, and edit bills while ensuring the running total was always accurate.

### The Solution
I built a custom Excel VBA UserForm that handles the entire billing lifecycle.

**Key Features:**
- **Automated Sequential Bill Numbering:** Generates unique bill numbers (e.g., `CB202610-001`) based on the selected date, automatically checking for existing bills to prevent duplicates.
- **Data Validation:** Ensures invalid dates cannot be entered and that users cannot bypass the form using the "X" button.
- **Dynamic Running Totals:** Automatically recalculates the running monthly total whenever a bill is added or edited, ensuring financial accuracy.
- **Search & Edit:** Allows users to search for an existing bill by number, pulling the data back into the form for editing without breaking the sequence.

### Code Highlight: Automated Bill Number Generation
When the user selects a date, the VBA script searches the hidden column `F` (which stores Year & Month). If it finds existing bills for that month, it increments the sequence. If not, it starts at `001`.

```vba
Private Sub txtDate_AfterUpdate()
    Dim seqnum As Variant
    Dim fndCashBilDate, fndCashBilNo
    Dim rowF As Long
    Dim freeaddress
    Dim ws As Worksheet
    seqnum = 2
    If txtDate.Value = "" Or txtDate.Value < #1/1/2019# Or Not IsDate(txtDate.Value) Then     '--- (3)
    MsgBox "Invalid date, Please re-enter again", Title:="Invalid Date"
    txtDate.SetFocus
    Exit Sub
    End If
    rowF = Range("F1048576").End(xlUp).Row

    Set ws = Worksheets("CashSalesBill")
    With Range("F3:F" & rowF)
    Set fndCashBilDate = .Find(what:=DatePart("yyyy", txtDate) & DatePart("m", txtDate), after:=ws.Range("F3"), _
    LookIn:=xlValues, searchorder:=xlByRows, searchdirection:=xlNext, _
    MatchCase:=False)
    If fndCashBilDate Is Nothing Then
    If Len(DatePart("m", txtDate.Value)) = 1 Then
    lbBillNoSeq.Caption = "CB" & DatePart("yyyy", txtDate.Value) _
    & "0" & DatePart("m", txtDate.Value) & "-001"
    Else
    lbBillNoSeq.Caption = "CB" & DatePart("yyyy", txtDate.Value) _
    & DatePart("m", txtDate.Value) & "-001"
    End If
    Set fndCashBilDate = Nothing
    Set ws = Nothing
    Exit Sub
    End If
    End With
    With Range("F3:F" & rowF)
    Set fndCashBilNo = .Find(what:=DatePart("yyyy", txtDate) & DatePart("m", txtDate), after:=ws.Range("F3"), _
    LookIn:=xlValues, searchorder:=xlByRows, searchdirection:=xlNext, _
    MatchCase:=False)
    If Not fndCashBilNo Is Nothing Then
    freeaddress = fndCashBilNo.Address
    Do
    Set fndCashBilNo = .FindNext(fndCashBilNo)
    If Not fndCashBilNo.Address = freeaddress Then
    seqnum = seqnum + 1
    End If
    Loop While Not fndCashBilNo Is Nothing And fndCashBilNo.Address <> freeaddress
    Set fndCashBilNo = Nothing
    Set ws = Nothing
    End If
    End With
    If Len(seqnum) = 1 Then
    If Len(DatePart("m", txtDate.Value)) = 1 Then
    lbBillNoSeq.Caption = "CB" & DatePart("yyyy", txtDate.Value) _
    & "0" & DatePart("m", txtDate.Value) & "-00" & seqnum
    Else
    lbBillNoSeq.Caption = "CB" & DatePart("yyyy", txtDate.Value) _
    & DatePart("m", txtDate.Value) & "-00" & seqnum
    End If
    ElseIf Len(seqnum) = 2 Then
    If Len(DatePart("m", txtDate.Value)) = 1 Then
    lbBillNoSeq.Caption = "CB" & DatePart("yyyy", txtDate.Value) _
    & "0" & DatePart("m", txtDate.Value) & "-0" & seqnum
    Else
    lbBillNoSeq.Caption = "CB" & DatePart("yyyy", txtDate.Value) _
    & DatePart("m", txtDate.Value) & "-0" & seqnum
    End If
    ElseIf Len(seqnum) = 3 Then
    If Len(DatePart("m", txtDate.Value)) = 1 Then
    lbBillNoSeq.Caption = "CB" & DatePart("yyyy", txtDate.Value) _
    & "0" & DatePart("m", txtDate.Value) & "-" & seqnum
    Else
    lbBillNoSeq.Caption = "CB" & DatePart("yyyy", txtDate.Value) _
    & DatePart("m", txtDate.Value) & "-" & seqnum
    End If
    End If
End Sub
```
### Code Highlight: Automated Bill Number Generation
Before saving an edit, the script verifies the bill number exists, updates the row, and loops through the rows to recalculate the running total for that specific month.

```vba
Private Sub cmdEdit_Click()
  Dim rowA, rowB, brow, Crow As Long
  Dim r, s, t As Long
  Dim Drow
  Dim FinDup
  Dim ws As Worksheet
  Application.ScreenUpdating = False
  Set ws = Worksheets("CashSalesBill")
  rowA = Range("A1048576").End(xlUp).Row
  With ws.Range("B3", Cells(rowA, "B"))
  Set FinDup = .Find(what:=ActiveCell.Value, after:=Range("B3"), _
  LookIn:=xlValues, Lookat:=xlWhole, searchorder:=xlByRows, _
  searchdirection:=xlNext, MatchCase:=False)
  If Not FinDup Is Nothing Then
  ActiveCell.Offset(0, -1) = txtDate.Value
  ActiveCell.NumberFormat = "dd/mm/yyyy;@"
  ActiveCell.Value = lbBillNoSeq.Caption
  ActiveCell.NumberFormat = "General"
  ActiveCell.EntireColumn.AutoFit
  ActiveCell.Offset(0, 1) = txtDescription.Value
  ActiveCell.Offset(0, 1).EntireColumn.AutoFit
  ActiveCell.Offset(0, 2) = txtAmount.Value
  ActiveCell.Offset(0, 2).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
  ActiveCell.Offset(0, 2).EntireColumn.AutoFit
  ActiveCell.Offset(0, 3).EntireRow.AutoFit
  ActiveCell.Offset(0, 3).EntireColumn.AutoFit
  End If
  t = 0 '------ Added
  If ActiveCell.Offset(t, 4).Value <> ActiveCell.Offset((t - 1), 4).Value Then '------   Added
  ActiveCell.Offset(0, 3).Value = txtAmount.Value
  s = 1
  Do
  ActiveCell.Offset(s, 3).Value = ActiveCell.Offset(s, 2) + ActiveCell.Offset(s - 1, 3)
  'r = r + 1
  s = s + 1
  Loop Until ActiveCell.Offset(s, 4).Value <> ActiveCell.Offset((s - 1), 4).Value
  Else
  r = 0
  s = 1
  Do
  ActiveCell.Offset(r, 3).Value = ActiveCell.Offset(r, 2) + ActiveCell.Offset(s - 2, 3)
  r = r + 1
  s = s + 1
  Loop Until ActiveCell.Offset(r, 4).Value <> ActiveCell.Offset((r - 1), 4).Value
  End If
  End With
  ws.Range("A1048576").End(xlUp).Select
  brow = ActiveCell.Row
  ws.Range("B1048576").End(xlUp).Select
  Crow = ActiveCell.Row
  With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
  .SortFields.Clear
  .SortFields.Add Key:=Range("B3", Cells(brow, "B")), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  .SortFields.Add Key:=Range("A3", Cells(Crow, "A")), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  .SetRange Range("A2", Cells(Crow, "D"))
  .Header = xlYes
  .MatchCase = False
  .Orientation = xlTopToBottom
  .SortMethod = xlPinYin
  .Apply
  End With
  Set Drow = ws.Range("D3", Cells(Crow, "D"))
  ws.Range("C" & (Crow + 1)) = "Total Cash Sales"
  ws.Range("D" & (Crow + 1)) = Application.WorksheetFunction.Sum(Drow)
  ws.Range("D" & (Crow + 1)).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
  Call UserForm_Initialize
  rowB = ws.Range("A1048576").End(xlUp).Row
  ws.Range("A" & rowB).Select
  Application.ScreenUpdating = True
  Set ws = Nothing
  ActiveWorkbook.Save
  cmdPosting.Enabled = True
  Unload Me
End Sub
```
### How to view the full code
The full VBA module and UserForm logic are available in the /vba_code folder of this repository.

🔗 [LinkedIn](https://www.linkedin.com/in/chin-kee-ming-588685148)
💻 [GitHub](https://github.com/chinkm/excel-vba-automation-showcase)
