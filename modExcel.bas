Attribute VB_Name = "modExcel"
Option Explicit

Public xlApp                    As New Excel.Application
Public xlBook                   As Excel.Workbook
Public xlSheet                  As Excel.Worksheet

'**************************************************************************
' Cute little function that takes a recordset and FilePathName and generates
' an Excel spreadsheet from the recordset. Nice.....
'
' The Gazman - Aug 2002
'**************************************************************************

Public Function SaveAsExcel(rsErr As ADODB.Recordset, sFileName As String, _
            sSheet As String, sOpen As String)
Dim fd          As Field
Dim Cellcnt     As Integer
Dim I           As Integer
Dim x           As Integer
Dim S           As Integer

On Error GoTo Err_Handler

Screen.MousePointer = vbHourglass

Set xlApp = New Excel.Application
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

'Get the field names
Cellcnt = 1
xlSheet.Name = sSheet
xlSheet.Cells(1.3, 1).Value = sSheet
xlSheet.Cells(1.3, 1).Font.Size = 8
xlSheet.Cells(1.3, 1).Font.Bold = True

Dim formC As Form
'1 = IRR
'2 = canvass
'3 = po records
'4 =  pr details

If who = 3 Then who3
If who = 4 Then who4


xlSheet.SaveAs sFileName ' Save the Worksheet.
xlBook.Close ' Close the Workbook
xlApp.Quit ' Close Microsoft Excel with the Quit method.

If sOpen = "YES" Then ' Open Excel Workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(sFileName)
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Application.Visible = True
Else
    Set xlApp = Nothing ' Release the Excel objects.
    Set xlBook = Nothing
    Set xlSheet = Nothing
End If

Screen.MousePointer = vbDefault
Err_Handler:
    If Err = 0 Then
        Screen.MousePointer = vbDefault
    Else
        MsgBox "An error has occurred! " & vbCrLf & vbCrLf & Err & ":" & Error & " ", vbExclamation
        Screen.MousePointer = vbDefault
    End If
End Function



Private Sub who4()
With FrmCutoff.ListView1
Dim countCol
countCol = 0
Dim I
Dim Cellcnt
Dim w
w = 1
I = 1
Cellcnt = 1
Do While w <= .ColumnHeaders.Count
    xlSheet.Cells(I, Cellcnt).Value = .ColumnHeaders(w).Text
    xlSheet.Cells(I, Cellcnt).Interior.ColorIndex = 33
    xlSheet.Cells(I, Cellcnt).Font.Size = 8
    xlSheet.Cells(I, Cellcnt).Font.Bold = True
    xlSheet.Cells(I, Cellcnt).BorderAround xlContinuous
    countCol = countCol + 1
   Cellcnt = Cellcnt + 1
   w = w + 1
Loop
w = 1


I = 2
Cellcnt = 1

Do While w <= .ListItems.Count
    
    xlSheet.Cells(I, Cellcnt).Value = .ListItems(w).Text
   Cellcnt = Cellcnt + 1
    Dim h
    h = 1
    
    Do While h <= (.ColumnHeaders.Count - 1)
    xlSheet.Cells(I, Cellcnt).Value = .ListItems(w).SubItems(h)
    xlSheet.Cells(I, Cellcnt).Font.Size = 8
    h = h + 1
     Cellcnt = Cellcnt + 1
    Loop
    Cellcnt = 1
    
     I = I + 1
    w = w + 1
Loop
w = 1
Cellcnt = 1
Do While w <= .ColumnHeaders.Count

    xlSheet.Columns(Cellcnt).AutoFit
    Cellcnt = Cellcnt + 1
    w = w + 1
Loop
End With
End Sub
Private Sub who3()
With FrmCutoff.ListView2
Dim countCol
countCol = 0
Dim I
Dim Cellcnt
Dim w
w = 1
I = 1
Cellcnt = 1
Do While w <= .ColumnHeaders.Count
    xlSheet.Cells(I, Cellcnt).Value = .ColumnHeaders(w).Text
    xlSheet.Cells(I, Cellcnt).Interior.ColorIndex = 33
    xlSheet.Cells(I, Cellcnt).Font.Size = 8
    xlSheet.Cells(I, Cellcnt).Font.Bold = True
    xlSheet.Cells(I, Cellcnt).BorderAround xlContinuous
    countCol = countCol + 1
   Cellcnt = Cellcnt + 1
   w = w + 1
Loop
w = 1


I = 2
Cellcnt = 1

Do While w <= .ListItems.Count
    
    xlSheet.Cells(I, Cellcnt).Value = .ListItems(w).Text
   Cellcnt = Cellcnt + 1
    Dim h
    h = 1
    
    Do While h <= (.ColumnHeaders.Count - 1)
    xlSheet.Cells(I, Cellcnt).Value = .ListItems(w).SubItems(h)
    xlSheet.Cells(I, Cellcnt).Font.Size = 8
    h = h + 1
     Cellcnt = Cellcnt + 1
    Loop
    Cellcnt = 1
    
     I = I + 1
    w = w + 1
Loop
w = 1
Cellcnt = 1
Do While w <= .ColumnHeaders.Count

    xlSheet.Columns(Cellcnt).AutoFit
    Cellcnt = Cellcnt + 1
    w = w + 1
Loop
End With
End Sub
