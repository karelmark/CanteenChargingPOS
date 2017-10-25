VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSales 
   Caption         =   "PKI Canteen Charging Systems"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   15180
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   8520
      Picture         =   "FrmSales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Import to Excel"
      Top             =   7800
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Range"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   15015
      Begin VB.CommandButton btn_Search 
         Caption         =   "Search"
         Height          =   375
         Left            =   4800
         TabIndex        =   2
         Top             =   300
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   127074305
         CurrentDate     =   41659
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   127074305
         CurrentDate     =   41659
      End
      Begin VB.Label Label5 
         Caption         =   "From :"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "To :"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   10610
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Charge"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Total Sales:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   10
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lbltotalsales 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   9
      Top             =   7800
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Canteen Sales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004000&
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "FrmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub btn_Search_Click()
 
Dim sqlstr As String


'sqlstr = "SELECT tx.* , E.Emplname as Lname, E.empfname as fname FROM tbl_transaction as tx INNER JOIN vwEmployeeMaster as E ON tx.idno = E.empno WHERE tx.transdate BETWEEN " & DateValue(DTPicker1.Value) & " AND  " & DateValue(DTPicker2.Value) & " "

'sqlstr = "SELECT * FROM tbl_transaction   WHERE tx.transdate BETWEEN " & DateValue(DTPicker1.Value) & " AND  " & DateValue(DTPicker2.Value) & "    ORDER BY transno DESC"
  
'sqlstr = "SELECT tx.transdate, tx.transtime,tx.txtotal , E.Emplname as Lname, E.empfname as fname FROM tbl_transaction as tx INNER JOIN vwEmployeeMaster as E ON tx.idno = E.empno  WHERE tx.transdate  >=" & DTPicker1.Value & " AND tx.transdate <=" & DTPicker2.Value & ""
sqlstr = "SELECT tx.transdate, tx.transtime,tx.txtotal  , E.Emplname as Lname, E.empfname as fname   FROM tbl_transaction as tx INNER JOIN vwEmployeeMaster as E ON tx.idno = E.empno    WHERE  tx.transdate between #" & DateValue(DTPicker1.Value) & "# AND #" & DateValue(DTPicker2.Value) & "# ORDER BY transno DESC"
'MsgBox DTPicker1.Value & " " & DTPicker2.Value, vbOKOnly, "show values"

Dim c As Integer

c = ListView2.Width
  
  ListView2.ColumnHeaders.Item(1).Width = c * 0.1
  ListView2.ColumnHeaders.Item(2).Width = c * 0.1
  ListView2.ColumnHeaders.Item(3).Width = c * 0.68
  ListView2.ColumnHeaders.Item(4).Width = c * 0.09
  
 Dim totalsale As Double
 totalsale = 0
  
Set rsac7 = Nothing
  

rsac7.Open sqlstr, acaccess
    
    If Not rsac7.EOF Then
    
            ListView2.ListItems.Clear
            
            Do While Not rsac7.EOF
            
            totalsale = totalsale + rsac7.Fields!txtotal
            
            Set lstitem = ListView2.ListItems.Add(, , Format$(rsac7.Fields!transdate, "mmm dd yyyy"))
            
                lstitem.SubItems(3) = Format$(rsac7.Fields!txtotal, "#,##0.00")
                lstitem.SubItems(1) = rsac7.Fields!transtime
                lstitem.SubItems(2) = rsac7.Fields!lname & " , " & rsac7.Fields!fname
               
            rsac7.MoveNext
            
            Loop
    Else
    MsgBox "No record found.", vbExclamation
    End If
    
    lbltotalsales.Caption = Format$(totalsale, "#,##0.00")
    rsac7.Close
    Set rsac7 = Nothing
End Sub

 
Private Sub Command4_Click()
 
 Dim oExcel As Object
 Dim oBook As Object
 Dim oSheet As Object
    
Dim sqlstr As String
sqlstr = "select tx.*, txtotal as t  , E.Emplname as Lname, E.empfname as fname from tbl_transaction as tx INNER JOIN vwEmployeeMaster as E ON tx.idno = E.empno where tx.recorddate between  #" & DateValue(DTPicker1.Value) & "#  and  #" & DateValue(DTPicker2.Value) & "#   order by transno Desc"
rsac7.Open sqlstr, acaccess1
    
        If Not rsac7.EOF Then
           
        
    
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set oSheet = oBook.Worksheets(1)
   oSheet.Range("A1:D1").Font.Bold = True
   oSheet.Range("A1:D1").Value = Array("Date", "Time", "Name", "Tx Total")
           
    Dim Y As Integer, X As Integer
    X = 2
    Y = 2
    Dim subt As Double
    
     Do While Not rsac7.EOF
     
        subt = rsac7.Fields!txtotal
        
        oSheet.Range("A" & Y & ":D" & Y).Value = Array(Format$(rsac7.Fields!transdate, "mmm dd yyyy"), CStr(rsac7.Fields!transtime), rsac7.Fields!lname & " , " & rsac7.Fields!fname, subt)
        
        Y = Y + 1
        
        rsac7.MoveNext
        
      Loop
        
  Dim saveasstring As String
  
   saveasstring = "C:\SALES-From" & Format$(DTPicker1.Value, "MMMDDYY") & "To" & Format$(DTPicker2.Value, "MMMDDYY") & ".xls"
    
   oBook.SaveAs saveasstring
   oExcel.Quit
   
   rsac7.MoveFirst
   rsac7.Close
   
     MsgBox "Report File on " + saveasstring, 64, "Info"
     
    Else
    MsgBox "No record found.", vbExclamation
    End If
End Sub
Private Sub Form_Load()
 
 DTPicker1.Value = Date
 DTPicker2.Value = Date
 
  Dim c As Integer
  c = ListView2.Width
  ListView2.ColumnHeaders.Item(1).Width = c * 0.1
  ListView2.ColumnHeaders.Item(2).Width = c * 0.1
  ListView2.ColumnHeaders.Item(3).Width = c * 0.4
  ListView2.ColumnHeaders.Item(4).Width = c * 0.3
   
  
  connectData
   
  
 
End Sub
Private Sub connectData()

 Set acaccess = New ADODB.Connection
 Set acdet = New ADODB.Connection
 Set acaccess1 = New ADODB.Connection
 
 
  If accSer.State = adStateOpen Then
            accSer.Close
  End If
  If acaccess.State = adStateOpen Then
            acaccess.Close
  End If
  If acaccess1.State = adStateOpen Then
            acaccess1.Close
  End If
     
    If acaccess.State = 1 Then acaccess.eClose
        If acaccess1.State = 1 Then acaccess1.eClose
          If acdet.State = 1 Then acdet.eClose
          
             acaccess.Open "DSN=canteen_offline"   ' MSACCESS
             acaccess1.Open "DSN=canteen_offline"
         
         
         
    dec

    Set rsaccess = New ADODB.Recordset
    Set rsaccDet = New ADODB.Recordset
    Set rsup = New ADODB.Recordset
    Set rbu = New ADODB.Recordset
    Set rtx = New ADODB.Recordset
End Sub
