VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransactionRecords 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Records"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14790
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   8535
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14535
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transaction Details"
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
         Height          =   4095
         Left            =   120
         TabIndex        =   9
         Top             =   4320
         Width           =   14295
         Begin VB.CommandButton Command5 
            Height          =   615
            Left            =   7800
            TabIndex        =   19
            Top             =   4080
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Height          =   615
            Left            =   6000
            TabIndex        =   18
            Top             =   4080
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cancel Transaction"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   11160
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3120
            Width           =   3015
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2535
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   4471
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16761024
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Rec  #"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item No."
               Object.Width           =   1305
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Itemcode"
               Object.Width           =   1305
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item Name"
               Object.Width           =   6068
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Unit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Price"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Subtotal"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label9 
            Caption         =   "0"
            Height          =   255
            Left            =   3240
            TabIndex        =   20
            Top             =   3840
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Top             =   4440
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Grand Total:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   4440
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1440
            TabIndex        =   13
            Top             =   4080
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Subsidy :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   4080
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Top             =   3720
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Subtotal :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   3720
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.TextBox txtsearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.ComboBox cbosearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmTransactionRecords.frx":0000
         Left            =   1200
         List            =   "frmTransactionRecords.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   278
         Width           =   3615
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3375
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16761024
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Idno"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   -2147483645
         CalendarTitleForeColor=   16777215
         Format          =   138805249
         CurrentDate     =   41655
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   7080
         TabIndex        =   17
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   -2147483645
         CalendarTitleForeColor=   16777215
         Format          =   138805249
         CurrentDate     =   41655
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "to :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6600
         TabIndex        =   21
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   278
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmTransactionRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txtotal As String
Dim transcode As String


Private Sub cbosearch_Click()
    clearlist
    If cbosearch.ListIndex = 0 Then
        
        txtsearch.Visible = False
        DTPicker1.Visible = True
        DTPicker2.Visible = True
        
   ElseIf cbosearch.ListIndex = 1 Then
   
        
        txtsearch.Visible = True
        
        DTPicker1.Visible = False
        DTPicker2.Visible = False
        
        
   End If
End Sub

Private Sub Command1_Click()
    clearlist
    rsfilter
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
If Not ListView1.ListItems.Count = 0 Then

If Label9.Caption = "0" Then
    
    frmAccesscode.Show 1
    
Else

Dim rsvoid As New ADODB.Recordset
Dim rstrans As New ADODB.Recordset
Dim voidtotal As Double
voidtotal = 0
    
    i = 1
            Do While i <= ListView1.ListItems.Count
            
            ListView1.ListItems.Item(i).Selected = True
            
            If ListView1.SelectedItem.Checked = True Then
                'MsgBox ListView1.SelectedItem.Text
                rsvoid.Open "select * from tbl_transdetails where  recno = " & ListView1.SelectedItem.Text & "", ac, adOpenKeyset, adLockOptimistic, adCmdText
                'rs3.EditMode
                rsvoid.MoveFirst
                voidtotal = voidtotal + Val(rsvoid.Fields!subtotal)
                rsvoid.Fields!status = "VOID"
                rsvoid.Update
                rsvoid.Close
            Else
                    'do nothing
                    
            End If
                i = i + 1
            Set rs3 = Nothing
            Loop

If voidtotal > 0 Then
        rstrans.Open "select * from tbl_transaction where  transno =  " & transcode & "", ac, adOpenKeyset, adLockOptimistic, adCmdText
        rstrans.MoveFirst
        rstrans.Fields!txtotal = Val(txtotal) - Val(voidtotal)
        rstrans.Update
        rstrans.Close

End If


MsgBox "Item(s) Selected are now Void!" & "Now Update txtable with code:" & transcode & " and voidtotal =" & voidtotal & " TOTAL: " & txtotal
clearlist
rsfilter



End If


Else
MsgBox "No record to cancel.", vbCritical
txtsearch.SetFocus
Exit Sub
End If

End Sub

Private Sub Form_Activate()

'txtsearch.SetFocus

Dim w1 As Integer
Dim w2 As Integer

w1 = ListView1.Width
w2 = ListView2.Width


ListView2.ColumnHeaders(1).Width = 0.07 * w2
ListView2.ColumnHeaders(2).Width = 0.12 * w2
ListView2.ColumnHeaders(3).Width = 0.15 * w2
ListView2.ColumnHeaders(4).Width = 0.12 * w2
ListView2.ColumnHeaders(5).Width = 0.4 * w2
ListView2.ColumnHeaders(6).Width = 0.12 * w2




ListView1.ColumnHeaders(1).Width = 0.1 * w1
ListView1.ColumnHeaders(2).Width = 0.1 * w1
ListView1.ColumnHeaders(3).Width = 0.1 * w1
ListView1.ColumnHeaders(4).Width = 0.27 * w1




DTPicker1.Value = DateValue(Date)
DTPicker2.Value = DateValue(Date)



End Sub

Private Sub rsfilter()
 
 'On Error GoTo e:
 
 
    Dim sby As String
    
    sby = "idno"
     
     
     If rsAC.State = adStateOpen Then
            rsAC.Close
     End If
     
    If cbosearch.ListIndex = 0 Then
    
    rsAC.Open "SELECT tx.transno, tx.transdate, tx.transtime , tx.idno , tx.txtotal as total , E.Emplname as Lname, E.empfname  as fname FROM tbl_transaction as tx INNER JOIN vwEmployeeMaster as E ON tx.idno = E.empno where tx.transdate between #" & DateValue(DTPicker1.Value) & "# AND #" & DateValue(DTPicker2.Value) & "# AND  tx.status = 1  order by transno DESC", ac
    
    ElseIf cbosearch.ListIndex = 1 Then
    
    rsAC.Open "SELECT tx.transno, tx.transdate, tx.transtime , tx.idno , tx.txtotal as total , E.Emplname as Lname, E.empfname  as fname FROM tbl_transaction as tx INNER JOIN vwEmployeeMaster as E ON tx.idno = E.empno where tx." & sby & " like '" & txtsearch.Text & "%' OR E.EmpLname LIKE '" & txtsearch.Text & "%' AND  tx.status = 1  order by transno DESC", ac
    
    
    End If
    
    
    
    
    If Not rsAC.EOF Then
        fill_lst
    Else
       MsgBox "No record found.", vbExclamation
    End If
    
    rsAC.Close
    Set rsAC = Nothing
 
 Exit Sub
 
e:
 
 MsgBox err.Description, vbOKOnly + vbCritical, "Error1!"
 


End Sub

Private Sub fill_lst()
Dim x As Integer
            ListView2.ListItems.Clear
            
            
            Do While Not rsAC.EOF
            Set lstitem = ListView2.ListItems.Add(, , rsAC.Fields!transno)
            lstitem.SubItems(1) = rsAC.Fields!transdate
            lstitem.SubItems(2) = rsAC.Fields!transtime
            lstitem.SubItems(3) = rsAC.Fields!idno
            lstitem.SubItems(4) = rsAC.Fields!lname + " , " + rsAC.Fields!fname
            lstitem.SubItems(5) = FormatNumber(rsAC.Fields!total)
          Set rs1 = Nothing
            
            
            
            rsAC.MoveNext
            
            Loop
End Sub



Private Sub Form_Load()
connectData

Label9.Caption = "0"
    
  cbosearch.ListIndex = 1
  
  txtsearch.Visible = True
  'txtsearch.SetFocus

End Sub
 

Private Sub ListView2_Click()

Dim x_subtotal As Currency
Dim x_subsidy As Currency
Dim x_grandtotal As Currency

x_subtotal = 0
transcode = ListView2.SelectedItem.Text
If Not ListView2.ListItems.Count = 0 Then
Set rsAC = Nothing
ListView1.ListItems.Clear


rsAC.Open "select * from tbl_transdetails where transno = '" & ListView2.SelectedItem & "' and status = 'OK'", ac, adOpenStatic, adLockOptimistic

Do While Not rsAC.EOF
                    Set lstitem = ListView1.ListItems.Add(, , rsAC.Fields!recno)
                             lstitem.SubItems(1) = rsAC.Fields!itemno
                              lstitem.SubItems(2) = rsAC.Fields!itemcode
                           
                           Set rs1 = Nothing
                           rs1.Open "select * from tbl_inventory where recno = " & rsAC.Fields!recno & "", ac
                           
                           If Not rs1.EOF Then
                           lstitem.SubItems(3) = rs1("itemname")
                          
                           End If
                           
                           Set rs1 = Nothing
                             lstitem.SubItems(4) = rsAC.Fields!qty
                             lstitem.SubItems(5) = rsAC.Fields!unitcode
                             lstitem.SubItems(6) = FormatNumber(rsAC.Fields!Price)
                             lstitem.SubItems(7) = FormatNumber(rsAC.Fields!subtotal)
                            
                            'lstitem.SubItems(8) = rsAC.Fields!recno

                             x_subtotal = x_subtotal + rsAC("subtotal")
                       
                    rsAC.MoveNext
                    
Loop


'Label4.Caption = ListView2.SelectedItem.SubItems(4)

Label2.Caption = FormatNumber(x_subtotal)
txtotal = x_subtotal
x_subsidy = Label4.Caption
x_grandtotal = x_subtotal - x_subsidy
Label8.Caption = FormatNumber(x_grandtotal)



'txtsearch.SetFocus
Else
MsgBox "No record to select.", vbCritical
txtsearch.SetFocus
Exit Sub
End If
End Sub


Private Sub clearlist()
ListView1.ListItems.Clear
ListView2.ListItems.Clear
Label4.Caption = "0.00"
Label2.Caption = "0.00"
Label8.Caption = "0.00"
Set rsAC = Nothing

End Sub

Private Sub txtsearch_Change()
clearlist

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
