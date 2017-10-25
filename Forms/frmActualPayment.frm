VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActualPayment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Canteen Charging System"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   15240
   Icon            =   "frmActualPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   15240
   Begin VB.TextBox txtremarks 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   10320
      MultiLine       =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "frmActualPayment.frx":058A
      Top             =   8160
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   120
      TabIndex        =   17
      Top             =   3080
      Width           =   8175
      Begin MSComctlLib.ListView lvpaymentlogs 
         Height          =   2535
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1320
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16761024
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tx No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvdetails 
         Height          =   2415
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3960
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16761024
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Subtotal"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase History"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label9 
         BackColor       =   &H00004000&
         Height          =   495
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   8160
      End
   End
   Begin VB.TextBox txtamount 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10320
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   7560
      Width           =   3615
   End
   Begin MSComctlLib.ListView LvSearch 
      Height          =   1815
      Left            =   5520
      TabIndex        =   1
      Top             =   960
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16761024
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EmpNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lastname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtsearch 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6540
      Left            =   8400
      TabIndex        =   6
      Top             =   3000
      Width           =   6735
      Begin VB.CommandButton btnsubmit 
         Caption         =   "Submit &Payment"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   5760
         Width           =   3495
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Photo"
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2775
         Begin VB.Image Image1 
            Height          =   1695
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Label lblcutoffbal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         TabIndex        =   31
         Top             =   3240
         Width           =   3615
      End
      Begin VB.Label lblrunningbal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         TabIndex        =   29
         Top             =   3240
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unpaid Total: "
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   3360
         Width           =   1305
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   25
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   23
         Top             =   5880
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Payment"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   7080
         X2              =   0
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lblempno 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3360
         TabIndex        =   14
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblempname 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3360
         TabIndex        =   13
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label lbltbal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   3120
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   11
         Top             =   1440
         Width           =   1740
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Details"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction History"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7080
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00004000&
         Height          =   495
         Left            =   -360
         TabIndex        =   8
         Top             =   180
         Width           =   7920
      End
      Begin VB.Label Label8 
         BackColor       =   &H00004000&
         Height          =   495
         Left            =   0
         TabIndex        =   16
         Top             =   3840
         Width           =   6735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   15015
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Search Employee:"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ACTUAL PAYMENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "frmActualPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim selectedid As String
Dim selectedname As String

Dim rslistlog As New ADODB.Recordset
Dim rsdetails As New ADODB.Recordset
Dim rsprevbal As New ADODB.Recordset
Dim rspay As New ADODB.Recordset


Dim tcode As String
Dim prevbal As Double

Dim prevsubsidy As Double
Dim prevsubtotal As Double

Private Sub enddate_Change()
    listlog
End Sub

Private Sub Form_Load()

connectData

'stdate.Value = DateAdd("d", -31, Now)
'enddate.Value = DateAdd("d", 2, Now)




prevbal = 0
prevsubsidy = 0
prevsubtotal = 0
tcode = Val(Format(Now, "MMddyyhhmmss"))

Image1.Picture = LoadPicture(App.Path & "\Photos\kao_logo.jpg")

Dim w As Integer
Dim w2 As Integer
Dim w3 As Integer


w = LvSearch.Width
w2 = lvpaymentlogs.Width
w3 = lvdetails.Width


LvSearch.ColumnHeaders(1).Width = 0.3 * w
LvSearch.ColumnHeaders(2).Width = 0.35 * w
LvSearch.ColumnHeaders(3).Width = 0.35 * w
 
LvSearch.LabelEdit = lvwManual

lvpaymentlogs.ColumnHeaders(1).Width = 0.25 * w2
lvpaymentlogs.ColumnHeaders(2).Width = 0.25 * w2
lvpaymentlogs.ColumnHeaders(3).Width = 0.25 * w2
lvpaymentlogs.ColumnHeaders(4).Width = 0.24 * w2
lvpaymentlogs.LabelEdit = lvwManual

lvdetails.ColumnHeaders(1).Width = 0.25 * w3
lvdetails.ColumnHeaders(2).Width = 0.25 * w3
lvdetails.ColumnHeaders(3).Width = 0.25 * w3
lvdetails.ColumnHeaders(4).Width = 0.24 * w3

lvdetails.LabelEdit = lvwManual


listsearch


End Sub

Private Sub listsearch()

Dim lstitem As ListItem

'rsACSearch.Open "SELECT * FROM vwemployeemaster WHERE MPayArea  = '99' AND emplname like '" & Trim$(txtsearch.Text) & "%' OR empno like '%" & Trim$(txtsearch.Text) & "%' order by empno asc ", ac, adOpenStatic

rsACSearch.Open "SELECT * FROM vwemployeemaster WHERE MPayArea  = '99'  AND (emplname like '" & Trim$(txtsearch.Text) & "%' OR empno like '%" & Trim$(txtsearch.Text) & "%')  order by empno asc ", ac, adOpenStatic
       
       If rsACSearch.RecordCount >= 1 Then
        LvSearch.ListItems.Clear
        
        If rsACSearch.RecordCount = 1 Then
            
            selectedid = rsACSearch.Fields!empno
            selectedname = rsACSearch.Fields!emplname & ", " & rsACSearch.Fields!empfname
            
            showdetails
            listlog
            
        Else
         cleardetails
        End If
                Do While Not rsACSearch.EOF
                    Set lstitem = LvSearch.ListItems.Add(, , rsACSearch.Fields!empno)
                             lstitem.SubItems(1) = rsACSearch.Fields!emplname
                             lstitem.SubItems(2) = rsACSearch.Fields!empfname
                    rsACSearch.MoveNext
                    
                Loop
        Else
                LvSearch.ListItems.Clear
        End If
        
        
    
'rsACSearch.Close
Set rsACSearch = Nothing
            

End Sub
Private Sub listlog()

Dim lstitem As ListItem

Dim showtotal As Double
        
showtotal = 0
lvpaymentlogs.ListItems.Clear
lvdetails.ListItems.Clear

'rslistlog.Open "SELECT isnull(payable,0) as payable,  isnull(balance,0) as balance, isnull(payments,0) as payments, cutoffstartdate, cutoffdate FROM tbl_logs WHERE void = 0 AND empid = " & selectedid & " ORDER BY recno DESC", ac, adOpenStatic
'rslistlog.Open "SELECT  charges, subsidy, ot_subsidy, other_subsidy, subsidystatus, subtotal, payrolldeduction,actualdeduction, empbalance, prev_bal ,cutoffstartdate, cutoffdate FROM tbl_logs WHERE void = 0 AND empid = " & selectedid & " AND void = 0  and MPAYAREA ='7M' AND PAYROLLDATE > '1/1/2015' ORDER BY recno DESC", ac, adOpenStatic

rslistlog.Open "SELECT * FROM tbl_transaction WHERE status =  1 and remarks ='-' AND idno = '" & selectedid & "' ORDER BY transno DESC ", ac, adOpenStatic
'AND transdate  between " & DateValue(stdate.Value) & "  AND " & DateValue(enddate.Value) & "
        
    If rslistlog.RecordCount >= 1 Then
        
        lvpaymentlogs.ListItems.Clear
        
        Dim tdate As Date
        Dim ttime As String
        Dim txtotal As Double
        
        
         
        Dim txno As Integer
    
        
        Do While Not rslistlog.EOF
        
        
        tdate = rslistlog.Fields("transdate")
        ttime = rslistlog.Fields("transtime")
        txtotal = rslistlog.Fields("txtotal")
                    
        txno = rslistlog.Fields("transno")
          
            Set lstitem = lvpaymentlogs.ListItems.Add(, , txno)
            
                    lstitem.SubItems(1) = tdate
                    lstitem.SubItems(2) = ttime
                    
                    lstitem.SubItems(3) = Format(txtotal, "#0.00")
                   
                   showtotal = showtotal + txtotal
                     
           rslistlog.MoveNext
             
        Loop
         
        End If
        
        
    
    
    rslistlog.Close
    
    Set rslistlog = Nothing
    
    lblcutoffbal.Caption = Format$(showtotal, "#0.00")
    txtamount.Text = Format$(showtotal, "#0.00")
    
    
    
End Sub

Private Sub lvpaymentlogs_Click()

Dim txno As Integer
Dim itemname As String

lvdetails.ListItems.Clear

If lvpaymentlogs.ListItems.Count > 0 Then
 
    txno = lvpaymentlogs.SelectedItem.Text
    Dim lsdetails As ListItem
        
    rsdetails.Open "SELECT * FROM tbl_transdetails WHERE transno ='" & txno & "' AND status = 'OK'", ac
                    
                    Do While Not rsdetails.EOF
                      
                        Set lsdetails = lvdetails.ListItems.Add(, , GetItemname(rsdetails.Fields("itemcode")))
                            
                            lsdetails.SubItems(2) = rsdetails.Fields("qty")
                            
                            lsdetails.SubItems(1) = Format$(rsdetails.Fields("price"), "#0.00")
                            
                            lsdetails.SubItems(3) = Format$(rsdetails.Fields("subtotal"), "#0.00")
                    
                    rsdetails.MoveNext
                    Loop
                    
    rsdetails.Close
                    
    End If

End Sub
Private Function GetItemname(str As String)
    Dim s As String
    
    s = ""
    
    rsprevbal.Open "SELECT TOP 1 itemname FROM tbl_inventory where recno =" & str & "", ac, adOpenDynamic

        Do While Not rsprevbal.EOF
                
            s = rsprevbal.Fields("itemname")
            
            rsprevbal.MoveNext
        Loop
    rsprevbal.Close
    
    Set rsprevbal = Nothing
   
        
        
    If s = "" Then
     GetItemname = "Unkown"
    Else
     GetItemname = s
    End If
    
End Function
Private Sub showdetails()

Dim picpath As String

On Error Resume Next

rsdetails.Open "SELECT * FROM vwEmployeeMaster WHERE empno = " & Trim(selectedid), ac, adOpenStatic, adLockReadOnly

If rsdetails.RecordCount >= 1 Then
 
        If Dir(App.Path & "\Photos\" & selectedid & ".jpg") <> "" Then
         
         picpath = App.Path & "\Photos\" & selectedid & ".jpg"
        
        Else
         
         picpath = App.Path & "\Photos\kao_logo.jpg"
        
        End If
        
        Image1.Picture = LoadPicture(picpath)
        
        lblempname.Caption = selectedname
        lblempno.Caption = selectedid
        
 Else
 cleardetails
End If


Set rsdetails = Nothing



End Sub
Private Sub cleardetails()

Image1.Picture = LoadPicture(App.Path & "\Photos\kao_logo.jpg")
lblempname.Caption = "-"
 lblempno.Caption = "-"
 lvpaymentlogs.ListItems.Clear
 lblcutoffbal.Caption = "0.00"
 selectedid = 0
 selectedname = ""
 lblrunningbal.Caption = "-"
 lbltbal.Caption = "-"
 txtamount.Text = "0.00"
 
End Sub
Private Function get_rbalance(ByVal id As String)
On Error Resume Next

Dim result As String
result = ""


rsprevbal.Open "SELECT sum(txtotal) as total FROM tbl_transaction WHERE status = 1 AND idno =" & id, ac, adOpenStatic, adLockOptimistic

If rsprevbal.RecordCount >= 1 Then
    
    lblrunningbal.Caption = Format(Val(rsprevbal.Fields("total")), "##0.00")
    get_rbalance = Format(Val(rsprevbal.Fields("total")), "##0.00")
    
Else
 lblrunningbal.Caption = "0.00"
End If
    
rsprevbal.Close

Set rsprevbal = Nothing
 
 
 
End Function
Private Sub get_prevbal(ByVal id As String)

On Error Resume Next

'rsprevbal.Open "SELECT TOP (1) *  FROM  tbl_logs WHERE void= 0 AND  EmpId = '" & Trim(id) & "' ORDER by recno DESC", ac, adOpenStatic, adLockReadOnly
rsprevbal.Open "SELECT TOP (1) *  FROM  tbl_logs WHERE   EmpId =  " & Trim(id) & "", ac, adOpenDynamic

If rsprevbal.RecordCount >= 1 Then
    Dim subtotal As Double
    Dim bal As Double
    
     If Val(rsprevbal.Fields("balance")) < 0 Then
         bal = 0
         prevsubsidy = 0
         prevsubtotal = 0
     Else
        bal = Val(rsprevbal.Fields("balance"))
        prevsubsidy = Val(rsprevbal.Fields("subsidy")) + Val(rsprevbal.Fields("ot_subsidy"))
         
        prevsubtotal = Val(rsprevbal.Fields("subtotal"))
        
     End If
     
    subtotal = bal '+ Val(rsprevbal.Fields("subsidy")) + Val(rsprevbal.Fields("ot_subsidy"))
  
    'lblcutoffbal.Caption = Format(Val(subtotal), "##0.00")
    
    prevbal = Format(Val(subtotal), "##0.00")
  
Else
 'lblcutoffbal.Caption = "0.00"
End If

rsprevbal.Close

Set rsprevbal = Nothing

End Sub

Private Function get_prevtxnumber(ByVal id As String)
Dim result As String
'On Error Resume Next
rsprevbal.Open "SELECT TOP (1) transcode, balance, subsidy, ot_subsidy  FROM  tbl_logs WHERE void= 0 AND  EmpId = " & Trim(id) & " order by recno DESC", ac, adOpenStatic, adLockReadOnly
If rsprevbal.RecordCount >= 1 Then
  result = rsprevbal.Fields("transcode")
Else
  result = False
End If

rsprevbal.Close
get_prevtxnumber = result


End Function

 

'''''''''''''''''''''''''''''''''''''''Events ''''''''''''''''''''''''''
Private Sub btnsubmit_Click()
 
 If Val(selectedid) = 0 Then
    
    MsgBox "Please Select an Employee", vbCritical + vbOKOnly, "Oops!"
    
 ElseIf Val(txtamount.Text) = 0 Then
    
    MsgBox "All are paid", vbCritical + vbOKOnly, "Oops!"
    txtamount.SetFocus
    
 Else
     
     If MsgBox("Are you sure about this ?", vbQuestion + vbYesNo, "Please Confirm") = vbYes Then
             
             
              MsgBox "Doing some ...."
              
              Dim sql As String
              sql = "UPDATE tbl_transaction SET status = 2 WHERE status = 1 AND idno ='" & selectedid & "'"
              acaccess.Execute sql
              
             MsgBox "PAID COMPLETELY"
              
             Set rs2 = Nothing
             rs2.Open "select * from tbl_transaction", ac, adOpenDynamic, adLockOptimistic
             rs2.AddNew
             rs2.Fields!transdate = Date
             rs2.Fields!transtime = Time
             rs2.Fields!incharge = FrmMainMenu.StatusBar1.Panels(2).Text
             rs2.Fields!cardno = "-"
             rs2.Fields!idno = selectedid
             rs2.Fields!status = "1"
             rs2.Fields!remarks = txtremarks.Text
             rs2.Fields!subsidy = 0
             rs2.Fields!txtotal = (FormatNumber(txtamount.Text, 4)) * -1
            
             rs2.Update
             rs2.Close
             
Set rs2 = Nothing

rs2.Open "select * from tbl_payments", ac, adOpenDynamic, adLockOptimistic
rs2.AddNew

rs2.Fields!EmpID = selectedid
rs2.Fields!amount = (FormatNumber(txtamount.Text, 4))
rs2.Fields!remarks = FrmMainMenu.StatusBar1.Panels(2).Text
rs2.Fields!cashier = FrmMainMenu.StatusBar1.Panels(2).Text
rs2.Fields!payment_date = Date
rs2.Fields!rec_date = Date
rs2.Fields!status = "1"
rs2.Fields!remarks = txtremarks.Text
 
rs2.Update
 
             
             
     End If
    
    LvSearch.SetFocus
    

End If


End Sub

Private Sub lblcutoffbal_Click()
 lblrunningbal.Caption = lblcutoffbal.Caption
End Sub

Private Sub lblempno_Change()
Dim bal As String
get_prevbal (selectedid)

lblrunningbal.Caption = get_rbalance(selectedid)
bal = get_rbalance(selectedid) + prevbal
lbltbal.Caption = Format(bal, "#0.00")


End Sub
 



Private Sub LvSearch_Click()

If LvSearch.ListItems.Count > 0 Then

    selectedid = LvSearch.SelectedItem.Text
    selectedname = LvSearch.SelectedItem.ListSubItems.Item(1) & ", " & LvSearch.SelectedItem.ListSubItems.Item(2)
    
    listlog
    
    showdetails
    
    get_rbalance (selectedid)
    
Else

    cleardetails

End If
End Sub

Private Sub LvSearch_GotFocus()

If LvSearch.ListItems.Count > 0 Then

    selectedid = LvSearch.SelectedItem.Text
    selectedname = LvSearch.SelectedItem.ListSubItems.Item(1) & ", " & LvSearch.SelectedItem.ListSubItems.Item(2)
    
    listlog
    showdetails
    get_rbalance (selectedid)
    
Else
    
    cleardetails
    
End If

' txtamount.SetFocus
 
End Sub
Private Sub LvSearch_KeyPress(KeyAscii As Integer)

If LvSearch.ListItems.Count > 0 Then

    If KeyAscii = 13 Then
        
        selectedid = LvSearch.SelectedItem.Text
        selectedname = LvSearch.SelectedItem.ListSubItems.Item(1) & ", " & LvSearch.SelectedItem.ListSubItems.Item(2)
        
        listlog
        
        showdetails
        
        txtamount.SetFocus
        get_rbalance (selectedid)
        
    End If
    
Else
    
    cleardetails
    
End If
End Sub

 
 

Private Sub stdate_Change()
listlog
End Sub

Private Sub txtamount_Click()
   
   txtsearch.SetFocus
   
 
    txtamount.Text = lblcutoffbal.Caption
    'SendKeys ("{home}+{end}")

End Sub

Private Sub txtamount_GotFocus()

txtamount.Text = lblcutoffbal.Caption
txtsearch.SetFocus
'SendKeys ("{home}+{end}")
End Sub
Private Sub txtamount_KeyPress(KeyAscii As Integer)
txtamount.Text = lblcutoffbal.Caption
txtsearch.SetFocus
If KeyAscii = 13 Then
 btnsubmit_Click
End If


End Sub
 
Private Sub txtremarks_Click()
'SendKeys ("{home}+{end}")
End Sub

Private Sub txtremarks_GotFocus()
'SendKeys ("{home}+{end}")
End Sub

Private Sub txtsearch_Change()
    listsearch
End Sub

Private Sub txtsearch_Click()
'SendKeys ("{home}+{end}")

End Sub

Private Sub txtsearch_GotFocus()
'SendKeys ("{home}+{end}")

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
         
         accSer.Open "DSN=update_ccs;UID=pip_connect;PWD=pipconnect"   ' SQL
         acaccess.Open "DSN=canteen_offline"   ' MSACCESS
         acaccess1.Open "DSN=canteen_offline"
         'acdet.Open "DSN=canteen_offline"   ' MSACCESS
         
    dec

    Set rsaccess = New ADODB.Recordset
    Set rsaccDet = New ADODB.Recordset
    Set rsup = New ADODB.Recordset
    Set rbu = New ADODB.Recordset
    Set rtx = New ADODB.Recordset
End Sub
