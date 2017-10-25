VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
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
      TabIndex        =   33
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
         Height          =   4935
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1320
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8705
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
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Payables"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Payment(s)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "To"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblsubtotal 
         Caption         =   "Label13"
         Height          =   375
         Left            =   6720
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblsubsidy 
         Caption         =   "Label13"
         Height          =   375
         Left            =   4800
         TabIndex        =   36
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction History"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1575
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
         Left            =   2280
         TabIndex        =   23
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off Charge"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   2055
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
         Charset         =   1
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
         Charset         =   1
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
      Begin MSComCtl2.DTPicker pickend 
         Height          =   375
         Left            =   3960
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   70582273
         CurrentDate     =   41754
      End
      Begin MSComCtl2.DTPicker pickstart 
         Height          =   375
         Left            =   1920
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   70582273
         CurrentDate     =   41754
      End
      Begin VB.CommandButton btnsubmit 
         Caption         =   "Submit &Deduction"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   14.25
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   4
         Top             =   5760
         Width           =   2295
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
         Left            =   3360
         TabIndex        =   35
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Balance"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   2640
         Width           =   3015
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
         TabIndex        =   32
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Last Cutoff to Present Balance"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   31
         Top             =   1920
         Width           =   3015
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
         TabIndex        =   30
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
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
         Left            =   3600
         TabIndex        =   29
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   27
         Top             =   4080
         Width           =   615
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
         Left            =   3240
         TabIndex        =   25
         Top             =   5880
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Payment"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   1
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
         Height          =   465
         Left            =   3360
         TabIndex        =   12
         Top             =   2880
         Width           =   3135
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
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
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
         Top             =   3480
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
            Charset         =   1
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

Dim tcode As String
Dim prevbal As Double

Dim prevsubsidy As Double
Dim prevsubtotal As Double




Private Sub Form_Load()
prevbal = 0
prevsubsidy = 0
prevsubtotal = 0
tcode = Val(Format(Now, "MMddyyhhmmss"))
Image1.Picture = LoadPicture(App.Path & "\Photos\kao_logo.jpg")
Dim w As Integer
Dim w2 As Integer
w = LvSearch.Width
w2 = lvpaymentlogs.Width
LvSearch.ColumnHeaders(1).Width = 0.25 * w
LvSearch.ColumnHeaders(2).Width = 0.365 * w
LvSearch.ColumnHeaders(3).Width = 0.35 * w
LvSearch.LabelEdit = lvwManual

lvpaymentlogs.ColumnHeaders(1).Width = 0.2 * w2
lvpaymentlogs.ColumnHeaders(2).Width = 0.2 * w2
lvpaymentlogs.ColumnHeaders(3).Width = 0.25 * w2
lvpaymentlogs.ColumnHeaders(4).Width = 0.32 * w2
lvpaymentlogs.LabelEdit = lvwManual
listsearch
End Sub

Private Sub listsearch()

Dim lstitem As ListItem
rsACSearch.Open "SELECT * FROM vwemployeemaster where emplname like '" & Trim$(txtsearch.Text) & "%' OR empno like '%" & Trim$(txtsearch.Text) & "%' order by empno asc ", ac, adOpenStatic
   
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
rslistlog.Open "SELECT isnull(payable,0) as payable,  isnull(balance,0) as balance, isnull(payments,0) as payments, cutoffstartdate, cutoffdate FROM tbl_logs WHERE void = 0 AND empid = " & selectedid & " ORDER BY recno DESC", ac, adOpenStatic
        If rslistlog.RecordCount >= 1 Then
        lvpaymentlogs.ListItems.Clear
        
        Do While Not rslistlog.EOF
          Dim bal As Double
             bal = Val(rslistlog.Fields("payable"))
             If bal < 0 Then
                bal = 0
             End If
             
             
            Set lstitem = lvpaymentlogs.ListItems.Add(, , Format(Val(rslistlog.Fields("balance")), "##0.00"))
                    lstitem.SubItems(1) = Format(bal, "##0.00")
                    lstitem.SubItems(2) = Format(Val(rslistlog.Fields("payments")), "##0.00")
                    lstitem.SubItems(3) = (rslistlog.Fields("cutoffstartdate")) & " - " & (rslistlog.Fields("cutoffdate"))
                   
                    
           rslistlog.MoveNext
            
            
        Loop
        
        
        End If
    Set rslistlog = Nothing
End Sub

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

'On Error Resume Next
rsprevbal.Open "SELECT TOP (1) balance,subtotal, subsidy, ot_subsidy, cutoffstartdate,cutoffdate  FROM  tbl_logs WHERE void= 0 AND  EmpId = " & Trim(id) & " order by recno DESC", ac, adOpenStatic, adLockReadOnly
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
        pickstart.Value = rsprevbal.Fields("cutoffstartdate")
        pickend.Value = rsprevbal.Fields("cutoffdate")
        prevsubtotal = Val(rsprevbal.Fields("subtotal"))
     End If
     
  subtotal = bal '+ Val(rsprevbal.Fields("subsidy")) + Val(rsprevbal.Fields("ot_subsidy"))
  
  lblcutoffbal.Caption = Format(Val(subtotal), "##0.00")
  lblsubsidy.Caption = Format(Val(prevsubsidy), "##0.00")
  lblsubtotal.Caption = Format(Val(prevsubtotal), "##0.00")
  
  prevbal = Format(Val(subtotal), "##0.00")
  
Else
 lblcutoffbal.Caption = "0.00"
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
    MsgBox "Please Provide Amount", vbCritical + vbOKOnly, "Oops!"
    txtamount.SetFocus
 Else
     If MsgBox("Are you sure about this ?", vbQuestion + vbYesNo, "Please Confirm") = vbYes Then
           
            Set rs2 = Nothing
            Dim dduct As Double
            Dim txcode As String
            Dim tsubtotal As String
            Dim tsubsidy As String
            Dim act_deduction As String
            Dim d_empbalance As String
            Dim prevsubtotal As String
            
            
            
            tsubtotal = Val(lblsubtotal.Caption)
            tsubsidy = Val(lblsubsidy.Caption)
            prevsubtotal = (Val(tsubtotal) + Val(tsubsidy)) - Val(lblcutoffbal.Caption)
            
            act_deduction = Val(txtamount.Text) - tsubsidy
            If act_deduction < 0 Then
                act_deduction = 0
            
            End If
            
            If act_deduction = 0 Then
            d_empbalance = tsubtotal
            Else
            act_deduction = Val(act_deduction) + Val(prevsubtotal)
            d_empbalance = Val(tsubtotal) - Val(act_deduction)
            
            End If
            txcode = get_prevtxnumber(selectedid)
            
           dduct = Val(lblcutoffbal.Caption) - Val(txtamount.Text)
           
           ac.Execute "UPDATE tbl_logs SET payments = '" & Val(txtamount.Text) & "'," & _
           "  balance = " & dduct & _
           " ,actualdeduction = " & Val(act_deduction) & _
           " ,empbalance = " & d_empbalance & _
           " WHERE transcode ='" & txcode & "' AND empId =" & selectedid
           
          
          
          
          '  rs2.Open "select * from tbl_logs WHERE transcode ='" & txcode & "'", ac, adOpenDynamic, adLockOptimistic
          '  rs2.Fields!payments = Val(txtamount.Text)
          '  rs2.Fields!balance = Val(lblcutoffbal.Caption) - Val(txtamount.Text)
          '  rs2.Fields!remarks = txtremarks.Text
                
          'rs2.Update
          'rs2.Close
            
            listlog
            get_prevbal (selectedid)
            lblempno_Change
    End If
    'txtamount.SetFocus
    LvSearch.SetFocus
    'SendKeys ("{home}+{end}")

End If


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
Else
cleardetails
End If

'    txtamount.SetFocus
End Sub
Private Sub LvSearch_KeyPress(KeyAscii As Integer)
If LvSearch.ListItems.Count > 0 Then
    If KeyAscii = 13 Then
        selectedid = LvSearch.SelectedItem.Text
        selectedname = LvSearch.SelectedItem.ListSubItems.Item(1) & ", " & LvSearch.SelectedItem.ListSubItems.Item(2)
        listlog
        showdetails
        txtamount.SetFocus
    End If
Else
cleardetails
End If
End Sub

 
Private Sub txtamount_Click()
SendKeys ("{home}+{end}")
End Sub

Private Sub txtamount_GotFocus()
SendKeys ("{home}+{end}")
End Sub
Private Sub txtamount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnsubmit_Click
End If

End Sub
 
Private Sub txtremarks_Click()
SendKeys ("{home}+{end}")
End Sub

Private Sub txtremarks_GotFocus()
SendKeys ("{home}+{end}")
End Sub

Private Sub txtsearch_Change()
listsearch
End Sub

Private Sub txtsearch_Click()
SendKeys ("{home}+{end}")

End Sub

Private Sub txtsearch_GotFocus()
SendKeys ("{home}+{end}")

End Sub
