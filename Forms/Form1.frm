VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCredit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PKI - Canteen Charging System"
   ClientHeight    =   10365
   ClientLeft      =   -4440
   ClientTop       =   -1500
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10245
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   15015
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   440
         Left            =   5400
         Picture         =   "Form1.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   360
         Width           =   440
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&New Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12480
         TabIndex        =   27
         Top             =   8760
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search &Employee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12480
         TabIndex        =   26
         Top             =   8040
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12480
         TabIndex        =   25
         Top             =   7200
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   3375
         Left            =   120
         TabIndex        =   7
         Top             =   5880
         Width           =   8775
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   5400
            Top             =   2880
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Photo"
            Height          =   2535
            Left            =   5280
            TabIndex        =   16
            Top             =   720
            Width           =   3255
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               Height          =   2130
               Left            =   480
               Picture         =   "Form1.frx":0E8E
               Stretch         =   -1  'True
               Top             =   240
               Width           =   2220
            End
         End
         Begin zkemkeeperCtl.CZKEM CZKEM1 
            Height          =   495
            Left            =   5520
            OleObjectBlob   =   "Form1.frx":28CC
            TabIndex        =   29
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lbl_asof 
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   32
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "OT - Subsidy :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7200
            TabIndex        =   28
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Subsidy :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7320
            TabIndex        =   23
            Top             =   1680
            Visible         =   0   'False
            Width           =   1335
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
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   1680
            TabIndex        =   22
            Top             =   2400
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Payables :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   21
            Top             =   2400
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Charge Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label13 
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
            Height          =   345
            Left            =   1680
            TabIndex        =   13
            Top             =   1920
            Width           =   3135
         End
         Begin VB.Label Label12 
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
            Height          =   330
            Left            =   1680
            TabIndex        =   12
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label Label11 
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
            Height          =   345
            Left            =   1680
            TabIndex        =   11
            Top             =   960
            Width           =   3135
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "RFID No. :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   10
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   9
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Emp. No. :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label15 
            BackColor       =   &H00004000&
            Height          =   495
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   8745
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   11640
         Top             =   360
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4815
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8493
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16761024
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ITEMNAME"
            Object.Width           =   11995
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "QTY"
            Object.Width           =   1341
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UNIT"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PRICE"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "SUBTOTAL"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "RECNO"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   11280
         Top             =   360
      End
      Begin VB.CommandButton tap 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&T"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         HelpContextID   =   1
         Left            =   14760
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6720
         Width           =   135
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   11880
         TabIndex        =   18
         Top             =   6120
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   14520
         TabIndex        =   30
         Top             =   600
         Width           =   45
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   14520
         TabIndex        =   6
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label26 
         Caption         =   "180"
         Height          =   255
         Left            =   9720
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   10080
         TabIndex        =   20
         Top             =   10440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10440
         TabIndex        =   19
         Top             =   9960
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         TabIndex        =   17
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "» OFFLINE «"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   13320
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CREDIT CHARGES"
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
      TabIndex        =   1
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "FrmCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xprice As Currency
Dim xqty As Currency
Dim xsubtotal As Currency
Dim xtotal As Currency
Dim vMachinenumber As String

Private Sub lv_recompute()
On Error GoTo err:

       Dim i As Integer
       
       ListView1.Refresh
        
       Dim xx_subtotal As Currency
       xx_subtotal = 0
       i = 1
       
  Do While i <= ListView1.ListItems.Count
  
        ListView1.ListItems.Item(i).Selected = True
        
        xx_subtotal = xx_subtotal + ListView1.SelectedItem.SubItems(5)
        
        ListView1.SelectedItem.Text = i
        i = i + 1
        
    Set rs3 = Nothing
  Loop
       Dim x_subsidy As Currency
       Dim x_grand As Currency
                                   
        x_grand = xx_subtotal - x_subsidy
        x_subsidy = Label21.Caption
                            
   Label20.Caption = FormatNumber(xx_subtotal)
   Label17.Caption = FormatNumber(x_grand)
Exit Sub

err:
  MsgBox "Error!" + err.Description, vbOKOnly + vbCritical, "Error!"
  
End Sub

Private Sub Command1_Click()
On Error GoTo err

Timer3.Enabled = False
 
FrmSearchEmployee.Show 1
Text1.SetFocus

Exit Sub

err:
  
  MsgBox "Error!" + err.Description, vbOKOnly + vbCritical, "Error!"
  

End Sub
 
 
Private Sub Command2_Click()
On Error GoTo e:

Dim i As Integer
Dim x_grand As Currency
Dim x_total As Currency

Me.Caption = "PKI - Canteen Charging System" & ""
 
If Not ListView1.ListItems.Count = 0 And Not Label11.Caption = "-" Then

        Set rs2 = Nothing
        rs2.Open "select * from tbl_transaction", ac, adOpenDynamic, adLockOptimistic
        rs2.AddNew
        rs2.Fields!transdate = Date
        rs2.Fields!transtime = Time
        rs2.Fields!incharge = FrmMainMenu.StatusBar1.Panels(2).Text
        rs2.Fields!cardno = Label13.Caption
        rs2.Fields!idno = Label11.Caption
        rs2.Fields!status = "1"
        rs2.Fields!remarks = "-"
        rs2.Fields!subsidy = FormatNumber(Label21.Caption, 2)
        rs2.Fields!txtotal = FormatNumber(Label17.Caption, 4)
       
        rs2.Update
        rs2.Close
        
        i = 1
  
  Do While i <= ListView1.ListItems.Count
  
  
        
        ListView1.ListItems.Item(i).Selected = True
        
        Set rs3 = Nothing
        
        rs3.Open "select * from tbl_transdetails", ac, adOpenDynamic, adLockOptimistic
        rs3.AddNew
        
        Set RS = Nothing
        
        RS.Open " select max(transno) as xtransno from tbl_transaction", ac, adOpenDynamic, adLockOptimistic
         
         rs3.Fields!transno = RS("xtransno")
         rs3.Fields!itemno = ListView1.SelectedItem
         rs3.Fields!itemcode = ListView1.SelectedItem.SubItems(6)
         rs3.Fields!qty = ListView1.SelectedItem.SubItems(2)
         rs3.Fields!unitcode = ListView1.SelectedItem.SubItems(3)
         rs3.Fields!Price = ListView1.SelectedItem.SubItems(4)
         rs3.Fields!subtotal = ListView1.SelectedItem.SubItems(5)
         rs3.Fields!status = "OK"
        
        
        rs3.Update
        rs3.Close
        
             
        i = i + 1
    Set rs3 = Nothing
  Loop
  

x_grand = 0
x_total = 0
MsgBox "Record successfully saved!", vbInformation
clearlist
Text1.SetFocus
Else
MsgBox "No record / employee to save.", vbCritical
Text1.SetFocus
Exit Sub
End If

Timer3.Enabled = False
Command2.Enabled = False

Exit Sub
e:
 
  MsgBox "Error!" + err.Description, vbOKOnly + vbCritical, "Error!"
  

End Sub
  

Private Sub Command4_Click()

Command2.Enabled = False
Timer3.Enabled = False

clearlist

End Sub

 

Private Sub Command5_Click()
frm_searchitem.Show 1

End Sub



Private Sub clearlist()
ListView1.ListItems.Clear
Image1.Picture = LoadPicture(App.Path & "\Photos\kao_Logo.jpg")
Label11.Caption = "-"
Label12.Caption = "-"
Label13.Caption = "-"
Label23.Caption = "0.00"
 
Label20.Caption = "0.00"
Label21.Caption = "0.00"
Label17.Caption = "0.00"
 


xprice = 0
xqty = 0
xsubtotal = 0
xtotal = 0

Text1.SetFocus



End Sub


Private Sub Form_activate()

Label6.Caption = FormatDateTime(Date, vbLongDate)
 
Text1.SetFocus

End Sub

Private Sub Form_Load()

On Error GoTo err

Dim lngComPort As Integer
Dim lngMachineNum As Integer
Dim lngBaudRate As Long
Dim bconn As Boolean
Dim devModel As String
Dim reader As Integer
Dim devNo As Integer
Dim ipAdd As String
Dim portNo As Long
Dim X As Long

 

Label26.Caption = "15"

dataconnect

Me.MousePointer = vbHourglass


            RS.Open "select * from tbl_deviceinfo", ac

            If Not RS.EOF Then


                devNo = RS.Fields!deviceno
                ipAdd = RS.Fields!ipaddress
                portNo = RS.Fields!portNo
    

            Else
                Set RS = Nothing
                MsgBox "No device detected.", vbInformation
            Exit Sub
            End If

 
bconn = CZKEM1.Connect_Net(Trim(ipAdd), portNo)

 

vMachinenumber = lngMachineNum
 

If bconn Then
 

        CZKEM1.BASE64 = 0
        CZKEM1.RegEvent 1, 32767
         

Else

    MsgBox "Can't connect to the specified device", vbCritical, "RFID Error"
    MsgBox "No device detected. Please contact MIS administrator for assistance.", vbCritical

End

End If
 
Me.MousePointer = vbDefault
Exit Sub

err:
MsgBox "Loading Form Error!" & err.Description, vbCritical + vbOKOnly, "ERROR!"


End Sub
 

Private Sub Label11_Change()
On Error GoTo error:

    
    If Not Label11.Caption = "-" Then
        Fetch_record
        Unload idplease
    End If

Exit Sub
   
error:

MsgBox "Oops!, SOMETHING WENT WRONG !!!" & vbNewLine & vbNewLine & err.Description, vbOKOnly, "ERROR"

End Sub
 
 
 

Private Sub Label13_Change()
 Unload idplease
End Sub
 

Private Sub Label20_Change()

    If Label21.Caption = "0.00" Then
        Label17.Caption = Label20.Caption
    Else
        Exit Sub
    End If
End Sub

 

Private Sub Label26_Change()

If Label26.Caption = 0 Then
    
    If Not ListView1.ListItems.Count = 0 And Label11.Caption <> "-" Then
            'do not close a record is pending to be saved
       Beep
       Label26.Caption = 60
        
       MsgBox "A Record must be saved!", vbCritical + vbOKOnly, "System Idle"
        
    
    Else
     
    Timer3.Enabled = False
    Label26.Caption = 180
    Timer2.Enabled = False
    
     
    
    End If
    
    'frmTimeout.Show 1
End If

End Sub

Private Sub ListView1_Click()
 Dim rcvqty As Variant
 Dim xy As Integer
Label26.Caption = 180
If Not ListView1.ListItems.Count = 0 Then
x_begin:

    rcvqty = InputBox("Please enter quantity of " & ListView1.SelectedItem.SubItems(1) + " CURRENT QTY :" + ListView1.SelectedItem.SubItems(2) & "", "Canteen Charging System", "" & ListView1.SelectedItem.SubItems(2) & "")
    If IsNumeric(rcvqty) Then
                    If Val(rcvqty) > 0 Then
                       xy = ListView1.SelectedItem.Index
                                    
                                    xqty = rcvqty
                                    xprice = ListView1.SelectedItem.SubItems(4)
                                    xsubtotal = FormatNumber(xqty * xprice)
                                               
                                    'a.SubItems(4) = FormatNumber(rs1("price"))
                                    ListView1.SelectedItem.SubItems(5) = FormatNumber(xsubtotal)
                                    ListView1.ListItems(xy).SubItems(2) = rcvqty
                                    lv_recompute
                                    
                                 Text1.SetFocus
                        Exit Sub
                    Else
                        MsgBox "Quantity entered is invalid.", vbCritical
                         Text1.SetFocus
                        Exit Sub
                    End If
                     Text1.SetFocus
                Else
                MsgBox "Please enter valid numeric value.", vbInformation
                GoTo x_begin
                End If



Else
MsgBox "No record selected.", vbCritical


Exit Sub
End If

Text1.SetFocus


End Sub



Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo err

Dim xy As Currency
Dim xsubtotal As Currency
Dim rcvqty As Variant
'xsubtotal = 0

If KeyAscii = 13 Then
'MsgBox ListView1.SelectedItem.Index

x_begin:

    rcvqty = InputBox("Please enter quantity if " & ListView1.SelectedItem.SubItems(1) + " CURRENT QTY : " + ListView1.SelectedItem.SubItems(2) & "", "Canteen Charging System", "" & ListView1.SelectedItem.SubItems(2) & "")
    If IsNumeric(rcvqty) Then
                    If Val(rcvqty) > 0 Then
                       xy = ListView1.SelectedItem.Index
                                    
                                    xqty = rcvqty
                                    xprice = ListView1.SelectedItem.SubItems(4)
                                    xsubtotal = FormatNumber(xqty * xprice)
                                    
                                                'a.SubItems(4) = FormatNumber(rs1("price"))
                                    ListView1.SelectedItem.SubItems(5) = FormatNumber(xsubtotal)
                                    ListView1.ListItems(xy).SubItems(2) = rcvqty
                                    lv_recompute
                                    Text1.SetFocus
                        Exit Sub
                    Else
                        MsgBox "Quantity entered is invalid.", vbCritical
                        Exit Sub
                    End If
                Else
                MsgBox "Please enter valid numeric value.", vbInformation
                GoTo x_begin
                End If

ElseIf KeyAscii = vbKeyDelete Or KeyAscii = 27 Then

   If vbYes = MsgBox("Are you sure to remove this Item", vbYesNo + vbCritical, "Please Confirm") Then
   
            ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
             
            lv_recompute
            Text1.SetFocus
            
        
   End If
   


End If

    Text1.SetFocus
     
    CZKEM1.EnableDevice vMachinenumber, False

Exit Sub
err:
Text1.SetFocus
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label26.Caption = 180
End Sub

Private Sub tap_Click()
 
CZKEM1.EnableDevice vMachinenumber, True
idplease.Show 1
 
End Sub

Private Sub Text1_Change()
 
 Timer2.Enabled = True
 
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Trim(Text1.Text) <> "" Then
        
        Refresh_lv1
        
    Else
        Text1.SetFocus
    End If
     
End If



End Sub



Private Sub Timer1_Timer()
 
    Label7.Caption = Time
    Me.MousePointer = vbDefault
   
   DoEvents
   DoEvents
 
End Sub

Private Sub Timer2_Timer()
   DoEvents
   DoEvents
    Me.MousePointer = vbDefault
   DoEvents
   DoEvents
End Sub


Private Sub CZKEM1_OnAttTransaction(ByVal EnrollNumber As Long, ByVal IsInValid As Long, ByVal AttState As Long, ByVal VerifyMethod As Long, ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long)
 
    Label11.Caption = EnrollNumber
    CZKEM1.EnableDevice vMachinenumber, False
    
End Sub

Private Sub CZKEM1_OnHIDNum(ByVal CardNumber As Long)
    
    Label13.Caption = CardNumber
    CZKEM1.EnableDevice vMachinenumber, False
    
End Sub

Private Sub Fetch_record()

 'On Error GoTo error:
On Error Resume Next

Dim xs As String
    
    xs = Label11.Caption
    Set rs1 = Nothing
    Set rs2 = Nothing

rs1.Open "SELECT emplname,empfname from VWemployeemaster where empno = '" & Label11.Caption & "'", ac, adOpenDynamic, adLockOptimistic
 
Dim xtotalpay As Currency
xtotalpay = 0
        
If Not rs1.EOF Then
            
            Label12.Caption = rs1.Fields!emplname + ", " + rs1.Fields!empfname
            Set rsAC = Nothing
            
         rsAC.Open "SELECT SUM(txtotal) as  totalPayables from tbl_transaction where idno = '" & Label11.Caption & "' and status = 1", ac, adOpenStatic, adLockReadOnly

Do While Not rsAC.EOF


    If (rsAC.Fields!TotalPayables = Null Or IsNull(rsAC.Fields!TotalPayables)) Then
    
        xtotalpay = 0

    Else

        xtotalpay = rsAC.Fields!TotalPayables

    End If
    rsAC.MoveNext
Loop

    Label23.Caption = FormatNumber$(xtotalpay)
     
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim xyz As String
        Dim picpath As String
        
        xyz = Dir(App.Path & "\Photos\" & xs & ".jpg")
        
        If Trim(xyz) <> "" Then
         
            picpath = App.Path & "\Photos\" & xs & ".jpg"
         
        Else
          
            picpath = App.Path & "\Photos\kao_logo.jpg"
          
        End If
 
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
  'picpath = App.Path & "\Photos\" & xs & ".jpg"
  
  Image1.Picture = LoadPicture(picpath)
  
  Timer3.Enabled = True
  
  'Command2.Enabled = True

Else

    Label12.Caption = ""
    
    MsgBox "Invalid IDNO. Please  Contact CCS Administrator.", vbCritical + vbOKOnly, "MIS ADMIN"
    
    Image1.Picture = LoadPicture(App.Path & "\Photos\kao_logo.jpg")
    
      
     
    
    Exit Sub
    
End If

Exit Sub

error:

  MsgBox err.Description, vbCritical + vbOKOnly, "Error!"
  

 
End Sub

Private Sub Refresh_lv1()

On Error Resume Next

Dim a As ListItem
 
    Me.Caption = "PKI - Canteen Charging System"
 
If ListView1.ListItems.Count = 0 Then

    xtotal = 0
    
End If

Dim recnumber As Integer

Set rs1 = Nothing
rs1.Open "select * from tbl_inventory where barcodeno = '" & Text1.Text & "'", ac, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
                  Set a = ListView1.ListItems.Add(, , ListView1.ListItems.Count + 1)
                  a.SubItems(1) = rs1("itemname")
                  Dim rcvqty As String
                  Dim x_qty As Long
                  x_qty = 1
x_begin:
             rcvqty = 1
                
             If IsNumeric(rcvqty) Then
                    If Val(rcvqty) > 0 Then
                         a.SubItems(2) = rcvqty
                   Else
                        MsgBox "Quantity entered is invalid.", vbCritical
                        Exit Sub
                    End If
                Else
                    MsgBox "Please enter valid numeric value.", vbInformation
                    GoTo x_begin
                End If
                                                
                a.SubItems(3) = rs1("unitcode")
                          
              xqty = rcvqty
              xprice = rs1("price")
              xsubtotal = FormatNumber(xqty * xprice)
              
                          a.SubItems(4) = FormatNumber(rs1("price"))
                          a.SubItems(5) = FormatNumber(xsubtotal)
                          a.SubItems(6) = rs1("recno")
        
             xtotal = xtotal + xsubtotal
             Dim x_subsidy As Currency
             Dim x_grand As Currency
             x_grand = 0
             
             x_subsidy = Label21.Caption
             x_grand = xtotal - x_subsidy
             Label20.Caption = FormatNumber(xtotal)
             
             Label17.Caption = FormatNumber(x_grand)
             
              lv_recompute
              Text1.Text = ""
              Text1.SetFocus
              
              Else
                MsgBox "Barcode no. dont exist.", vbCritical
                Text1.Text = ""
                Text1.SetFocus
                Exit Sub
              End If
              
                                    
End Sub
 
 
Private Sub Timer3_Timer()
 DoEvents
 DoEvents
 
Command2.Enabled = True

End Sub
