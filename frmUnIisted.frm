VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUnlisted 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload Unlisted Transactions"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   11505
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   4320
      Width           =   4455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   139001857
      CurrentDate     =   42627
   End
   Begin VB.Frame Frame1 
      Caption         =   "EMPLOYEE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2895
      Begin VB.TextBox EID 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   50
         TabIndex        =   9
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox EName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   50
         TabIndex        =   8
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   240
         Picture         =   "frmUnIisted.frx":0000
         Stretch         =   -1  'True
         Top             =   260
         Width           =   2220
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ADD ITEM"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox grndtotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   8760
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SAVE"
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&SELECT"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Unit Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Subtotal"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "NOTES / REASONS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   12
      Top             =   4080
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATE LISTED:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   1290
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7695
      TabIndex        =   4
      Top             =   4200
      Width           =   960
   End
End
Attribute VB_Name = "frmUnlisted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmEmployee.Show 1

End Sub

Private Sub Command2_Click()

'On Error GoTo e:

Dim i As Integer
Dim x_grand As Currency
Dim x_total As Currency

If Trim(Text1.Text) <> "" Then

If Not ListView1.ListItems.Count = 0 And Not Trim(EID.Text) = "" Then

        Set rs2 = Nothing
        rs2.Open "select * from tbl_transaction", ac, adOpenDynamic, adLockOptimistic
        rs2.AddNew
        rs2.Fields!transdate = Date
        rs2.Fields!transtime = Time
        rs2.Fields!incharge = FrmMainMenu.StatusBar1.Panels(2).Text
        rs2.Fields!cardno = "-"
        rs2.Fields!idno = EID.Text
        rs2.Fields!status = "1"
        rs2.Fields!remarks = Text1.Text & " Dated:" & DateValue(DTPicker1.Value)
        
        rs2.Fields!subsidy = 0
        rs2.Fields!txtotal = grndtotal.Text
       
        rs2.Update
        rs2.Close
        
        i = 1
        
  
  Do While i <= ListView1.ListItems.Count
  
  
        
        ListView1.ListItems.Item(i).Selected = True
        
        Set rs3 = Nothing
        
        rs3.Open "select * from tbl_transdetails", ac, adOpenDynamic, adLockOptimistic
        rs3.AddNew
        
        Set RS = Nothing
        
        RS.Open " SELECT max(transno) as xtransno from tbl_transaction", ac, adOpenDynamic, adLockOptimistic
         
       
         rs3.Fields!transno = RS("xtransno")
         rs3.Fields!itemno = ListView1.SelectedItem
         rs3.Fields!itemcode = ListView1.SelectedItem.SubItems(1)
         rs3.Fields!qty = ListView1.SelectedItem.SubItems(4)
         rs3.Fields!unitcode = ListView1.SelectedItem.SubItems(5)
         rs3.Fields!Price = ListView1.SelectedItem.SubItems(3)
         rs3.Fields!subtotal = ListView1.SelectedItem.SubItems(6)
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
'Text1.SetFocus
Else
MsgBox "No record / employee to save.", vbCritical
Text1.SetFocus
Exit Sub
End If

Else
    MsgBox "Please provide reasons for this transaction", vbCritical, "Oops!"
    
End If
 
Exit Sub
e:
 
  MsgBox "Error!" + err.Description, vbOKOnly + vbCritical, "Error!"
End Sub

Private Sub Command3_Click()
 clearlist
 Text1.Text = ""
 
End Sub

Private Sub Command4_Click()

frmadditem.Show 1

End Sub
  
 
Private Sub lv_recompute()
On Error GoTo err:

       Dim i As Integer
       
       ListView1.Refresh
        
       Dim xx_subtotal As Currency
       xx_subtotal = 0
       i = 1
       
  Do While i <= ListView1.ListItems.Count
  
        ListView1.ListItems.Item(i).Selected = True
        
        xx_subtotal = xx_subtotal + ListView1.SelectedItem.SubItems(6)
        
        ListView1.SelectedItem.Text = i
        i = i + 1
         
  Loop
       
       Dim x_grand As Currency
                                   
        x_grand = xx_subtotal
       
    grndtotal.Text = FormatNumber(x_grand)
Exit Sub

err:
  MsgBox "Error!" + err.Description, vbOKOnly + vbCritical, "Error!"
End Sub

Private Sub EID_Change()

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim xyz As String
        Dim picpath As String
        
        xyz = Dir(App.Path & "\Photos\" & Trim(EID.Text) & ".jpg")
        
        If Trim(xyz) <> "" Then
         
            picpath = App.Path & "\Photos\" & Trim(EID.Text) & ".jpg"
         
        Else
          
            picpath = App.Path & "\Photos\kao_logo.jpg"
          
        End If
 
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
  'picpath = App.Path & "\Photos\" & xs & ".jpg"
  
  Image1.Picture = LoadPicture(picpath)
End Sub

Private Sub EID_GotFocus()
Text1.SetFocus

End Sub

Private Sub EName_Change()
'Text1.SetFocus
End Sub

Private Sub Form_Activate()
Dim w As Double
w = ListView1.Width
ListView1.ColumnHeaders(1).Width = 0.1 * w
ListView1.ColumnHeaders(2).Width = 0.15 * w
ListView1.ColumnHeaders(3).Width = 0.25 * w
ListView1.ColumnHeaders(4).Width = 0.15 * w
ListView1.ColumnHeaders(5).Width = 0.1 * w
ListView1.ColumnHeaders(6).Width = 0.15 * w
ListView1.ColumnHeaders(7).Width = 0.2 * w


lv_recompute

End Sub

Private Sub Form_Click()
lv_recompute
End Sub

Private Sub Form_GotFocus()
lv_recompute
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Or KeyAscii = 27 Then

   If vbYes = MsgBox("Are you sure to remove this Item", vbYesNo + vbCritical, "Please Confirm") Then
   
            ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
             
            lv_recompute
            
            
        
   End If
End If
   
End Sub
Private Sub clearlist()

'Text1.Text = ""

EID.Text = ""
EName.Text = ""
grndtotal = 0
ListView1.ListItems.Clear
Image1.Picture = LoadPicture(App.Path & "\Photos\kao_Logo.jpg")
End Sub
