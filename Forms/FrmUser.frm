VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User settings"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10954
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Admin / Users"
      TabPicture(0)   =   "FrmUser.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   5730
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10215
         Begin VB.CommandButton cmdclose 
            Caption         =   "Close"
            Height          =   855
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2760
            Width           =   1695
         End
         Begin VB.CommandButton cmddelete 
            Caption         =   "Delete"
            Height          =   855
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "Modify"
            Height          =   855
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "New"
            Height          =   855
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   1695
         End
         Begin VB.Frame Frame2 
            Caption         =   "Username"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   735
            Left            =   3840
            TabIndex        =   30
            Top             =   840
            Width           =   4215
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   3855
            End
            Begin VB.Label Label2 
               BackColor       =   &H00808080&
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   32
               Top             =   360
               Width           =   3855
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   735
            Left            =   3840
            TabIndex        =   27
            Top             =   1560
            Width           =   4215
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               IMEMode         =   3  'DISABLE
               Left            =   120
               PasswordChar    =   "*"
               TabIndex        =   28
               Top             =   195
               Width           =   3855
            End
            Begin VB.Label Label2 
               BackColor       =   &H00808080&
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   29
               Top             =   360
               Width           =   3855
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Confirm Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   735
            Left            =   3840
            TabIndex        =   24
            Top             =   2280
            Width           =   4215
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               IMEMode         =   3  'DISABLE
               Left            =   120
               PasswordChar    =   "*"
               TabIndex        =   25
               Top             =   240
               Width           =   3855
            End
            Begin VB.Label Label2 
               BackColor       =   &H00808080&
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   26
               Top             =   360
               Width           =   3855
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Complete Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   735
            Left            =   3840
            TabIndex        =   21
            Top             =   120
            Width           =   4215
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   3855
            End
            Begin VB.Label Label2 
               BackColor       =   &H00808080&
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   23
               Top             =   360
               Width           =   3855
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Access Forms"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1605
            Left            =   120
            TabIndex        =   5
            Top             =   3960
            Width           =   9975
            Begin VB.CheckBox recipe 
               Caption         =   "Recipe"
               Height          =   330
               Left            =   120
               TabIndex        =   20
               Top             =   210
               Width           =   1380
            End
            Begin VB.CheckBox Sales 
               Caption         =   "Sales"
               Height          =   435
               Left            =   105
               TabIndex        =   19
               Top             =   525
               Width           =   1065
            End
            Begin VB.CheckBox receipt 
               Caption         =   "Receipt"
               Height          =   435
               Left            =   120
               TabIndex        =   18
               Top             =   960
               Width           =   1170
            End
            Begin VB.CheckBox classification 
               Caption         =   "Classification"
               Height          =   330
               Left            =   1470
               TabIndex        =   17
               Top             =   210
               Width           =   1380
            End
            Begin VB.CheckBox Customer 
               Caption         =   "Customer"
               Height          =   330
               Left            =   1470
               TabIndex        =   16
               Top             =   630
               Width           =   1275
            End
            Begin VB.CheckBox Employee 
               Caption         =   "Employee"
               Height          =   435
               Left            =   1470
               TabIndex        =   15
               Top             =   945
               Width           =   1275
            End
            Begin VB.CheckBox expense 
               Caption         =   "Expense"
               Height          =   330
               Left            =   2835
               TabIndex        =   14
               Top             =   210
               Width           =   1170
            End
            Begin VB.CheckBox inventory 
               Caption         =   "Inventory"
               Height          =   330
               Left            =   2835
               TabIndex        =   13
               Top             =   630
               Width           =   1275
            End
            Begin VB.CheckBox orderslip 
               Caption         =   "Order slip gen."
               Height          =   330
               Left            =   2835
               TabIndex        =   12
               Top             =   1050
               Width           =   1485
            End
            Begin VB.CheckBox printing 
               Caption         =   "Printing"
               Height          =   330
               Left            =   4410
               TabIndex        =   11
               Top             =   210
               Width           =   1275
            End
            Begin VB.CheckBox vendor 
               Caption         =   "Vendor"
               Height          =   330
               Left            =   4410
               TabIndex        =   10
               Top             =   630
               Width           =   1275
            End
            Begin VB.CheckBox users 
               Caption         =   "Users"
               Height          =   330
               Left            =   4410
               TabIndex        =   9
               Top             =   1050
               Width           =   1065
            End
            Begin VB.CheckBox purchase 
               Caption         =   "Purchase Inventory"
               Height          =   330
               Left            =   5670
               TabIndex        =   8
               Top             =   210
               Width           =   2010
            End
            Begin VB.CheckBox inventorylist 
               Caption         =   "Inventory List"
               Height          =   450
               Left            =   5670
               TabIndex        =   7
               Top             =   600
               Width           =   2010
            End
            Begin VB.CheckBox Assembly 
               Caption         =   "Assembly"
               Height          =   450
               Left            =   5670
               TabIndex        =   6
               Top             =   960
               Width           =   2010
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Login Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   735
            Left            =   3840
            TabIndex        =   2
            Top             =   3000
            Width           =   4215
            Begin VB.ComboBox Combo1 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "FrmUser.frx":001C
               Left            =   120
               List            =   "FrmUser.frx":0035
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   240
               Width           =   3855
            End
            Begin VB.Label Label2 
               BackColor       =   &H00808080&
               Height          =   300
               Index           =   4
               Left            =   240
               TabIndex        =   4
               Top             =   360
               Width           =   3855
            End
         End
         Begin MSComctlLib.ListView listUsers 
            Height          =   3495
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   6165
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   3705
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Login Type"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'MsgBox "Ambot"
End Sub
Private Sub clearusers()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
recipe.Value = 0
Sales.Value = 0
receipt.Value = 0
classification.Value = 0
Customer.Value = 0
Employee.Value = 0
expense.Value = 0
inventory.Value = 0
orderslip.Value = 0
printing.Value = 0
vendor.Value = 0
users.Value = 0
purchase.Value = 0
recipe.Value = 0
inventorylist.Value = 0
Assembly.Value = 0

End Sub
Private Sub enableuser()
Text1.BackColor = vbWhite
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite

Combo1.BackColor = vbWhite
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True

Combo1.Enabled = True
End Sub
Private Sub disableuser()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Combo1.Enabled = False
Text4.Enabled = False

Text1.BackColor = &HE0E0E0
Text2.BackColor = &HE0E0E0
Text3.BackColor = &HE0E0E0
Combo1.BackColor = &HE0E0E0
Text4.BackColor = &HE0E0E0

End Sub

Private Sub cmdadd_Click()
    If cmdadd.Caption = "&Add" Then
        
    enableuser
    clearusers
    cmdadd.Caption = "&Save"
    cmdclose.Caption = "&Cancel"
    
    Text4.SetFocus
        ElseIf cmdadd.Caption = "&Save" Then
            If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
                MsgBox "Please Complete the Fields..!!", vbInformation, "Cannot Save"
                If Text1.Text = "" Then
                    Text1.SetFocus
                ElseIf Text2.Text = "" Then
                
                    Text2.SetFocus
                Else
                    Text3.SetFocus
                End If
            ElseIf Trim$(Text2.Text) <> Trim$(Text3.Text) Then
                MsgBox "Access Code and Confirmation Not the Same..!!", vbInformation, "Cannot Save"
                Text3.SetFocus
            ElseIf Combo1.Text = "" Then
                MsgBox "Login Type Missing..!!", vbInformation, "Cannot Save"
           
        Else
        
        sqlstring = "INSERT INTO tbl_LogIn values ('" & Trim$(Text1.Text) & "','" & Trim$(Text2.Text) & "','" & UCase(Trim$(Text4.Text)) & "','" & Combo1.Text & "','" & recipe.Value & "','" & Sales.Value & "','" & receipt.Value & "','" & classification.Value & "','" & Customer.Value & "','" & Employee.Value & "','" & expense.Value & "','" & inventory.Value & "','" & orderslip.Value & "','" & printing.Value & "','" & vendor.Value & "','" & users.Value & "','" & purchase.Value & "','" & inventorylist.Value & "','" & Assembly.Value & "')"
        
        ac.Execute sqlstring
        MsgBox " New User Added.!!", vbInformation, "Save"
        'CheckUserType
        
        
        disableuser
        clearusers
        refUsers
        cmdadd.Caption = "&Add"
        cmdclose.Caption = "&Close"
       
      End If
    End If
End Sub


Private Sub cmdclose_Click()
  
    If cmdclose.Caption = "&Cancel" Then
            clearusers
            disableuser
            cmdadd.Caption = "&Add"
            cmdadd.Enabled = True
            cmdedit.Enabled = False
            cmddelete.Enabled = False
            cmdedit.Caption = "&Modify"
    cmdclose.Caption = "&Close"
    Else
    Unload Me
    
    End If
End Sub

Private Sub cmdDelete_Click()
    'rsAC.Open "SELECT Name FROM tbl_LogIn", AC, adOpenStatic
  
    If MsgBox("Do you really want to delete " & UCase(Text1.Text) + "---" + Combo1.Text + "?", vbYesNo, "Confirm Delete") = vbYes Then
   
    sqlstring = "Delete from tbl_LogIn where Name = '" & Text1.Text & "'"
    ac.Execute sqlstring
    'MsgBox "File Successfully DELETED!!!", vbInformation, "Deleted"
    refUsers
    cmdclose.Caption = "&Close"
    cmddelete.Enabled = False
    cmdedit.Enabled = False
    cmdadd.Enabled = True
    cmdadd.Caption = "&Add"
    cmdedit.Caption = "&Modify"
    clearusers
    End If
    'Set rsAC = Nothing
End Sub

Private Sub cmdEdit_Click()
If cmdedit.Caption = "&Modify" Then
    'Text2.SetFocus
    cmddelete.Enabled = False
    cmdedit.Caption = "&Update"
    Text2.Enabled = True
    Text3.Enabled = True
    Text2.BackColor = vbWhite
    Text3.BackColor = vbWhite
    
    Text2.SetFocus
    'enableuser
ElseIf cmdedit.Caption = "&Update" Then
    If (Trim$(Text2.Text)) <> (Trim$(Text3.Text)) Then
       MsgBox "Access Code and Confirmation Not the Same..!!", vbInformation, "Cannot Save"
       Text3.SetFocus
    Else
    ac.Execute "UPDATE tbl_LogIn SET assembly = '" & Assembly.Value & "', AccessCode ='" & Text2.Text & "', recipe = '" & recipe.Value & "', sales = '" & Sales.Value & "', receipt = '" & receipt.Value & "', classification = '" & classification.Value & "', customer = '" & Customer.Value & "', employee = '" & Employee.Value & "', expense = '" & expense.Value & "', inventory = '" & inventory.Value & "', orderslip = '" & orderslip.Value & "', printing = '" & printing.Value & "', vendor = '" & vendor.Value & "', users = '" & users.Value & "', purchase = '" & purchase.Value & "', inventorylist = '" & inventorylist.Value & "'  WHERE Name ='" & Text1.Text & "' and status = '" & Combo1.Text & "'"
    MsgBox "User Successfully Updated..!!", vbInformation, "Updated"
    cmdedit.Caption = "&Modify"
    cmdedit.Enabled = False
    cmdclose.Caption = "&Close"
    cmdadd.Enabled = True
    clearusers
    disableuser
    End If
End If

End Sub

Private Sub Combo1_gotfocus()
Combo1.BackColor = &HC0E0FF
End Sub
Private Sub Combo1_LostFocus()
Combo1.BackColor = vbWhite
End Sub

Private Sub form_activate()
     'Me.Top = 750
    ' Me.Left = 0
      'With frmUsers
      '  .Move (Screen.Width - Width) \ 2, (Screen.Height - 8000) \ 8
    'End With
End Sub
Private Sub Form_Load()
'Me.Top = 750
'Me.Left = 0
refUsers
disableuser
cmdedit.Enabled = False
cmddelete.Enabled = False
'CheckUserType

End Sub

Private Sub refUsers()
 rsAC.Open "SELECT * FROM tbl_Login ", ac, adOpenStatic
    Dim ctr As String
    Dim a As Integer
    'a = 1
    'ctr = 0
    
    If rsAC.RecordCount >= 1 Then
        listUsers.ListItems.Clear
                Do While Not rsAC.EOF
                    Set lstitem = listUsers.ListItems.Add(, , rsAC.Fields!UserName)
                        lstitem.SubItems(1) = rsAC.Fields!status
                        
                     
                    rsAC.MoveNext
                    'ctr = ctr + 1
                Loop
            Else
                listUsers.ListItems.Clear
            End If
            'lblcounter.Caption = ctr
            Set rsAC = Nothing
End Sub






Private Sub listUsers_Click()
If Not listUsers.ListItems.Count = 0 Then
    rsAC.Open "SELECT * FROM tbl_LogIn WHERE Name LIKE '" & listUsers.SelectedItem.Text & "'", ac
    If Not rsAC.EOF And Not listUsers.ListItems.Count = 0 Then
        Text1.Text = rsAC.Fields!Name
        Text2.Text = rsAC.Fields!AccessCode
        Combo1.Text = rsAC.Fields!status
        Text4.Text = rsAC.Fields!FullName
        
        
        recipe.Value = rsAC.Fields!recipe
        Sales.Value = rsAC.Fields!Sales
        receipt.Value = rsAC.Fields!receipt
        classification.Value = rsAC.Fields!classification
        Customer.Value = rsAC.Fields!Customer
        Employee.Value = rsAC.Fields!Employee
        expense.Value = rsAC.Fields!expense
        inventory.Value = rsAC.Fields!inventory
        orderslip.Value = rsAC.Fields!orderslip
        printing.Value = rsAC.Fields!printing
        vendor.Value = rsAC.Fields!vendor
        users.Value = rsAC.Fields!users
        purchase.Value = rsAC.Fields!users
        inventorylist.Value = rsAC.Fields!inventorylist
        Assembly.Value = rsAC.Fields!Assembly
        
    End If
    Set rsAC = Nothing
End If
cmdadd.Enabled = False
cmdedit.Enabled = True
cmddelete.Enabled = True
cmdclose.Caption = "&Cancel"


End Sub


Private Sub Text1_GotFocus()



Text1.BackColor = &HC0E0FF



End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys "{tab}"
  End If
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = vbWhite
End Sub

Private Sub Text2_GotFocus()
Text2.BackColor = &HC0E0FF
End Sub

Private Sub Text2_LostFocus()
Text2.BackColor = vbWhite
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = &HC0E0FF
End Sub
Private Sub Text3_LostFocus()
Text3.BackColor = vbWhite
End Sub

Private Sub Text4_GotFocus()
Text4.BackColor = &HC0E0FF
End Sub

Private Sub Text4_LostFocus()
Text4.BackColor = vbWhite
End Sub

 _
 _
 _
 _
 _
 _

