VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmInventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Inventory Record"
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
      Height          =   6255
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton Command2 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtsearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cbosearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmInventory.frx":0000
         Left            =   5160
         List            =   "FrmInventory.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   240
         Width           =   2895
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5295
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   9340
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Itemname"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Unitcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Barcode No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ClassID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Classification"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Added By"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Inactive"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Search :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   26
         Top             =   315
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "Search By :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4095
         TabIndex        =   25
         Top             =   315
         Width           =   1095
      End
   End
   Begin VB.CommandButton Cmd_list 
      Caption         =   "List"
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
      Left            =   4200
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Cmd_Save 
      Caption         =   "Save"
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
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton Cmd_Edit 
      Caption         =   "Edit"
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   10
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Cmd_Delete 
      Caption         =   "Delete"
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Cmd_Add 
      Caption         =   "Add"
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
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1575
      TabIndex        =   3
      Top             =   600
      Width           =   4560
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   1575
      TabIndex        =   0
      Top             =   120
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
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
      Left            =   4200
      TabIndex        =   13
      Top             =   5280
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "FrmInventory.frx":0028
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text71"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Combo2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "text7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.ComboBox text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "FrmInventory.frx":0044
         Left            =   1800
         List            =   "FrmInventory.frx":0063
         TabIndex        =   6
         Text            =   "pc"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   7
         Top             =   2400
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "FrmInventory.frx":00A3
         Left            =   1800
         List            =   "FrmInventory.frx":00A5
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   3210
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text71 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4800
         TabIndex        =   29
         Text            =   "pc"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5040
         TabIndex        =   28
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode No.  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   27
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unit  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Classification:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "In-Active:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblItemName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   18
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblItemID 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Item ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6375
      TabIndex        =   17
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "FrmInventory"
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
Dim rsrs As New ADODB.Recordset


Dim tcode As String
Dim prevbal As Double

Dim prevsubsidy As Double
Dim prevsubtotal As Double
Private Sub Cmd_Add_Click()
EnableText
Cmd_Save.Enabled = True
Command1.Caption = "Cancel"
Cmd_Edit.Enabled = False
Cmd_Delete.Enabled = False
ItemCodes
Text2.SetFocus
End Sub
Private Sub ItemCodes()
           ' trucker id
         Set RS = Nothing
         RS.Open "select max(recno) as xidno from tbl_inventory", ac, adOpenStatic, adLockOptimistic
         
         
         If Not RS.EOF Then
         
         Text1.Text = RS("xidno") + 1
         End If
         
                              
                'ending
End Sub


Private Sub Cmd_Delete_Click()
If MsgBox("Do you really want to delete  " & Text1.Text + " " + "?", vbYesNo, "Confirm Delete") = vbYes Then
 
   
ac.Execute "delete  from tbl_inventory where recno = '" & Text1.Text & "'"

reflist
frmProgressBar.Show 1
DisableText

Cmd_Edit.Enabled = False
Cmd_Delete.Enabled = False
Command1.Caption = "Close"
End If
End Sub
Private Sub EnableText()
On Error Resume Next
 
Text1.Text = ""
Text2.Text = ""
Text3.Text = "0"
Text4.Text = "0"
text7.Text = "pc"

 
text7.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Combo2.Enabled = True
Combo2.BackColor = vbWhite

text7.BackColor = vbWhite
 
'Text1.BackColor = vbWhite
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Cmd_Add.Enabled = False

Text2.SetFocus
End Sub

Private Sub DisableText()
On Error Resume Next

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
text7.Enabled = False

Combo2.Enabled = False
Combo2.Text = ""


Combo2.BackColor = &HC0C0C0
'Text1.BackColor = &HC0C0C0
Text2.BackColor = &HC0C0C0
Text3.BackColor = &HC0C0C0
Text4.BackColor = &HC0C0C0
text7.BackColor = &HC0C0C0

Text1.Text = ""
Text2.Text = ""
Text3.Text = "0"
Text4.Text = "0"
text7.Text = "pc"

End Sub

Private Sub Cmd_Edit_Click()
Dim status As String
 
If Check1.Value = 1 Then
status = "INACTIVE"
ElseIf Check1.Value = 0 Then
status = "ACTIVE"
End If

ac.Execute "UPDATE tbl_inventory set inactive = '" & status & "', itemname = '" & UCase(Text2.Text) & "', classification = '" & Combo2.Text & "', price = '" & Text3.Text & "', unitcode =  '" & text7.Text & "' , class_id =  '" & Label3.Caption & "' , barcodeno =  '" & Text4.Text & "' WHERE  recno =  " & Text1.Text

reflist

frmProgressBar.Show 1
DisableText

Cmd_Edit.Enabled = False
Cmd_Delete.Enabled = False
Command1.Caption = "Close"
Cmd_Add.Enabled = True


End Sub

Private Sub Cmd_list_Click()
Frame1.Visible = True

End Sub

Private Sub Cmd_Save_Click()

On Error GoTo errhandler:


Dim status As String

If Check1.Value = 1 Then
status = "INACTIVE"
ElseIf Check1.Value = 0 Then
status = "ACTIVE"
End If

If Text1.Text = "" Then
MsgBox "Please enter item number.", vbInformation
ElseIf Combo2.Text = "" Then
MsgBox "Please select classification name.", vbInformation
Exit Sub
Else
Dim icode As String

rs4.Open "Select TOP 1 recno FROM tbl_inventory ORDER BY recno DESC", acaccess, adOpenForwardOnly, adLockPessimistic

If Not rs4.EOF Then
        icode = CStr(rs4.Fields!recno + 1)
Else
        icode = "1000"
End If


acaccess1.Execute "insert into tbl_inventory(recno,itemname, unitcode, price, barcodeno, addedby, addeddate, addedtime, inactive, class_id, classification) values ('" & icode & "','" & UCase(Text2.Text) & "','" & UCase((text7.Text)) & "','" & Text3.Text & "','" & Text4.Text & "','" & FrmMainMenu.StatusBar1.Panels(2).Text & "','" & Date & "','" & Time & "','" & status & "','" & Label3.Caption & "','" & Combo2.Text & "')"
 
            

reflist
frmProgressBar.Show 1
DisableText
Cmd_Save.Enabled = False
Command1.Caption = "Close"
Cmd_Add.Enabled = True

 

End If

Exit Sub
errhandler:
MsgBox err.Description, vbExclamation
Exit Sub

End Sub
Private Sub reflist()

On Error GoTo error:

 If rsAC.State = adStateOpen Then
            rsAC.Close
  End If
  Dim lstitem As ListItem
rsAC.Open "SELECT * FROM tbl_inventory order by ItemName asc ", ac, adOpenStatic
   
       If rsAC.RecordCount >= 1 Then
        ListView2.ListItems.Clear
                
                Do While Not rsAC.EOF
                    Set lstitem = ListView2.ListItems.Add(, , rsAC.Fields!recno)
                             
                             lstitem.SubItems(1) = rsAC.Fields!itemname
                             lstitem.SubItems(2) = rsAC.Fields!unitcode
                             lstitem.SubItems(3) = FormatNumber(rsAC.Fields!Price)
                             lstitem.SubItems(4) = rsAC.Fields!barcodeno
                             lstitem.SubItems(5) = rsAC.Fields!class_id
                             lstitem.SubItems(6) = rsAC.Fields!classification
                             lstitem.SubItems(7) = rsAC.Fields!addedby
                             lstitem.SubItems(8) = rsAC.Fields!inactive
                                
                    rsAC.MoveNext
                    
                Loop
            Else
                ListView2.ListItems.Clear
                Set rsAC = Nothing
            End If
            
            Set rsAC = Nothing
Exit Sub
            
error:
            MsgBox err.Description, vbCritical, "Oops!"
            
End Sub
Private Sub Combo2_Click()
Set rsAC = Nothing
rsAC.Open "Select * FROM tbl_classification where classification = '" & Combo2.Text & "'  order by classification asc", ac, adOpenStatic
   
   If Not rsAC.EOF Then
    Label3.Caption = rsAC.Fields!recno
   End If
   
    Set rsAC = Nothing
End Sub

Private Sub Combo2_DropDown()
rsAC.Open "Select * FROM tbl_classification  order by classification asc", ac, adOpenStatic
    Combo2.Clear
    Do While Not rsAC.EOF
        Combo2.AddItem (rsAC.Fields!classification)
    
        rsAC.MoveNext
    Loop
    Set rsAC = Nothing
End Sub

Private Sub Combo2_GotFocus()
Combo2.SelStart = 0
Combo2.SelLength = Len(Combo2.Text)
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Cancel" Then
    DisableText
    Cmd_Save.Enabled = False
    Cmd_Edit.Enabled = False
    Cmd_Delete.Enabled = False
    Command1.Caption = "Close"
    Cmd_Add.Enabled = True
    
    
ElseIf Command1.Caption = "Close" Then

Unload Me

End If
End Sub

Private Sub Command2_Click()
Frame1.Visible = False

End Sub

Private Sub rsfilter()
    Dim sby As String
    If cbosearch.ListIndex = -1 Then cbosearch.ListIndex = 0
    
 Select Case cbosearch.ListIndex

Case 0
sby = "classification"
Case 1
sby = "itemname"



End Select
If cbosearch.ListIndex = 0 Then
    rsAC.Open "select * from tbl_inventory where " & sby & " like '" & txtsearch & "%' order by itemname asc", ac, adOpenStatic, adLockReadOnly
    fill_lst
    Set rsAC = Nothing
End If

If cbosearch.ListIndex = 1 Then
    rsAC.Open "select * from tbl_inventory where " & sby & " like '" & txtsearch & "%' order by itemname asc", ac, adOpenStatic, adLockReadOnly
    fill_lst
    Set rsAC = Nothing
End If




End Sub
Private Sub fill_lst()
Dim X As Integer
ListView2.ListItems.Clear
   Dim lstitem As ListItem
   Do While Not rsAC.EOF
                    Set lstitem = ListView2.ListItems.Add(, , rsAC.Fields!recno)
                             lstitem.SubItems(1) = rsAC.Fields!itemname
                             lstitem.SubItems(2) = rsAC.Fields!unitcode
                             lstitem.SubItems(3) = FormatNumber(rsAC.Fields!Price)
                             lstitem.SubItems(4) = rsAC.Fields!barcodeno
                            ' lstitem.SubItems(4) = rsAC.Fields!status
                             lstitem.SubItems(5) = rsAC.Fields!class_id
                             lstitem.SubItems(6) = rsAC.Fields!classification
                             lstitem.SubItems(7) = rsAC.Fields!addedby
                             lstitem.SubItems(8) = rsAC.Fields!inactive
                             
                             'lstitem.SubItems(8) = rsAC.Fields!classification
                             
                             
                     
                    rsAC.MoveNext
                    
                Loop


End Sub

Private Sub Form_activate()
'Text2.SetFocus

End Sub

Private Sub Form_Load()
connectData
reflist
DisableText
Cmd_Save.Enabled = False


'DTPicker1.Value = Date
End Sub

Private Sub ListView2_Click()
If Not ListView2.ListItems.Count = 0 Then
        EnableText
        Command1.Caption = "Cancel"
        
        Cmd_Edit.Enabled = True
        Cmd_Delete.Enabled = True
        Cmd_Save.Enabled = False
        
       
        Text1.Text = ListView2.SelectedItem
        Text2.Text = ListView2.SelectedItem.SubItems(1)
        Label3.Caption = ListView2.SelectedItem.SubItems(5)
        Combo2.Text = ListView2.SelectedItem.SubItems(6)
        
        Text3.Text = ListView2.SelectedItem.SubItems(3)
        text7.Text = ListView2.SelectedItem.SubItems(2)
        Text4.Text = ListView2.SelectedItem.SubItems(4)
        
        
        If ListView2.SelectedItem.SubItems(8) = "ACTIVE" Then
            Check1.Value = 0
        ElseIf ListView2.SelectedItem.SubItems(8) = "INACTIVE" Then
            Check1.Value = 1
        End If
        
        
Frame1.Visible = False
        
            
            
Else
Exit Sub

End If
End Sub

 

 
Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

 

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub text7_GotFocus()
text7.SelStart = 0
text7.SelLength = Len(text7.Text)
End Sub

Private Sub txtsearch_Change()
rsfilter
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
