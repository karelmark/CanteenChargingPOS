VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmadditem 
   Caption         =   "Available Items"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Top             =   525
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13785
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Barcode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Itemcode"
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item Name"
         Object.Width           =   6237
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "UnitCode"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmadditem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub reflist()
'On Error GoTo err

rsACSearch.Open "SELECT * FROM tbl_inventory where itemname like '" & Trim$(Text1.Text) & "%'", ac, adOpenStatic
   
       If rsACSearch.RecordCount >= 1 Then
        ListView1.ListItems.Clear
                Do While Not rsACSearch.EOF
                    Set lstitem = ListView1.ListItems.Add(, , rsACSearch.Fields!barcodeno)
                             
                             lstitem.SubItems(1) = rsACSearch.Fields!recno
                             
                             lstitem.SubItems(2) = rsACSearch.Fields!itemname
                             
                             lstitem.SubItems(3) = FormatNumber(rsACSearch.Fields!Price, 2)
                             lstitem.SubItems(4) = rsACSearch.Fields!unitcode
                             
                    rsACSearch.MoveNext
                    
                Loop
            Else
                ListView1.ListItems.Clear
                Set rsACSearch = Nothing
            End If
            
'rsACSearch.Close
Set rsACSearch = Nothing
           
Exit Sub

err:

MsgBox "Error! reflist", vbCritical + vbOKOnly, "Error"

    
    
End Sub

 
Private Sub Form_Load()
Call reflist
End Sub

Private Sub ListView1_DblClick()
If Not ListView1.ListItems.Count = 0 Then

Dim selitem As String
Dim selname As String
Dim a As ListItem

selitem = ListView1.SelectedItem
selname = ListView1.SelectedItem.SubItems(2)

Dim qty As Double
qty = InputBox("Enter Quantity", "Provide # of items", 1)
Dim sbtotal As Double

 
        With frmUnlisted
            Set a = .ListView1.ListItems.Add(, , "")
              
            a.SubItems(1) = ListView1.SelectedItem.SubItems(1)
            
            a.SubItems(2) = selname
            
            a.SubItems(3) = ListView1.SelectedItem.SubItems(3)
            
            a.SubItems(4) = qty
            
            a.SubItems(5) = ListView1.SelectedItem.SubItems(4)
            
            sbtotal = CDbl(ListView1.SelectedItem.SubItems(3)) * qty
            
            a.SubItems(6) = Format$(sbtotal, "0.00")
            
        End With
        
        Unload Me
        'FrmCredit.Text1.SetFocus
        'FrmCredit.Timer3.Enabled = True

Else
MsgBox "No record selected.", vbInformation
Exit Sub
End If
End Sub

Private Sub Text1_Change()
Call reflist
End Sub

