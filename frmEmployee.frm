VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmployee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee List"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9330
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Last Name"
         Object.Width           =   6237
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "First Name"
         Object.Width           =   6068
      EndProperty
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
DoEvents
End Sub

Private Sub Form_Load()
validpin = False
reflist2
 
 
End Sub
Private Sub reflist()
DoEvents
DoEvents
On Error Resume Next
Set rsACSearchSearch = Nothing

rsACSearchSearch.Open "SELECT * FROM vwemployeemaster order by emplname asc ", ac, adOpenStatic
   
       If rsACSearchSearch.RecordCount >= 1 Then
        ListView1.ListItems.Clear
                Do While Not rsACSearchSearch.EOF
                    Set lstitem = ListView1.ListItems.Add(, , rsACSearchSearch.Fields!empno)
                             lstitem.SubItems(1) = rsACSearch.Fields!emplname
                             lstitem.SubItems(2) = rsACSearch.Fields!empfname
                    rsACSearchSearch.MoveNext
               Loop
            Else
                ListView1.ListItems.Clear
                Set rsACSearch = Nothing
            End If
            
 Set rsACSearch = Nothing
            
    
    
End Sub
Private Function verifypin(ByVal empno As String, ByVal pincode As String)
Dim result As Boolean
result = False

If Val(empno) <> 0 Then
    
    Set rsPin = Nothing
    rsPin.Open "SELECT * FROM tbl_pincode WHERE PinCode = '" & Trim(pincode) & "' AND empno ='" & Val(Trim(empno)) & "'", ac, adOpenStatic
               If rsPin.RecordCount >= 1 Then
                             result = True
               End If
                
     Set rsPin = Nothing
End If

verifypin = result


End Function
Private Sub Form_Unload(Cancel As Integer)
        'FrmCredit.Text1.SetFocus
        FrmCredit.Timer3.Enabled = True
End Sub

Private Sub ListView1_DblClick()
If Not ListView1.ListItems.Count = 0 Then

Dim selitem As String
Dim selname As String

selitem = ListView1.SelectedItem
selname = ListView1.SelectedItem.SubItems(1) + " " + ListView1.SelectedItem.SubItems(2)


 
        With frmUnlisted
            .EID.Text = selitem
            .EName.Text = selname
        End With
        
        Unload Me
        
 
 

End If


 

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
 
'MsgBox KeyAscii
If KeyAscii = 13 Then
If Not ListView1.ListItems.Count = 0 Then


Dim selitem As String
Dim selname As String

selitem = ListView1.SelectedItem
selname = ListView1.SelectedItem.SubItems(1) + " " + ListView1.SelectedItem.SubItems(2)


pin_pinempid = ListView1.SelectedItem
frmverifiypin.Show 1
If validpin Then

With FrmCredit
             .Label11.Caption = selitem
            .Label12.Caption = selname
End With
Unload FrmSearchEmployee
FrmCredit.Text1.SetFocus

Else
MsgBox "Invalid Pin. Please try Again!"
End If

Else
MsgBox "No record selected.", vbInformation
Exit Sub
End If

End If


End Sub

Private Sub Text1_Change()
'If Not Text1.Text = "" Then
reflist2
DoEvents
'Else
'reflist
'End If

End Sub

Private Sub Text1_Click()
    Text1.Text = ""
End Sub
Private Sub reflist2()
 
'On Error GoTo err
 
DoEvents
rsACSearch.Open "SELECT * FROM vwemployeemaster where emplname like '" & Trim$(Text1.Text) & "%' OR empno like '" & Trim$(Text1.Text) & "%' order by emplname asc ", ac, adOpenStatic
   
       If rsACSearch.RecordCount >= 1 Then
        ListView1.ListItems.Clear
                Do While Not rsACSearch.EOF
                    Set lstitem = ListView1.ListItems.Add(, , rsACSearch.Fields!empno)
                             lstitem.SubItems(1) = rsACSearch.Fields!emplname
                             lstitem.SubItems(2) = rsACSearch.Fields!empfname
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

MsgBox "Error!", vbCritical + vbOKOnly, "Error"

  


End Sub

