VERSION 5.00
Begin VB.Form frmAccesscode 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2.117
   ScaleMode       =   0  'User
   ScaleWidth      =   4.657
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESS CODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Height          =   3135
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmAccesscode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = ""
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Text1.Text = "" Then
            MsgBox "Please enter access code.", vbCritical, "AccessCode Error"
            Exit Sub
        End If
    'rsAC.Open "Select FullName,LoginType from tbl_LogIn Where AccessCode = '" & Trim$(Text1.Text) & "',ac  "
    
    rsAC2.Open "Select * from tbl_LogIn where userlevel = 'Administrator' and password = '" & Trim$(Text1.Text) & "'", ac, adOpenStatic
    'If rsAC2.EOF = True Then
     '   MsgBox "Invalid Access Code..!!", vbCritical, "Access Denied"
    If Not rsAC2.EOF Then
       Unload Me
       MsgBox "You can now cancel selected transaction no."
       
       frmTransactionRecords.Label9.Caption = "1"
       
        'frmNewUnitPrice.Show 1
        'FrmIncome.Show 1
  
    Else
     Text1.Text = ""
     MsgBox "Invalid access code.", vbCritical, "Access Denied"
    
    End If
    
    
        Set rsAC2 = Nothing
    End If
End Sub

