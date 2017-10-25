VERSION 5.00
Begin VB.Form frmverifiypin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Verify"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   840
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER YOUR PINCODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmverifiypin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Click()
    SendKeys ("{Home} + {End}")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 pin_pincode = Text1.Text
 
 If verifypin(pin_pinempid, pin_pincode) Then
        validpin = True
        
 Else
       validpin = False
 End If
 
 DoEvents
 Unload Me
 
 
 
 
ElseIf KeyAscii = vbKeyDelete Or KeyAscii = 27 Then
 Unload Me
End If

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
