VERSION 5.00
Begin VB.Form idplease 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLEASE TAP YOUR EMPLOYEE ID CARD ON THE RFID READER !"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6255
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   1200
      Width           =   6015
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   120
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7920
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3060
      Left            =   1440
      Picture         =   "idplease.frx":0000
      ScaleHeight     =   204
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   204
      TabIndex        =   2
      Top             =   1560
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SCANNING EMPLOYEE ID "
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   225
      TabIndex        =   0
      Top             =   120
      Width           =   5865
   End
End
Attribute VB_Name = "idplease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Dim counter As Integer


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete Or KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    FrmCredit.CZKEM1.EnableDevice vMachinenumber, False
    Unload Me
    
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Picture2_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = vbKeyDelete Or KeyAscii = 27 Then
        Unload Me
 End If
 
 
End Sub
 
Private Sub Timer2_Timer()
  
   If Picture2.Top > 4800 Then
     counter = -120
   End If
   If Picture2.Top <= 1200 Then
    counter = 120
   End If
   
   
   
   Picture2.Top = Picture2.Top + counter
 
'    If Shape1.Top = 960 And Shape1.Left = 240 Then
'
'        Shape1.Top = 2640
'        Shape1.Left = 3000
'
'        Shape2.Top = 960
'        Shape2.Left = 240
'
'    ElseIf Shape1.Top = 2640 And Shape2.Left = 3000 Then
'
'        Shape1.Top = 2640
'        Shape1.Left = 240
'
'        Shape2.Top = 960
'        Shape2.Left = 3000
'
'
'    ElseIf Shape1.Top = 2640 And Shape2.Left = 3000 Then
'        Shape1.Top = 2640
'        Shape1.Left = 240
'
'        Shape2.Top = 960
'        Shape2.Left = 3000
'
'    ElseIf Shape1.Top = 2640 And Shape2.Left = 240 Then
'
'        Shape1.Top = 960
'         Shape1.Top = 240
'
'
'        Shape2.Top = 2640
'         Shape2.Left = 3000
'
'
'    End If
 
End Sub
