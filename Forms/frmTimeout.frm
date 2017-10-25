VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Begin VB.Form frmTimeout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtuser 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtpassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   3135
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   855
      Left            =   120
      OleObjectBlob   =   "frmTimeout.frx":0000
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Due to inactivity current session log-out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Enter password :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmTimeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub form_activate()
txtpassword.Text = ""
txtpassword.SetFocus

End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
Dim machno As String

machno = 0

If KeyAscii = 13 Then

    If txtuser.Text = "" Then
                            MsgBox "Please enter user name.", 48, "Password Error"
                            txtuser.SetFocus
                            Exit Sub
                            Else
                                
                                str_User = "'" & Trim(FrmMainMenu.StatusBar1.Panels(2).Text) & "'"
                                str_Password = "'" & Trim(txtpassword.Text) & "'"
                                getusername = Trim(txtuser.Text)
                                
                                
                                Set rsLog_in = ac.Execute(" SELECT * from tbl_login where username = " & Trim$(str_User) & " and password = " & Trim(str_Password))
                                    If rsLog_in.EOF = False Then
                                        ' Screen.MousePointer = 11
                                         userlogin = True
                                         username_h = LCase(txtuser.Text)
                                                                                 
                                         Unload Me
                                         
                                         CZKEM1.EnableDevice machno, True
                                         MsgBox "Device connected", vbInformation
                                         FrmCredit.Timer3.Enabled = True
                                         
                                         
                                         
                                         FrmCredit.Timer2.Enabled = True
                                         'FrmMainMenu.Show
                                         'frmUsers.Show 1
                                         'FrmMainMenu.StatusBar1.Panels(2).Text = txtuser.Text
                                         
                                         'FrmMainMenu.Show 1
                             
                                        Exit Sub
                                    Else
                                            MsgBox "Invalid password.", 48, "Access Denied"
                                            txtpassword.SetFocus
                                            'SendKeys ("{home}+{end}")
                                    End If
                     End If
End If

End Sub

