VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Frmpassword 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pilipinas Kao, Inc "
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5220
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   Icon            =   "frmpassword.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "login"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmpassword.frx":076A
   ScaleHeight     =   3195
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   13
      PasswordChar    =   "="
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   20
      Scrolling       =   1
   End
   Begin VB.Frame Frame3 
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
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4980
      Begin LVbuttons.LaVolpeButton cmdLogin 
         Default         =   -1  'True
         Height          =   360
         Left            =   3120
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "Login"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmpassword.frx":0BAC
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.TextBox txtuser 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   240
         Top             =   2040
      End
      Begin LVbuttons.LaVolpeButton cmdCancel 
         Height          =   360
         Left            =   -360
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmpassword.frx":0BC8
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.CommandButton cmdCancel1 
         Caption         =   "Cancel"
         Height          =   300
         Left            =   -240
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdLogin1 
         Caption         =   "Login"
         Height          =   300
         Left            =   -240
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Not yet Register? Click Here"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   -240
         TabIndex        =   13
         ToolTipText     =   "Click Here"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "frmpassword.frx":0BE4
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "Label4"
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   2385
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "Label4"
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   2385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   1080
         Width           =   690
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANTEEN CHARGING SYSTEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "Frmpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Check for network availability

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdLogin_Click()

On Error Resume Next

Dim str_User, str_Password
Dim dwFlags As Long
Dim sNameBuf As String
Dim lPos As Long
sNameBuf = String$(513, 0)
 
                    If txtuser.Text = "" Then
                            MsgBox "Please enter user name.", 48, "Password Error"
                            txtuser.SetFocus
                            Exit Sub
                            Else
                                
                                str_User = "'" & Trim(txtuser.Text) & "'"
                                str_Password = "'" & Trim(txtpassword.Text) & "'"
                                getusername = Trim(txtuser.Text)
                                 
                                Set rsLog_in = ac.Execute(" SELECT * from tbl_login where username = " & Trim$(str_User) & " and password = " & Trim(str_Password))
                                    If rsLog_in.EOF = False Then
                                         ' Screen.MousePointer = 11
                                         userlogin = True
                                         username_h = LCase(txtuser.Text)
                                         Me.Visible = False
                                         FrmMainMenu.StatusBar1.Panels(2).Text = txtuser.Text
                                         Unload Me
                                         
                                         FrmMainMenu.Show 1
                                     Exit Sub
                                    Else
                                            MsgBox "Invalid username or password.", 48, "Access Denied"
                                            txtpassword.SetFocus
                                            
                                            
                                            SendKeys ("{home}+{end}")
                                    End If
                     End If

            
End Sub

Private Sub Command1_Click()

'
'    SelectPrinter = True
'    For i = 0 To Printers.count - 1
'        If Printers(i).DeviceName = printer_name Then
'            Set Printer = Printers(i)
'            SelectPrinter = False
'            Exit For
'        End If
'    Next i
'
'
'Exit Function
'Dim str As String
'Dim prntr As Printer
'For Each prntr In Printers
'
''MsgBox prntr.DeviceName
'str = prntr.DeviceName
'Next
'
'For Each prntr In Printers
'    If str = prntr.DeviceName Then
'    Set Printer = prntr
'    End If
'Next

'IssuanceReport "ISS-22281"
'set_POapprover "038666"

End Sub


Private Sub Form_Unload(Cancel As Integer)
 'End
End Sub
 

Private Sub Form_Load()
dataconnect
Set RS = Nothing

End Sub

Private Sub LaVolpeButton1_Click()
End
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
'On Error GoTo wala
'dataconnect

'Dim strvalid
'strvalid = "0123456789.qwertyuiop[]lkjhgfdsazxcvbnmQWERTYUIOPLKJHGFDSAZXCVBNM[];,./?><{}"
'If KeyAscii > 26 Then
'    If InStr(strvalid, Chr(KeyAscii)) = 0 Then
'       KeyAscii = 0
'   End If
'End If
'
'If txtuser.Text = "paranopa" And txtpassword.Text = "shipwreck" Then
'    Unload Me
'    FrmMainMenu.Show 1
'
'End If
'
'Exit Sub
End Sub



Private Sub txtUser_GotFocus()
'txtUser.BackColor = &HC0E0FF
txtuser.SelStart = 0
txtuser.SelLength = Len(txtuser.Text)
End Sub

Private Sub txtUser_Keypress(KeyAscii As Integer)
    
    Dim strvalid
    strvalid = "0123456789.qwertyuiop[]lkjhgfdsazxcvbnmQWERTYUIOPLKJHGFDSAZXCVBNM[];,./?><{}"
    If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    End If

    If KeyAscii = 27 Then
        End
    Exit Sub
    End If
    If KeyAscii = 13 Then
  SendKeys "{tab}"
  Else
  tCase txtuser, KeyAscii
End If
 End Sub


                           
