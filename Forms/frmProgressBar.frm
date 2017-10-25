VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgressBar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1800
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   1320
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   10
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Processing ........"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ProgressBar1.Max = 20
    ProgressBar1.Min = 0
End Sub



Private Sub Timer1_Timer()
    If ProgressBar1.Value < 20 Then
        ProgressBar1.Value = ProgressBar1.Value + 1
    Else
        Timer1.Enabled = False
        Unload Me
   
    
    End If
End Sub

Private Sub Timer2_Timer()
If Timer2.Interval = 100 Then
       Label2.Visible = Not Label2.Visible

    End If
End Sub
