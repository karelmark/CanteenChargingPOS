VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMainMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome                                                          "
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   16065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   BeginProperty Font 
      Name            =   "MV Boli"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMainMenu.frx":0000
   LinkTopic       =   "PRIMOS"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   10830
   ScaleWidth      =   16065
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6885
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12144
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Transactions"
      TabPicture(0)   =   "FrmMainMenu.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Inquiry"
      TabPicture(1)   =   "FrmMainMenu.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Maintenance"
      TabPicture(2)   =   "FrmMainMenu.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6255
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Width           =   11415
         Begin VB.CommandButton Command8 
            Appearance      =   0  'Flat
            Caption         =   "Manual List Transactions"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   3840
            Width           =   3615
         End
         Begin VB.CommandButton Command23 
            Appearance      =   0  'Flat
            Caption         =   "Upload Access Offline "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1920
            Width           =   3615
         End
         Begin VB.CommandButton Command22 
            Appearance      =   0  'Flat
            Caption         =   "Update offline table(s)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   2880
            Width           =   3615
         End
         Begin VB.CommandButton Command17 
            Appearance      =   0  'Flat
            Caption         =   "Back Up"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3840
            TabIndex        =   30
            Top             =   3960
            Width           =   3615
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3840
            TabIndex        =   24
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3840
            TabIndex        =   23
            Top             =   2040
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.Label Label21 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3840
            TabIndex        =   22
            Top             =   3000
            Visible         =   0   'False
            Width           =   3615
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6255
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   11415
         Begin VB.CommandButton Command7 
            Appearance      =   0  'Flat
            Caption         =   "Charges"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3840
            TabIndex        =   27
            Top             =   1920
            Width           =   3735
         End
         Begin VB.CommandButton Command10 
            Appearance      =   0  'Flat
            Caption         =   "Sales"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3840
            TabIndex        =   16
            Top             =   960
            Width           =   3735
         End
         Begin VB.CommandButton Command9 
            Appearance      =   0  'Flat
            Caption         =   "Reports"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3840
            TabIndex        =   14
            Top             =   2880
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3960
            TabIndex        =   28
            Top             =   2040
            Width           =   3735
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3960
            TabIndex        =   17
            Top             =   3000
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3960
            TabIndex        =   15
            Top             =   1080
            Width           =   3735
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   11175
         Begin VB.CommandButton Command6 
            Appearance      =   0  'Flat
            Caption         =   "F4 - Actual Payment         "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   4680
            Width           =   3735
         End
         Begin VB.CommandButton Command5 
            Appearance      =   0  'Flat
            Caption         =   "F5 -Charges                      "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   9840
            Width           =   3375
         End
         Begin VB.CommandButton Command4 
            Appearance      =   0  'Flat
            Caption         =   "F4 - Set Cut-off                    "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   8640
            Width           =   3375
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            Caption         =   "F3 - Inventory                     "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   3600
            Width           =   3735
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            Caption         =   "F2 - Transaction Records"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   2520
            Width           =   3735
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Caption         =   "F1 - Canteen Charging     "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   1440
            Width           =   3735
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3840
            TabIndex        =   26
            Top             =   4920
            Width           =   3735
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   13
            Top             =   10080
            Width           =   3375
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3840
            TabIndex        =   10
            Top             =   3840
            Width           =   3735
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3840
            TabIndex        =   0
            Top             =   2760
            Width           =   3735
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3840
            TabIndex        =   5
            Top             =   1560
            Width           =   3735
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   10215
      Width           =   16065
      _ExtentX        =   28337
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13318
            MinWidth        =   13318
            Text            =   "Canteen Charging System | Pilipinas Kao , Inc."
            TextSave        =   "Canteen Charging System | Pilipinas Kao , Inc."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "9/15/2016"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "8:17 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   25
      Top             =   6480
      Width           =   3735
   End
   Begin VB.Menu cmdaccount 
      Caption         =   "Account"
      Visible         =   0   'False
   End
   Begin VB.Menu Transactions 
      Caption         =   "Transactions"
      Begin VB.Menu Pr 
         Caption         =   "Canteen Charging"
         Shortcut        =   {F1}
      End
      Begin VB.Menu TR 
         Caption         =   "Transaction Records"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Inventory 
         Caption         =   "Inventory"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Apayment 
         Caption         =   "Actual Payment"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu Inq 
      Caption         =   "Inquiry"
   End
   Begin VB.Menu user 
      Caption         =   "User"
      Begin VB.Menu logout 
         Caption         =   "&Logout"
      End
      Begin VB.Menu turnoff 
         Caption         =   "Turn Off Computer"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FrmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x_count As Integer
 

Private Sub Apayment_Click()
    frmActualPayment.Show 1
End Sub

 

Private Sub Command1_Click()
    
    Dim dwFlags As Long
    Dim sNameBuf As String
    Dim lPos As Long
    sNameBuf = String$(513, 0)
    FrmCredit.Show 1
End Sub
Private Sub Command10_Click()
     FrmSales.Show 1
End Sub
 
 
Private Sub Command8_Click()
frmUnlisted.Show 1
End Sub

Private Sub Form_Activate()
    StatusBar1.Panels(3).Text = LCase(username_h)
      
End Sub
Private Sub Command2_Click()
    frmTransactionRecords.Show 1
End Sub
Private Sub Command22_Click()
    frm_updateoff.Show 1
End Sub
Private Sub Command23_Click()
    frm_uploadacess.Show 1
End Sub
Private Sub Command3_Click()
    FrmInventory.Show 1
End Sub
Private Sub Command4_Click()
    FrmCutoff.Show 1
End Sub
Private Sub Command6_Click()
    frmActualPayment.Show 1
End Sub
Private Sub Command7_Click()
    frm_summaryofcharge.Show 1
End Sub
Private Sub Command9_Click()
    FrmPrint.Show 1
End Sub
Private Sub Cutoff_Click()
    FrmCutoff.Show 1
End Sub
Private Sub Exit_Click()
    End
End Sub
Private Sub Form_Load()

    Dim sDate As String
    Dim sDate1 As String
    sDate = Format(Now, "mm") & "/01/" & Format(Now, "yyyy")
    gICDatePrev = sDate
    If Not Val(Format(Now, "mm")) = 1 Then
        sDate1 = (Val(Format(Now, "mm")) - 1) & "/01/" & Format(Now, "yyyy")
    Else
        sDate1 = (Val(Format(Now, "mm")) - 1) & "/01/" & (Val(Format(Now, "yyyy")) - 1)
    End If
    
    gICDateCurr = sDate1
    
     
    x_count = 0
     
    Dim w As Double
    
    w = Me.Width
    StatusBar1.Panels(1).Width = 0.7 * w
    StatusBar1.Panels(2).Width = 0.05 * w
    StatusBar1.Panels(3).Width = 0.1 * w
    StatusBar1.Panels(4).Width = 0.15 * w

End Sub
Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
If MsgBox("Are you sure you want to Logout? ", vbYesNo + vbQuestion) = vbYes Then
    End
Else
   Cancel = True  'Returns to the Current form
End If

End Sub
 
Private Sub Inventory_Click()
    FrmInventory.Show 1
End Sub

Private Sub logout_Click()

    Me.Hide
    Frmpassword.Show 1

End Sub
 
 
Private Sub Timer1_Timer()
    StatusBar1.Panels.Item(5).Text = Time
End Sub
 

Private Sub turnoff_Click()
    
    If "gracias" = InputBox("Please Input Password to TURN OFF THE COMPUTER:", "Please Confirm") Then
            Shell ("shutdown.exe -s")
    Else
            MsgBox "TURN OFF FAIL!", vbOKOnly + vbCritical, "Invalid Password!"
    End If

        
End Sub

 
