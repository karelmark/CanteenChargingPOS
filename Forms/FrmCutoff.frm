VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCutoff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Cut-off"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16005
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   16005
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView2 
      Height          =   3495
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   6165
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Idno"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total Charges"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total Subsidy"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total OTsubsidy"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total Deductions "
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView5 
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "File transfer to Excel"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   9360
      Width           =   12375
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmCutoff.frx":0000
         Left            =   120
         List            =   "FrmCutoff.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   285
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11280
         Picture         =   "FrmCutoff.frx":0021
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Import to Excel"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         Picture         =   "FrmCutoff.frx":078B
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Save As"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtpath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Top             =   285
         Width           =   8535
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11760
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   7011
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TransNo"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Idno"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   5009
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Itemcode"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Itemname"
         Object.Width           =   4480
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Qty"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Price"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Subtotal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Subsidy"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cut-off Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   15735
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   495
         Left            =   13920
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Post Cut-Off"
         Height          =   495
         Left            =   12240
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   300
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   72286209
         CurrentDate     =   41659
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   72286209
         CurrentDate     =   41659
      End
      Begin VB.Label Label4 
         Caption         =   "To :"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "From :"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   1455
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   14280
      TabIndex        =   12
      Top             =   9960
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   14280
      TabIndex        =   11
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Total Subsidy :"
      Height          =   255
      Left            =   12600
      TabIndex        =   10
      Top             =   9960
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Total Records :"
      Height          =   255
      Left            =   12600
      TabIndex        =   9
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Canteen Charges Cut-off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16095
   End
End
Attribute VB_Name = "FrmCutoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Me.MousePointer = 11
reflist
'reflist2

reflist5
reflist_final
'reflist3
Me.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Command4_Click()
If Combo1.ListIndex = 0 Then
                who = 3
                If txtpath.Text = "" Then
                    MsgBox "Please enter an Extract file name and location.", vbExclamation, "File Name"
                Else
                    modExcel.SaveAsExcel RS, txtpath, Me.Caption, "YES"
                End If
ElseIf Combo1.ListIndex = 1 Then
            who = 4
            If txtpath.Text = "" Then
                MsgBox "Please enter an Extract file name and location.", vbExclamation, "File Name"
            Else
                modExcel.SaveAsExcel RS, txtpath, Me.Caption, "YES"
            End If
End If

End Sub

Private Sub Command5_Click()
On Error GoTo Err_Handler

CommonDialog1.CancelError = True
txtpath.Text = ""

CommonDialog1.Filename = Replace(Me.Caption, " ", "") & "_" & "sample" & "_" & Format(Now(), "ddmm") & ".xls"

CommonDialog1.Filter = "Microsoft Excel .xls (*.xls)|*.xls"
CommonDialog1.ShowOpen

If CommonDialog1.Filename <> "" Then txtpath.Text = CommonDialog1.Filename

Exit Sub

Err_Handler:

    If err = 32755 Then
        txtpath.Text = ""
    Else
        MsgBox "An error has occurred! " & vbCrLf & vbCrLf & err & ": " & Error & " ", vbExclamation
    End If
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date



End Sub
Private Sub reflist3()
ListView3.ListItems.Clear

I = 1
  Do While I <= ListView5.ListItems.Count
        ListView5.ListItems.Item(I).Selected = True
        
    Set rsAC = Nothing
            rsAC.Open "select * from vwovertime where empno = '" & ListView5.SelectedItem & "' and otfpay between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "' order by empno", ac, adOpenStatic
            
            Dim x_otfsub As Currency
            Dim x_otfhr As Currency
            
            x_otfsub = 0
            x_otfhr = 0
            
              Do While Not rsAC.EOF
                    
                         
                         
                         
                         
                       Set lstitem = ListView3.ListItems.Add(, , rsAC.Fields!empno)
                       
                             lstitem.SubItems(1) = (rsAC("otfhrs"))
                          
                                   
                    
                    Set rs2 = Nothing
                    rs2.Open "select otfhrs from vwovertime where otfid = '" & rsAC("otfid") & "' and empno =  '" & rsAC("empno") & "' ", ac, adOpenStatic, adLockReadOnly
                   '
                    If Not rs2.EOF Then
                        x_otfhr = rs2("otfhrs")
                   
                        If x_otfhr >= 2 And x_otfhr <= 9 Then
                        x_otfsub = 40
                        
                        'MsgBox x_otfsub
                        lstitem.SubItems(2) = x_otfsub
                        ElseIf x_otfhr >= 10 And x_otfhr <= 17 Then
                        x_otfsub = 80
                        lstitem.SubItems(2) = x_otfsub
                        ElseIf x_otfhr >= 18 And x_otfhr <= 24 Then
                        x_otfsub = 120
                        lstitem.SubItems(2) = x_otfsub
                        Else
                        lstitem.SubItems(2) = 0
                        
                        
                        End If
                        
                        
                    
                    End If
                                
                
                
                
                           
                
               
                rsAC.MoveNext
                Loop
            Set rsAC = Nothing
                
       
       
       
       
      
             
        I = I + 1
    Set rs1 = Nothing
  Loop
End Sub
Private Sub reflist_final()
Dim xsubtotal As Currency
Dim xsubsidy As Currency
Dim xdeduct As Currency



ListView2.ListItems.Clear
xsubtotal_ctr = 0
xsubsidy = 0
xdeduct = 0

I = 1
  Do While I <= ListView5.ListItems.Count
        ListView5.ListItems.Item(I).Selected = True
    xsubtotal_ctr = 0
        
    Set rs1 = Nothing
       rs1.Open "select transno, idno  from tbl_transaction where idno = '" & ListView5.SelectedItem & "' and transdate between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "' order by transno ", ac, adOpenStatic, adLockReadOnly
         Set lstitem_1 = ListView2.ListItems.Add(, , rs1.Fields!idno)
                    
             
             Set rs3 = Nothing
                    
                    rs3.Open "select * from vwemployeemaster where empno = '" & rs1("idno") & "'", ac, adOpenStatic, adLockReadOnly
                    If Not rs3.EOF Then
                        lstitem_1.SubItems(1) = rs3.Fields!emplname + ", " + rs3.Fields!empfname + " " + rs3.Fields!empMname
                    End If
            
             Set rs3 = Nothing
                
                Do While Not rs1.EOF
                    Set rs2 = Nothing
                    rs2.Open "select transno , sum(subtotal) as xsubtotal from tbl_transdetails where transno = '" & rs1("transno") & "' group by transno", ac, adOpenStatic, adLockReadOnly
                    If Not rs2.EOF Then
                        xsubtotal_ctr = xsubtotal_ctr + rs2("xsubtotal")
                    End If
                                
                rs1.MoveNext
                Loop
                
            
                
                
                lstitem_1.SubItems(2) = FormatNumber(xsubtotal_ctr)
          lstitem_1.SubItems(3) = FormatNumber(ListView5.SelectedItem.SubItems(1))
         
       
        
   
       
       
            Set rsAC = Nothing
            rsAC.Open "select * from vwovertime where empno = '" & ListView5.SelectedItem & "' and otfpay between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "' order by empno", ac, adOpenStatic
            
            Dim x_otfsub As Currency
            Dim x_otfhr As Currency
            Dim x_otcounter As Currency
            
            x_otfsub = 0
            x_otfhr = 0
            x_otcounter = 0
            
              Do While Not rsAC.EOF
                    
                         
                         
                         
                         
                       Set lstitem = ListView3.ListItems.Add(, , rsAC.Fields!empno)
                       
                             lstitem.SubItems(1) = (rsAC("otfhrs"))
                          
                                   
                    
                    Set rs2 = Nothing
                    rs2.Open "select otfhrs from vwovertime where otfid = '" & rsAC("otfid") & "' and empno =  '" & rsAC("empno") & "' ", ac, adOpenStatic, adLockReadOnly
                   '
                    If Not rs2.EOF Then
                        x_otfhr = rs2("otfhrs")
                   
                        If x_otfhr >= 2 And x_otfhr <= 9 Then
                        x_otfsub = 40
                        
                        'MsgBox x_otfsub
                        lstitem.SubItems(2) = x_otfsub
                        ElseIf x_otfhr >= 10 And x_otfhr <= 17 Then
                        x_otfsub = 80
                        lstitem.SubItems(2) = x_otfsub
                        ElseIf x_otfhr >= 18 And x_otfhr <= 24 Then
                        x_otfsub = 120
                        lstitem.SubItems(2) = x_otfsub
                        Else
                         x_otfsub = 0
                        lstitem.SubItems(2) = 0
                        
                        
                        End If
                        
                   x_otcounter = x_otcounter + x_otfsub
                    
                    End If
                                
    
                rsAC.MoveNext
                 
                Loop
            Set rsAC = Nothing
        
                     
            xsubsidy = ListView5.SelectedItem.SubItems(1)
           xdeduct = (xsubtotal_ctr - (xsubsidy + x_otcounter))
           
        
            
           lstitem_1.SubItems(4) = FormatNumber(x_otcounter)
              lstitem_1.SubItems(5) = FormatNumber(xdeduct)
       
        I = I + 1
    Set rs1 = Nothing
  Loop
End Sub
Private Sub reflist5()
ListView5.ListItems.Clear

rsAC.Open "select idno, sum(subsidy) as xsubsidy from tbl_transaction where transdate between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "' group by idno order by idno", ac, adOpenStatic, adLockReadOnly
   
       If rsAC.RecordCount >= 1 Then
            ListView5.ListItems.Clear
                Do While Not rsAC.EOF
                    Set lstitem = ListView5.ListItems.Add(, , rsAC.Fields!idno)
                       
                             lstitem.SubItems(1) = FormatNumber(rsAC("xsubsidy"))
                          
                             
                             
                          
                  
                    rsAC.MoveNext
                    
                Loop
                
                
            Else
      
                ListView5.ListItems.Clear
                Set rsAC = Nothing
   
            End If
            
            Set rsAC = Nothing
End Sub


Private Sub reflist()
Dim xsub As Currency
xsub = 0

'On Error Resume Next
rsAC.Open "select tbl_transaction.transno, itemcode,qty,unitcode,price,subtotal,subsidy,transdate,idno from tbl_transdetails, tbl_transaction where tbl_transdetails.transno = tbl_transaction.transno and transdate between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "' order by transno", ac, adOpenStatic, adLockReadOnly
   
       If rsAC.RecordCount >= 1 Then
            ListView1.ListItems.Clear
                Do While Not rsAC.EOF
                    Set lstitem = ListView1.ListItems.Add(, , rsAC.Fields!TransNo)
                             lstitem.SubItems(1) = rsAC.Fields!idno
                            Set rs1 = Nothing
                            rs1.Open "select * from vwemployeemaster where empno = '" & rsAC.Fields!idno & "'", ac, adOpenStatic, adLockReadOnly
                            If Not rs1.EOF Then
                             lstitem.SubItems(2) = rs1.Fields!emplname + ", " + rs1.Fields!empfname + " " + rs1.Fields!empMname
                            End If
                            Set rs1 = Nothing
                            
                             lstitem.SubItems(3) = rsAC.Fields!itemcode
                            
                            rs1.Open "select itemname from tbl_inventory where recno = '" & rsAC.Fields!itemcode & "'", ac, adOpenStatic, adLockReadOnly
                            If Not rs1.EOF Then
                             lstitem.SubItems(4) = rs1.Fields!itemname
                            End If
                            Set rs1 = Nothing
                             
                             
                             
                             lstitem.SubItems(5) = rsAC.Fields!qty
                             lstitem.SubItems(6) = rsAC.Fields!unitcode
                             lstitem.SubItems(7) = FormatNumber(rsAC.Fields!price)
                             lstitem.SubItems(8) = FormatNumber(rsAC.Fields!subtotal)
                             lstitem.SubItems(9) = FormatNumber(rsAC.Fields!subsidy)
                             lstitem.SubItems(10) = rsAC.Fields!transdate
                             
                          
                     xsub = xsub + rsAC("subsidy")
                     
                    rsAC.MoveNext
                    
                Loop
                Label8.Caption = ListView1.ListItems.Count
                Label9.Caption = FormatNumber(xsub)
                
                
            Else
            MsgBox "No record found.", vbExclamation
                ListView1.ListItems.Clear
                   Label8.Caption = ListView1.ListItems.Count
               Label9.Caption = "0"
               
               
                Set rsAC = Nothing
            End If
            
            Set rsAC = Nothing
End Sub

Private Sub reflist2()
'On Error Resume Next
rsAC.Open "select transno, idno, subsidy from tbl_transaction where transdate between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "' order by transno", ac, adOpenStatic, adLockReadOnly
   
       If rsAC.RecordCount >= 1 Then
            ListView3.ListItems.Clear
                Do While Not rsAC.EOF
                    Set lstitem = ListView3.ListItems.Add(, , rsAC.Fields!TransNo)
                       
                             lstitem.SubItems(1) = rsAC.Fields!idno
                             lstitem.SubItems(2) = FormatNumber(rsAC.Fields!subsidy)
                             
                             
                          
                  
                    rsAC.MoveNext
                    
                Loop
                
                
            Else
      
                ListView3.ListItems.Clear
                Set rsAC = Nothing
   
            End If
            
            Set rsAC = Nothing
End Sub

 _
 _

