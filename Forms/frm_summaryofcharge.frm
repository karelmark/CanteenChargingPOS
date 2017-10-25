VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_summaryofcharge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Summary of Charges"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   15735
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Payroll Date"
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   15615
      Begin VB.ComboBox cmb_pdates 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton btn_Search 
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lv_soc 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   10186
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
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Prev. Balance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Charges"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Paid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Remaining"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "Grand Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   9000
      Width           =   3375
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      Top             =   9000
      Width           =   3375
   End
   Begin VB.Label Label6 
      Caption         =   "Previous Bal."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   7680
      Width           =   3375
   End
   Begin VB.Label lblprevbal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   7680
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Forwarded Bal."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   11
      Top             =   8160
      Width           =   2535
   End
   Begin VB.Label lblfbal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   12120
      TabIndex        =   10
      Top             =   8160
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Total Charges:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   8400
      Width           =   3375
   End
   Begin VB.Label lbltotalcharges 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   8400
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Summary of Charges"
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
      TabIndex        =   5
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblrcvables 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12120
      TabIndex        =   4
      Top             =   7560
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Total Payment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   3
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004000&
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   15975
   End
End
Attribute VB_Name = "frm_summaryofcharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Search_Click()
    If cmb_pdates.Text <> "" Then
    lv_soc.ListItems.Clear
    Dim sql As String
    
    sql = "SELECT emplname,empfname,total_charges,actual_payment , current_rem_bal ,prev_unpaid_charges ,charges ,actual_payment FROM tbl_logs as l INNER JOIN vwEmployeeMaster as e ON l.empid = e.empno WHERE void = 0 and payrolldate = #" & DateValue(cmb_pdates.Text) & "# ORDER BY e.mpayarea, e.empno Asc"
    
    rsoff.Open sql, acaccess
    
    
    Dim totalsubsidy As Double
    Dim actualdeduction As Double
    
    
    Dim prev As Double
    Dim charge As Double
    Dim total As Double
    Dim paid As Double
    Dim remaining As Double
    
    Dim t_prev As Double
    Dim t_charge As Double
    Dim t_total As Double
    Dim t_paid As Double
    Dim t_remaining As Double
    
    ''
    
    Dim prev2 As Double
     
    
   ' Dim charges As Double, pbal As Variant, rbal As Double, paid As Double, cbal As Double
   
    If Not rsoff.EOF Then
    
       t_prev = 0#
       t_charge = 0#
       t_total = 0#
       t_paid = 0#
       t_remaining = 0#
        
         Do While Not rsoff.EOF
                  
                Dim empname As String
                empname = rsoff.Fields!emplname & " , " & rsoff.Fields!empfname
                
                Dim totalcharge As Currency
                Dim actualpayment As Currency
                Dim currentbalnce As Currency
              
                totalcharge = rsoff.Fields!total_charges
                If IsNull(rsoff.Fields!actual_payment) Then
                
                actualpayment = 0
                
                Else
                  
                  actualpayment = rsoff.Fields!actual_payment
                  
                End If
                
                If IsNull(rsoff.Fields!current_rem_bal) Then
                    
                    currentbalnce = 0
                    
                Else
                
                currentbalnce = rsoff.Fields!current_rem_bal
                
                End If
                
                Dim lstitem As ListItem
                
                
                Set lstitem = lv_soc.ListItems.Add(, , empname)
                
                
                lstitem.SubItems(1) = FormatNumber(rsoff.Fields!prev_unpaid_charges, 2)
                lstitem.SubItems(2) = FormatNumber(rsoff.Fields!charges, 2)
                lstitem.SubItems(3) = FormatNumber(totalcharge, 2)
                lstitem.SubItems(4) = FormatNumber(actualpayment, 2)
                lstitem.SubItems(5) = FormatNumber(currentbalnce, 2)
                 
                t_prev = t_prev + rsoff.Fields!prev_unpaid_charges
                t_charge = t_charge + rsoff.Fields!charges
                t_total = t_total + totalcharge
                t_paid = t_paid + actualpayment
                t_remaining = t_remaining + currentbalnce
                
            rsoff.MoveNext
        Loop
        
    
    End If
    rsoff.Close
    Set rsoff = Nothing
    lblprevbal = FormatNumber(t_prev, 2)
    lbltotal.Caption = FormatNumber(t_total, 2)
    lblrcvables = FormatNumber(t_paid, 2)
    lbltotalcharges = FormatNumber(t_charge, 2)
    lblfbal = FormatNumber(t_remaining, 2)
    Else
        MsgBox "Please Select Payroll Date"
    End If
    
End Sub

Private Sub Form_Load()
connectData
Dim stringsql As String
stringsql = "Select Distinct payrolldate From tbl_logs WHERE void = 0 ORDER BY payrolldate Desc"
 
rsoff.Open stringsql, acaccess1
If Not rsoff.EOF Then
            Do While Not rsoff.EOF
             
              cmb_pdates.AddItem (rsoff.Fields!payrolldate)
              
            rsoff.MoveNext
            Loop
    Else
    MsgBox "No record found.", vbExclamation
    End If
rsoff.Close
Set rsoff = Nothing
 Dim w As Double
 w = lv_soc.Width

 lv_soc.ColumnHeaders.Item(1).Width = w * 0.25
 lv_soc.ColumnHeaders.Item(2).Width = w * 0.14
 lv_soc.ColumnHeaders.Item(3).Width = w * 0.14
 lv_soc.ColumnHeaders.Item(4).Width = w * 0.14
 lv_soc.ColumnHeaders.Item(5).Width = w * 0.14
 lv_soc.ColumnHeaders.Item(6).Width = w * 0.15
 
 





End Sub
Private Sub connectData()

 Set acaccess = New ADODB.Connection
 Set acdet = New ADODB.Connection
 Set acaccess1 = New ADODB.Connection
 
 
  If accSer.State = adStateOpen Then
            accSer.Close
  End If
  If acaccess.State = adStateOpen Then
            acaccess.Close
  End If
  If acaccess1.State = adStateOpen Then
            acaccess1.Close
  End If
     
    If acaccess.State = 1 Then acaccess.eClose
        If acaccess1.State = 1 Then acaccess1.eClose
            If acdet.State = 1 Then acdet.eClose
         
          
         acaccess.Open "DSN=canteen_offline"   ' MSACCESS
         acaccess1.Open "DSN=canteen_offline"
         'acdet.Open "DSN=canteen_offline"   ' MSACCESS
         
    dec

    Set rsaccess = New ADODB.Recordset
    Set rsaccDet = New ADODB.Recordset
    Set rsup = New ADODB.Recordset
    Set rbu = New ADODB.Recordset
    Set rtx = New ADODB.Recordset
End Sub
  

