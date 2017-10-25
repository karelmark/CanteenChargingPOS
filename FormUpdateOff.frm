VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_updateoff 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UPDATE TABLES"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "FormUpdateOff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Download 
      BackColor       =   &H00FFFFFF&
      Caption         =   "UPDATE"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin MSComctlLib.ListView lvlogs 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   12303
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Doing ..."
         Object.Width           =   7011
      EndProperty
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE  DATA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frm_updateoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Download_Click()
Download.Enabled = False
        Dim lstitem As ListItem
        Dim sql As String
        
If MsgBox("Are you sure you want to do this ?", vbCritical + vbYesNo, "Please Confirm ...") = vbYes Then
    
    Dim maxlogs As Integer
    maxlogs = 0
             
             
sql = "SELECT *  FROM tbl_logs WHERE offline_copy = 0 Order by recno"
rsup.Open sql, accSer, adOpenStatic, adLockReadOnly
 
Set lstitem = lvlogs.ListItems.Add(, , rsup.RecordCount & " records to update...")
maxlogs = rsup.RecordCount
 
Do While Not rsup.EOF

'UPDATE table offline
   Set rs2 = Nothing
   
        DoEvents
        DoEvents
        
        rs2.Open "select * from tbl_logs", acaccess, adOpenDynamic, adLockOptimistic
        rs2.AddNew
  
        rs2.Fields!transcode = rsup.Fields!transcode
        rs2.Fields!recdate = rsup.Fields!recdate
        rs2.Fields!rectime = rsup.Fields!rectime
        rs2.Fields!EmpID = rsup.Fields!EmpID
        rs2.Fields!incharge = rsup.Fields!incharge
        rs2.Fields!payrolldate = rsup.Fields!payrolldate
        rs2.Fields!cutoffstartdate = rsup.Fields!cutoffstartdate
        rs2.Fields!cutoffdate = rsup.Fields!cutoffdate
        rs2.Fields!prev_bal = rsup.Fields!prev_bal
        rs2.Fields!prev_unpaid_charges = rsup.Fields!prev_unpaid_charges
        rs2.Fields!total_charges = rsup.Fields!total_charges
        rs2.Fields!subtotal = rsup.Fields!subtotal
        rs2.Fields!payrolldeduction = rsup.Fields!payrolldeduction
        rs2.Fields!actualdeduction = rsup.Fields!actualdeduction
        rs2.Fields!deduction_adjustment = rsup.Fields!deduction_adjustment
        rs2.Fields!empbalance = rsup.Fields!empbalance
        rs2.Fields!actual_payment = rsup.Fields!actual_payment
        rs2.Fields!current_rem_bal = rsup.Fields!current_rem_bal
        rs2.Fields!payable = rsup.Fields!payable
        rs2.Fields!payments = rsup.Fields!payments
        rs2.Fields!balance = rsup.Fields!balance
        rs2.Fields!paymentdate = rsup.Fields!paymentdate
        rs2.Fields!remarks = rsup.Fields!remarks
        rs2.Fields!void = rsup.Fields!void
        
        rs2.Update
        
        DoEvents
        DoEvents
        
        
        rs2.Close
        
        DoEvents
        DoEvents
         
                   
rsup.MoveNext
                
Loop

 
accSer.Execute "Update tbl_logs SET offline_copy  = 1 WHERE offline_copy = 0 "

Set lstitem = lvlogs.ListItems.Add(, , "Done Updating Cut-off " & rsup.RecordCount & " new records ")
             
rsup.Close
Set rsup = Nothing


    'TO UPDATE tables offline
    'Employee
    'Transaction ?
Set lstitem = lvlogs.ListItems.Add(, , "Checking New Employee's")
sql = "SELECT *  FROM vwEmployeeMaster"
rsup.Open sql, accSer, adOpenStatic, adLockReadOnly


Do While Not rsup.EOF

'Check if exist offline
Set rtx = Nothing
     rtx.Open "SELECT Count(*) as cnt FROM vwEmployeeMaster WHERE empNo = '" & rsup.Fields!empno & "'", acaccess, adOpenDynamic, adLockOptimistic
     
    If (rtx.Fields!cnt = 0) Then
        'Update Offline
        Set rs2 = Nothing
        DoEvents
        DoEvents
        
        rs2.Open "select * from vwEmployeeMaster", acaccess, adOpenDynamic, adLockOptimistic
        rs2.AddNew
  
        rs2.Fields!empno = rsup.Fields!empno
        rs2.Fields!emplname = rsup.Fields!emplname
        rs2.Fields!empfname = rsup.Fields!empfname
        rs2.Fields!EmpMname = rsup.Fields!EmpMname
        rs2.Fields!Pwd = rsup.Fields!Pwd
        rs2.Fields!EmpEmail = rsup.Fields!EmpEmail
        rs2.Fields!MpayArea = rsup.Fields!MpayArea
        rs2.Fields!EmpCat = rsup.Fields!EmpCat
         
        rs2.Update
        
        DoEvents
        DoEvents
        
        rs2.Close
        
        DoEvents
        DoEvents
        
        Set lstitem = lvlogs.ListItems.Add(, , "Added " & rsup.Fields!empno)
      
      End If
      
rsup.MoveNext
                
Loop
 
Set lstitem = lvlogs.ListItems.Add(, , "Done Checking/Updating Employee Data")

             
rsup.Close
Set rsup = Nothing
rtx.Close
Set rtx = Nothing

  
    


     
End If

Download.Enabled = True
        
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
         
             accSer.Open "DSN=update_ccs;UID=ccs_connect;PWD=ccs"   ' SQL
         
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

Private Sub Form_Load()
 Call connectData
End Sub

