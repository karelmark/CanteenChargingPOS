VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_uploadacess 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recover Unrecorded Sales"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   14130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame CoverUpdate 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   14175
      Begin MSComctlLib.ListView ListView1 
         Height          =   6135
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   10821
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   7095
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   13815
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Update Server"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6000
         Width           =   2535
      End
      Begin MSComctlLib.ListView lv_tx_offline 
         Height          =   4215
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7435
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TX No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lv_tx_details 
         Height          =   4215
         Left            =   7200
         TabIndex        =   4
         Top             =   840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7435
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   735
         Left            =   11160
         TabIndex        =   13
         Top             =   6240
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   13920
         X2              =   0
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "Total Sales:"
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
         Left            =   720
         TabIndex        =   10
         Top             =   5280
         Width           =   2055
      End
      Begin VB.Label lbltotalsales 
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
         Left            =   2880
         TabIndex        =   9
         Top             =   5280
         Width           =   3855
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "SubTotal:"
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
         Left            =   8160
         TabIndex        =   8
         Top             =   5280
         Width           =   2055
      End
      Begin VB.Label Label6 
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
         Left            =   9960
         TabIndex        =   7
         Top             =   5280
         Width           =   3855
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         Caption         =   "Transaction"
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
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Details"
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
         Left            =   7200
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Offline Charges"
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004000&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15975
   End
End
Attribute VB_Name = "frm_uploadacess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub reflist()

ListView1.ListItems.Clear


On Error GoTo X:

rsaccess.Open "SELECT  * FROM tbl_transaction WHERE uploadstatus = 0 ", acaccess, adOpenStatic
   
       Dim totalsales As Double
       totalsales = 0
       
       If rsaccess.RecordCount >= 1 Then
        
       lv_tx_offline.ListItems.Clear
       Dim w As Double
       
       w = lv_tx_offline.Width
       
       lv_tx_offline.ColumnHeaders.Item(1).Width = 0.1 * w
       lv_tx_offline.ColumnHeaders.Item(2).Width = 0.2 * w
       lv_tx_offline.ColumnHeaders.Item(3).Width = 0.3 * w
       lv_tx_offline.ColumnHeaders.Item(4).Width = 0.2 * w
       lv_tx_offline.ColumnHeaders.Item(5).Width = 0.2 * w
       
    Dim lstitem As ListItem
    
    Do While Not rsaccess.EOF
                    
                 Set lstitem = lv_tx_offline.ListItems.Add(, , rsaccess.Fields!transno)
                             
                             lstitem.SubItems(1) = rsaccess.Fields!idno
                             lstitem.SubItems(2) = rsaccess.Fields!transdate
                             lstitem.SubItems(3) = rsaccess.Fields!transtime
                             lstitem.SubItems(4) = Format(CDbl(rsaccess.Fields!txtotal), "#,##0.00")
                             
                             totalsales = totalsales + CDbl(rsaccess.Fields!txtotal)
                             
                             rsaccess.MoveNext
                    
                    
                Loop
            Else
            
                lv_tx_offline.ListItems.Clear
                Set rsaccess = Nothing
                
            End If
            
    lbltotalsales.Caption = "P " + Format(totalsales, "#,##0.00")
    
    
    If rsaccess.State = adStateOpen Then
            
            rsaccess.Close
            
    End If
    
    Exit Sub
            
X:

  Set rsAC = Nothing
  MsgBox err.Description
  
End Sub
  

Private Sub Command1_Click()

On Error GoTo err:

Dim pbar As Double
pbar = 0

If vbYes = MsgBox("Are you sure about this?", vbCritical + vbYesNo, "Please Confirm!") Then

 Command1.Enabled = False
 CoverUpdate.Visible = True
  
 
        DoEvents
        DoEvents
        DoEvents
        
 Dim X As String
 X = lv_tx_offline.ListItems.Count
 
 Dim i As String
 i = 1
 
           
Set rtx = Nothing

rtx.Open "SELECT * FROM tbl_transaction WHERE uploadstatus = 0 ", acaccess, adOpenDynamic, adLockOptimistic
  
ListView1.ListItems.Add , , "Uploading ..."


While (Not rtx.EOF)
        
        DoEvents
        DoEvents
        DoEvents
        
        Set rsup = Nothing
        
        rsup.Open "SELECT * FROM tbl_transaction", accSer, adOpenDynamic, adLockOptimistic
        rsup.AddNew
            
            rsup.Fields!transdate = rtx.Fields!transdate
            rsup.Fields!transtime = rtx.Fields!transtime
            rsup.Fields!incharge = rtx.Fields!incharge
            rsup.Fields!cardno = rtx.Fields!cardno
            rsup.Fields!idno = rtx.Fields!idno
            rsup.Fields!status = rtx.Fields!status
            rsup.Fields!remarks = rtx.Fields!remarks
            rsup.Fields!subsidy = 0
            rsup.Fields!txtotal = rtx.Fields!txtotal
            
        rsup.Update
        
        ListView1.ListItems.Add , , rtx.Fields!idno & " " & rtx.Fields!txtotal & " -- DONE --"
        
        
        rsup.Close
        
        Dim txcode  As String
        txcode = rtx.Fields!transno
          
        Dim details As Variant
        
        rbu.Open "SELECT * FROM tbl_transdetails WHERE transno = '" + txcode + "'", acaccess1, adOpenDynamic, adLockOptimistic
       
        Do While Not rbu.EOF
          
               Set RS = Nothing
               RS.Open "SELECT Top 1 transno  from tbl_transaction ORDER BY transno DESC", accSer, adOpenDynamic, adLockOptimistic
               
               Dim stat As String
               stat = "OK"
               
               If (rbu.status <> 0) Then
                    stat = "CANCELLED"
               Else
                    stat = "OK"
               End If
               
               
               
               Dim sql As String
               'sql = "INSERT INTO tbl_transdetails(recno, transno,itemno,itemcode,qty,unitcode,Price,subtotal,status) VALUES( '12" & RS.Fields!transno & "', '" & RS.Fields!transno & "', '" & rbu.Fields!itemno & "','" & rbu.Fields!itemcode & "'," & rbu.Fields!qty & ",'" & rbu.Fields!unitcode & "'," & rbu.Fields!Price & "," & rbu.Fields!subtotal & ",'" & rbu.status & "')"
               sql = "INSERT INTO tbl_transdetails( transno,itemno,itemcode,qty,unitcode,Price,subtotal,status) VALUES(   '" & RS.Fields!transno & "', '" & rbu.Fields!itemno & "','" & rbu.Fields!itemcode & "'," & rbu.Fields!qty & ",'" & rbu.Fields!unitcode & "'," & rbu.Fields!Price & "," & rbu.Fields!subtotal & ",'" & stat & "')"
               
               accSer.Execute sql
               
               rbu.MoveNext
               
          Loop
         
        rbu.Close
        
        i = i + 1

        acaccess.Execute "UPDATE tbl_transdetails SET uploaded = 1 WHERE transno ='" + txcode + "'"
        acaccess.Execute "UPDATE tbl_transaction SET uploadstatus = 1 WHERE transno =" + txcode + ""
          
          
        reflist
        
        rtx.MoveNext
'
'        pbar = i / CInt(rtx.Fields.Count)
'        pbar = pbar * 100
'
'        prgress = Format$(pbar, "#0.#")
'
'
'        lblprogress.Caption = prgress
'        ProgressBar1.Value = prgress
'
        DoEvents
        DoEvents
        DoEvents
        
Wend
  
rtx.Close

Set rtx = Nothing
 
'ProgressBar1.Value = ProgressBar1.Max
 
CoverUpdate.Visible = False

MsgBox "Upload Finish! ", vbOKCancel, "Done Updating the server ...!"

lv_tx_details.ListItems.Clear
Command1.Enabled = True

  
End If

  Exit Sub

err:

MsgBox err.Description
Command1.Enabled = True
CoverUpdate.Visible = False
  
End Sub
 

Private Sub Form_activate()
CoverUpdate.Visible = False
Call reflist
End Sub

 

Private Sub lv_tx_offline_Click()

On Error GoTo X:

    rsaccDet.Open "SELECT  * FROM tbl_transdetails WHERE transno = '" + CStr(lv_tx_offline.SelectedItem.Text) + "'", acaccess, adOpenStatic
    
    lv_tx_details.ListItems.Clear
       Dim totaldesales As Double
       Dim detotalsales As Double
       Dim lstitem As ListItem
       
      detotalsales = 0
       
         If rsaccDet.RecordCount >= 1 Then
       
                Do While Not rsaccDet.EOF
                    
                    rsaccess.Open "Select itemname FROM tbl_inventory WHERE recno = " + rsaccDet.Fields!itemcode + "", acaccess, adOpenDynamic
                    
                    Set lstitem = lv_tx_details.ListItems.Add(, , rsaccDet.Fields!itemno)
                             
                             lstitem.SubItems(1) = rsaccess.Fields!itemname
                             lstitem.SubItems(2) = rsaccDet.Fields!qty
                             lstitem.SubItems(3) = Format(CDbl(rsaccDet.Fields!Price), "#,##0.00")
                             lstitem.SubItems(4) = Format(CDbl(rsaccDet.Fields!subtotal), "#,##0.00")
                             
                            detotalsales = detotalsales + CDbl(rsaccDet.Fields!subtotal)
                             
                             
                    rsaccDet.MoveNext
                    rsaccess.Close
                    
                Loop
                
            Else
                
                Set rsaccDet = Nothing
                
            End If
       Label6.Caption = "P " + Format(detotalsales, "#,##0.00")
    
     If rsaccDet.State = adStateOpen Then
      rsaccDet.Close
     End If
    
     Set rsaccDet = Nothing
      
    Exit Sub
X:
 
  MsgBox err.Description
  
End Sub

Private Sub Form_Load()

Call connectData
Call reflist


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
         
         accSer.Open "DSN=update_ccs;UID=pip_connect;PWD=pipconnect"   ' SQL
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
  

 
