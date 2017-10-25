VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   Icon            =   "FrmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   9240
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   435
      Left            =   1920
      TabIndex        =   14
      Top             =   1920
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      Format          =   73334785
      CurrentDate     =   37987
      MinDate         =   36526
   End
   Begin VB.ComboBox payrolldate 
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
      ItemData        =   "FrmPrint.frx":030A
      Left            =   1920
      List            =   "FrmPrint.frx":030C
      TabIndex        =   13
      Top             =   1440
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4920
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "FrmPrint.frx":030E
      Left            =   4365
      List            =   "FrmPrint.frx":03CF
      TabIndex        =   11
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "FrmPrint.frx":054D
      Left            =   2055
      List            =   "FrmPrint.frx":0575
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   435
      Left            =   4320
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      Format          =   73334785
      CurrentDate     =   35796
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Payroll Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Monthly"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Between"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2475
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Today"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   4800
         Picture         =   "FrmPrint.frx":05DB
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1800
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "FrmPrint.frx":08E5
         Left            =   120
         List            =   "FrmPrint.frx":08FE
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   315
         Width           =   4455
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   435
      Left            =   1920
      TabIndex        =   1
      Top             =   2400
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      Format          =   73334785
      CurrentDate     =   38353
   End
   Begin VB.Label Label3 
      Caption         =   "Yr :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4050
      TabIndex        =   8
      Top             =   4455
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      TabIndex        =   0
      Top             =   2520
      Width           =   375
   End
End
Attribute VB_Name = "FrmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Command1_Click()
 On Error GoTo errhandler:
If Not Combo1.Text = "" Then
                      
                    If Combo1.Text = "Sales" Then
                        
                                  If Option1.Value = True Then
                                                de1.conn1.Open
                                                 de1.Sales_daily_Grouping (DTPicker3.Value)
                                                 
                                                With RptSales_daily
                                                
                                                        With .Sections("Section4").Controls
                                                             .Item(3).Caption = DTPicker3.Value
                                                          
                                                        End With
                                                .Show 1
                                                End With
                                                de1.rsSales_daily_Grouping.Close
                                                de1.conn1.Close
                                  ElseIf Option2.Value = True Then
                                                    
                                                    de1.conn1.Open
                                                    de1.Sales_between_Grouping DTPicker1.Value, DTPicker2.Value
                                                    
                                                    With RptSales_between
                                                    
                                                            With .Sections("Section4").Controls
                                                                 .Item(3).Caption = DTPicker1.Value
                                                                .Item(8).Caption = DTPicker2.Value
                                                              
                                                            End With
                                                    .Show 1
                                                    End With
                                                    de1.rsSales_between_Grouping.Close
                                                    de1.conn1.Close
                                                                      
                                  End If
                     ElseIf Combo1.Text = "Inventory" Then
                     de1.conn1.Open
                                If Option1.Value = True Then
                                                de1.Inventory_Day_Grouping DTPicker3.Value
                                                With RptInventory_day
                                                
                                                        With .Sections("Section4").Controls
                                                             .Item(7).Caption = DTPicker3.Value
                                                          
                                                        End With
                                                .Show 1
                                                End With
                                                de1.rsInventory_Day_Grouping.Close
                                  
                                  ElseIf Option2.Value = True Then
                                                    
                                              Call dNewinventory(DTPicker1.Value, DTPicker2.Value)
                    
                                                                    
                                  End If
                         de1.conn1.Close
                    ElseIf Combo1.Text = "Cutoff History" Then
                     
                     Call cutoffhistory
                     
                     
                    
                    ElseIf Combo1.Text = "Summary of Charges" Then
                         If Option1.Value Then
                                summ_charges DTPicker3.Value, DTPicker3.Value
                          ElseIf Option2.Value Then
                                summ_charges DTPicker1.Value, DTPicker2.Value
                          Else
                                summ_charges payrolldate.Text, payrolldate.Text
                          End If
                          
                    ElseIf Combo1.Text = "Billing Statement" Then
                    
                                Dim dt1 As Date
                                Dim dt2 As Date
                                
                                If Option1.Value Then
                                    dt1 = DTPicker3.Value
                                    dt2 = DTPicker3.Value
                                ElseIf Option2.Value Then
                                    dt1 = DTPicker1.Value
                                    dt2 = DTPicker2.Value
                                End If
                                    reportbill dt1, dt2
                    End If
 
            Else
                    MsgBox "Please select what to print.", vbExclamation
                    Exit Sub
            End If

            Exit Sub
errhandler:
            MsgBox err.Description
            Exit Sub
End Sub

Private Sub summ_charges(ByVal dt1 As Date, ByVal dt2 As Date)

Dim SQL As String

'SQL = " SELECT  EmpNo as ID, EmpLName + ' ,' + EmpFName AS Fullname , B.cutoffstartdate as startdate , B.cutoffdate as enddate,  " & _
'      " case WHEN isnull(B.prev_bal,0) > 0 then isnull(B.prev_bal,0) else 0 end as prevbal, B.charges as charges, b.charges + case WHEN isnull(B.prev_bal,0) > 0 then isnull(B.prev_bal,0) else 0 end  as payable, B.payments as payments, case WHEN isnull(B.prev_bal,0) > 0 then isnull(B.prev_bal,0) else 0 end + B.charges - B.payments as subtotal " & _
'      " FROM    vwEmployeeMaster AS A INNER JOIN tbl_logs as B ON A.empno = B.empid " & _
'      " WHERE A.Mpayarea <> '99' AND B.void = 0 AND B.payrolldate = '" & payrolldate.Text & "' " & _
'      " ORDER BY EmpNo, transcode "


SQL = " SELECT  EmpNo as ID, EmpLName + ' ,' + EmpFName AS Fullname , B.cutoffstartdate as startdate , B.cutoffdate as enddate,  " & _
      " isnull(B.prev_bal2,0)  as prevbal, B.charges as charges, b.charges + isnull(B.prev_bal2,0)  as payable, B.payments as payments, isnull(B.prev_bal2,0) + B.charges - B.payments as subtotal " & _
      " FROM    vwEmployeeMaster AS A INNER JOIN tbl_logs as B ON A.empno = B.empid " & _
      " WHERE A.Mpayarea <> '99'  AND B.void = 0 AND B.payrolldate = '" & payrolldate.Text & "' " & _
      " ORDER BY  Mpayarea, EmpNo, transcode "
 
            
rsPin.Open SQL, ac, adOpenStatic, adLockReadOnly
With rptcharge
    Set .DataSource = Nothing
        .DataMember = ""
    Set .DataSource = rsPin.DataSource
    With .Sections("Section4")
    
        .Controls("label12").Caption = Format(CDate(rsPin.Fields("startdate")), "mmmm dd, yyyy")
        .Controls("label15").Caption = Format(CDate(rsPin.Fields("enddate")), "mmmm dd, yyyy")
        .Controls("payrolldate").Caption = "Payroll Date: " & Format(CDate(dt1), "mmmm dd, yyyy")
    
    End With
End With

rptcharge.Show 1
rsPin.Close

                                           

End Sub

Private Sub reportbill(ByVal dt1 As Date, ByVal dt2 As Date)
On Error GoTo er

Dim dSQL As String
Dim rsbill As New ADODB.Recordset
dSQL = "SELECT T1.Txtotal, T2.PrevTotal AS dprevbal,T2.PrevTotal AS PrevTotal, T1.Txtotal + T2.PrevTotal AS GrandTotal FROM (SELECT SUM(txtotal) AS Txtotal  FROM  tbl_transaction  WHERE  (transdate BETWEEN  '" & dt1 & "' AND '" & dt2 & "')) AS T1 CROSS JOIN (SELECT     TOP (1) SUM(balance) AS PrevTotal  FROM tbl_logs  WHERE  (transcode =   (SELECT     TOP (1) transcode   FROM  tbl_logs AS tbl_logs_1  WHERE   (void = 0) AND (paymentdate IS NOT NULL)  ORDER BY transcode DESC))) AS T2"
rsbill.Open dSQL, ac, adOpenStatic

With rptBilling
Set .DataSource = Nothing
    .DataMember = ""
Set .DataSource = rsbill.DataSource
     
    With .Sections("Section2")
    
    .Controls("label7").Caption = "To: Pilipinas Kao Inc."
    .Controls("label8").Caption = "From: Gracia's Foodhaus and Bakeshop"
    .Controls("label2").Caption = "RE: Billing period of " & Format(CDate(dt1), "mmmm dd, yyyy") & " to " & Format(CDate(dt2), "mmmm dd, yyyy")
    
    End With
    
    With .Sections("Section1")
     .Controls("label3").Caption = "This is to bill you the amount of P"
'     .Controls("label10").Caption = "for the period of " & _
'          Format(CDate(dt1), "mmmm dd, yyyy") & " to " & Format(CDate(dt2), "mmmm dd, yyyy") & " for the canteen services to PKI."
    .Controls("label10").Caption = " For the canteen services to PKI."
    .Controls("drange").Caption = "Billing period of " & Format(CDate(dt1), "mmmm dd, yyyy") & " to " & Format(CDate(dt2), "mmmm dd, yyyy")
'    .Controls("lbltxtotal").Caption = Format(Val(gxtotal), "#,###0.00")
'    .Controls("lblprevtotal").Caption = Format(Val(gptotal), "#,##0.00")
'    .Controls("lblgtotal").Caption = Format(Val(grandtotal), "#,##0.00")
'    .Controls("gtotal").Caption = Format(Val(grandtotal), "#,##0.00")
'
    
    End With
    With .Sections("Section5")
        .Controls("label5").Caption = "Neil A. Adran"
    End With
   
End With
rptBilling.Show 1

rsbill.Close
Exit Sub

er:
  MsgBox err.Description

End Sub
Private Sub cutoffhistory()
On Error GoTo er

Dim dSQL As String
Dim rshistory As New ADODB.Recordset
dSQL = "SELECT transcode, cutoffstartdate, cutoffdate, SUM(balance) AS Balance, SUM(charges) AS Charges, SUM(payable) AS Payable, SUM(payments) AS Payment From tbl_logs Where (void = 0) GROUP BY transcode, cutoffstartdate, cutoffdate ORDER BY transcode DESC"

rshistory.Open dSQL, ac, adOpenStatic

With rptcutofflogs
Set .DataSource = Nothing
    .DataMember = ""
Set .DataSource = rshistory.DataSource
     
    With .Sections("Section2")
    
    ' .Controls("label7").Caption = "To: Pilipinas Kao Inc."
    ' .Controls("label8").Caption = "From: Gracia's Foodhaus and Bakeshop"
    ' .Controls("label2").Caption = "RE: Billing period of " & Format(CDate(dt1), "mmmm dd, yyyy") & " to " & Format(CDate(dt2), "mmmm dd, yyyy")
    
    End With
    
    With .Sections("Section1")
    ' .Controls("label3").Caption = "This is to bill you the amount of P"
 
    
    End With
    With .Sections("Section5")
        '.Controls("label5").Caption = "Neil A. Adran"
    End With
   
End With
rptcutofflogs.Show 1

rshistory.Close
Exit Sub

er:
  MsgBox err.Description

End Sub

Private Sub dNewinventory(ByVal dt1 As Date, ByVal dt2 As Date)
On Error GoTo er

Dim dSQL As String
Dim rsnewinv As New ADODB.Recordset
dSQL = "SELECT   C.Itemname, A.price, SUM(A.qty) AS qty, SUM(A.subtotal) AS Subtotal FROM  tbl_transdetails AS A INNER JOIN   tbl_transaction AS B ON A.transno = B.transno INNER JOIN Tbl_inventory AS C ON A.itemcode = C.Recno WHERE     (B.transdate  BETWEEN  '" & dt1 & "' AND '" & dt2 & "') AND (A.status = 'OK') GROUP BY  C.Itemname, A.price ORDER BY C.Itemname"

rsnewinv.Open dSQL, ac, adOpenStatic

With rptNewInventory
Set .DataSource = Nothing
    .DataMember = ""
Set .DataSource = rsnewinv.DataSource
     
    With .Sections("Section4")

'    .Controls("label7").Caption = "To: Pilipinas Kao Inc."
'    .Controls("label8").Caption = "From: Gracia's Foodhaus and Bakeshop"
    .Controls("label12").Caption = "Covered Dates From: " & Format(CDate(dt1), "mmmm dd, yyyy") & " to " & Format(CDate(dt2), "mmmm dd, yyyy")

    End With
'
'    With .Sections("Section1")
'    ' .Controls("label3").Caption = "This is to bill you the amount of P"
'
'
'    End With
'    With .Sections("Section5")
'        '.Controls("label5").Caption = "Neil A. Adran"
'    End With
   
End With
rptNewInventory.Show 1

rsnewinv.Close
Exit Sub

er:
  MsgBox err.Description

End Sub

 

Private Sub Form_Load()
On Error GoTo err:
'Option1.Value = True
'Combo4.Text = Format(Date, "YYYY")
DTPicker3.Value = Date
DTPicker1.Value = Date
Label3.Visible = False
DTPicker3.Visible = True
Combo2.Visible = False
Combo3.Visible = False
'Combo4.Visible = False


Label2.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
  
Combo2.Visible = False
Combo3.Visible = False


Dim SQL As String

 
SQL = "SELECT DISTINCT  payrolldate FROM  tbl_logs WHERE void = 0 ORDER BY payrolldate DESC"
rsUtil.Open SQL, ac, adOpenStatic, adLockReadOnly

 
While Not rsUtil.EOF
        payrolldate.AddItem (rsUtil.Fields("payrolldate"))
        rsUtil.MoveNext
Wend
rsUtil.Close
Exit Sub

err:
    rsUtil.Close
    MsgBox err.Description
        
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Option1_Click()
Label3.Visible = False
DTPicker3.Visible = True




    
Combo2.Visible = False
Combo3.Visible = False
payrolldate.Visible = False


Label2.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False




Combo2.Visible = False
Combo3.Visible = False
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
DTPicker3.Visible = False

Label2.Visible = True
DTPicker1.Visible = True
DTPicker2.Visible = True
DTPicker1.Value = Date
DTPicker2.Value = Date

Label3.Visible = False

    Combo2.Visible = False
Combo3.Visible = False
'Combo4.Visible = False
  
End If


End Sub

Private Sub Option3_Click()

If Option3.Value = True Then
DTPicker3.Visible = False
Label3.Visible = True
Combo2.Visible = True

Combo2.Visible = True
Combo3.Visible = True

Combo2.Text = Format(Date, "MMMM")
Combo3.Text = Format(Date, "yyyy")
DTPicker3.Visible = False


Label2.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False


Combo4.Visible = False
    
End If

End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
   
'  Combo4.Visible = True
 DTPicker3.Visible = False
 
    
Label3.Visible = False



Label2.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False

Combo2.Visible = False
Combo3.Visible = False

End If

End Sub

