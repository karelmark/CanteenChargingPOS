Attribute VB_Name = "Module1"
Option Explicit

 'Public poSendMail As vbSendMail.clsSendMail
Public who As Integer
Public ac As New ADODB.Connection

Public acoff As New ADODB.Connection

Public rsLog_in As New ADODB.Recordset
Public rs1  As New ADODB.Recordset
Public rs2  As New ADODB.Recordset
Public rs3  As New ADODB.Recordset
Public rs4  As New ADODB.Recordset
Public rs5  As New ADODB.Recordset
Public rs6  As New ADODB.Recordset
Public rs7  As New ADODB.Recordset
Public rs8  As New ADODB.Recordset
Public rs9  As New ADODB.Recordset
Public rs10 As New ADODB.Recordset
Public RS  As New ADODB.Recordset
Public rsNew As New ADODB.Recordset
Public sl As New ADODB.Recordset
Public wali As String
Public varmonth As String
Public counter As String
Public ControlNo, AdJuStMeNt As String
Public rsPOS  As New ADODB.Recordset
Public rsAC  As New ADODB.Recordset
Public rsACSearch  As New ADODB.Recordset
Public rsPin As New ADODB.Recordset
Public rsUtil As New ADODB.Recordset
Public rsoff As New ADODB.Recordset

Public rsAC1  As New ADODB.Recordset
Public rsAC2  As New ADODB.Recordset
'Public rs1  As New ADODB.Recordset

Public rsAC3  As New ADODB.Recordset
Public rsAC4  As New ADODB.Recordset
Public rsAC5  As New ADODB.Recordset
Public rsAC6  As New ADODB.Recordset
Public rsac7 As New ADODB.Recordset
Public sql, sqlstring As String
Public pureControlNo As String
Public IsProccess As Boolean
Public ProccessLevel As Integer
Public pritem_limit As Long
Public username_h As String
Public pur_encharge As String
Public pr_leadtime As Long
Public cv_leadtime As Long
Public po_leadtime As Long
Public po_app_leadtime As Long
Public po_del_leadtime As Long
Public user_status As String
'Login
Public getusername As String
Public getPurSuvFName  As String
Public getPurSuvLName  As String

Public userlogin As Boolean
Public getUID As Long
Public cgetUID As Long
Public getUCODE As String
Public RcvNo As String
Public fEntryNo As String

'frmfuel
Public Total_fuel As Currency
Public Total_QueryBags As Currency
Public QueryPrnt As String
Public TotalRRbags As Double
Public TotalPayables As Currency
Public TempRValue As Integer
Public TempRBags As Integer
Public TotalRecievableBags As Double
Public maxAmount As Currency
Public gICDatePrev As String
Public gICDateCurr As String
Public GetRcvIC As String
Public gotPop As Boolean

Public gDateStart As Date
Public gPrevDateTemp As Date
Public gPrevDate As Date
Public gDateEnd As Date

Public validpin As Boolean
Public pin_pinempid As String
Public pin_pincode As String

Public offlinestatus As Boolean


'OFFLINE
Public acaccess As New ADODB.Connection
Public acaccess1 As New ADODB.Connection
Public acdet As New ADODB.Connection

Public accSer As New ADODB.Connection

Public rsaccess  As New ADODB.Recordset
Public rsaccDet As New ADODB.Recordset
Public rsup As New ADODB.Recordset
Public rbu As New ADODB.Recordset
Public rtx As New ADODB.Recordset


Type OrientStructure

 
      Orientation As Long
      Pad As String * 16

End Type
   
' Enter the following Declare statement on one, single line:
Declare Function Escape% Lib "GDI" (ByVal hDc%, ByVal nEsc%, ByVal nLen%, lpData As OrientStructure, lpOut As Any)
'PUT THIS SUB IN A .BAS MODULE
Public Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Long, ByVal dwReserved As Long) As Long 'Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long 'this function used with IE4
Private Declare Function GetRTTAndHopCount _
    Lib "iphlpapi.dll" _
   (ByVal lDestIPAddr As Long, _
    ByRef lHopCount As Long, _
    ByVal lMaxHops As Long, _
    ByRef lRTT As Long) As Long
        
Private Declare Function inet_addr _
    Lib "wsock32.dll" _
   (ByVal cp As String) As Long
 
  
Public Sub Main()

offlinestatus = False
Frmpassword.Show

End Sub
 
  
Public Sub dataconnect()

On Error GoTo errbo
Dim dwFlags As Long
Dim sNameBuf As String
Dim lPos As Long
sNameBuf = String$(513, 0)
'If InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&) Then

'connect offline

'If Ping("10.30.10.233") Then
'    Set ac = New ADODB.Connection
'    If ac.State = 1 Then ac.eClose
'
'       If offlinestatus Then
'         ac.Open "DSN=canteen_offline"   ' SQL
'            FrmMainMenu.Offline.Checked = True
'       Else
'       If Ping("10.30.10.233") Then
'         ac.Open "DSN=ccs_connect;UID=ccs_connect;PWD=ccs"   ' SQL
'          FrmMainMenu.Offline.Checked = False
'       Else
'            ac.Open "DSN=canteen_offline"   ' SQL
'            FrmMainMenu.Offline.Checked = Truemark
'
'        End If
'       End If
'        dec
'
'
'
'Else

'MsgBox "Please Check Network Connection!", vbCritical + vbOKOnly, "Network Problem, Working Offline!"
'FrmCredit.Timer2.Enabled = False
'FrmMainMenu.Offline.Checked = True
    Set ac = New ADODB.Connection
    If ac.State = 1 Then ac.eClose
       
        'ac.Open "DSN=ccs_connect;UID=ccs_connect;PWD=ccs"   ' SQL
         ac.Open "DSN=canteen_offline"   ' MSACCESS
    dec
'End If

    Exit Sub

errbo:
        MsgBox err.Description & " " & err.Number
     
End Sub

Public Sub dec()

Set RS = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set rs4 = New ADODB.Recordset
Set rs5 = New ADODB.Recordset
Set rs6 = New ADODB.Recordset
Set rs7 = New ADODB.Recordset
Set rs8 = New ADODB.Recordset
Set rs9 = New ADODB.Recordset
Set rs10 = New ADODB.Recordset

End Sub
 

Public Function tCase(str As String, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Exit Function
    End If
        
    tCase = KeyAscii
                                                                                                                                                                    
End Function

Public Sub ValidNumeric(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case 97
Case 110
Case 47
Case 13
Case 32
Case 48 To 57
 Case Else
  MsgBox "Invalid Input.Please Enter Numeric Types Only..", vbOKOnly + vbExclamation
  KeyAscii = 0
End Select
End Sub

Public Function validTxt(kl As Integer, Txt_Type As Integer)
Dim valStr As String
If Txt_Type = 1 Then
valStr = "0123456789.QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm\[];,./?><{()}!@#$%^&*- "
Else
valStr = "0123456789."
End If
If kl > 26 Then
    If InStr(valStr, Chr(kl)) = 0 Then
       kl = 0
   End If
End If
End Function

Public Sub getPurSuv()
Dim a_rs As ADODB.Recordset
Set a_rs = New ADODB.Recordset
    a_rs.Open "Select fname, lname from tbl_maintenance, tbl_userlogin where username = setvalue and code = 'PURSUV' ", ac
    If Not a_rs.EOF() Then
     getPurSuvFName = a_rs!fname
     getPurSuvLName = a_rs!lname
    End If
End Sub



