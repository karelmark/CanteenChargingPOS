VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form RFID2 
   Caption         =   "Form1"
   ClientHeight    =   11025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   11025
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8760
      TabIndex        =   20
      Text            =   "003"
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   8760
      ScaleHeight     =   1875
      ScaleWidth      =   1755
      TabIndex        =   19
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   8880
      TabIndex        =   18
      Top             =   3120
      Width           =   1575
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1440
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      EOFEnable       =   -1  'True
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   1815
      Left            =   6240
      OleObjectBlob   =   "RFID2.frx":0000
      TabIndex        =   17
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   840
      TabIndex        =   16
      Text            =   "Text4"
      Top             =   10560
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3120
      Top             =   9960
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   10080
      Width           =   2655
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   9360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ClearGlog"
      Height          =   495
      Left            =   7200
      TabIndex        =   14
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   7200
      TabIndex        =   13
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7200
      TabIndex        =   12
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CheckBox chkDownload 
      Caption         =   "Check1"
      Height          =   195
      Left            =   8280
      TabIndex        =   11
      Top             =   5160
      Width           =   1215
   End
   Begin MSComctlLib.ListView listview1 
      Height          =   2175
      Left            =   720
      TabIndex        =   10
      Top             =   7200
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
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
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ComboBox cmbconnect 
      Height          =   315
      ItemData        =   "RFID2.frx":0024
      Left            =   3000
      List            =   "RFID2.frx":002E
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox machineno 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Text            =   "1"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox comport 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "1"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdDownloadLog 
      Caption         =   "Command2"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdDownload01 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   480
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Width           =   6375
   End
   Begin VB.TextBox ItemDatabase 
      Height          =   975
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2040
      Width           =   6375
   End
   Begin VB.Label lblStatus 
      Caption         =   "lblStatus"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   3240
      Width           =   6615
   End
End
Attribute VB_Name = "RFID2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngComPort As Integer
Dim lngMachineNum As Integer
Dim lngBaudRate As Long
Dim bconn As Boolean

Private Sub cmdConnect_Click()
Dim lngComPort As Integer
Dim lngMachineNum As Integer
Dim lngBaudRate As Long
Dim bconn As Boolean
Dim devModel As String
Dim reader As Integer
Dim devNo As Integer
Dim ipAdd As String
Dim portNo As Long

'Me.MousePointer = vbHourglass

If cmdConnect.Caption = "Connect" Then

'DoEvents

lblStatus.Caption = "Connecting Device..."
lngComPort = CLng(Val(comport.Text))
lngMachineNum = CLng(Val(machineno.Text))
lngBaudRate = 115200 'default value of Baudrate
devNo = "1"
ipAdd = "10.30.10.199"
portNo = "4370"

       
   
   
'If cmbconnect.ListIndex = 0 Then
'this will connect the device through COM Port
'bconn = CZKEM1.Connect_Com(lngComPort, lngMachineNum, CLng(lngBaudRate)) 'THIS IS WHERE THE ERROR is
     
     bconn = CZKEM1.Connect_Net(Trim(ipAdd), 4370)
                
                'ElseIf cmbconnect.ListIndex = 1 Then
                                            'this will connect the device through Network
                'bConn = CZKEM1.Connect_Net(Trim(txtIP.Text), 4370)
'End If
vMachinenumber = lngMachineNum
'if connection is set, then changing the caption of the connect button
If bconn Then
cmdConnect.Caption = StrConv("Dis" & cmdConnect.Caption, vbProperCase)
cmdDownloadLog.Enabled = True
lblStatus.Caption = "Connected to the device."
CZKEM1.BASE64 = 0
Timer1.Enabled = True
'CZKEM1.RegEvent 1, 32767
Text1.SetFocus
Else
MsgBox "Can't connect to the specified device", vbCritical, "RFID Error"
lblStatus.Caption = ""
End
End If
CZKEM1.EnableDevice vMachinenumber, True

Else
lblStatus.Caption = "Disconnecting Device..."
cmdConnect.Caption = StrConv(Replace(cmdConnect.Caption, "Dis", ""), vbProperCase)
CZKEM1.Disconnect
'CZKEM1.EnableDevice cboMachineNo.Text, True
lblStatus.Caption = ""
End If
'Me.MousePointer = vbDefault
End Sub



Private Sub cmdDownload01_Click()
If lStatusKonek Then
If chkDownload.Value = 0 Then
cmdDownload01.Enabled = False
'MousePointer = vbHourglass
End If

Dim dwEnrollNumber As Long
Dim dwVerifyMode As Long
Dim dwInOutMode As Long
Dim timeStr As String
Dim I As Long
Dim lAddNew As Boolean

Dim v1 As String, v2 As Long, v3 As Long

lvx.Refresh
If chkDownload.Value = 1 Then
lvx.ListItems.Clear
End If

v1 = lngMachineNum
v2 = lngBaudRate
v3 = lngComPort

If CZKEM1.ReadGeneralLogData(v1) Then
I = 1
CZKEM1.ReadAllUserID v1 ''cboNoMesin01
While CZKEM1.GetGeneralLogDataStr(v1, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)

lvx.ListItems.Add I, , dwEnrollNumber
With lvx.ListItems(I)
.SubItems(1) = IIf(IsNull(timeStr), "", timeStr)
.SubItems(2) = IIf(IsNull(v1), "", v1)
.SubItems(3) = IIf(IsNull(dwVerifyMode), "", IIf(dwVerifyMode = 1, "Fingerprint", "Password"))
DoEvents
End With

Dim d_TimeStr As Date
d_TimeStr = CDate(Left(Right(Left(timeStr, 10), 2) & "-" & Mid(Left(timeStr, 10), 6, 2) & "-" & Left(Left(timeStr, 10), 4) & " " & Right(Trim(timeStr), 8), Len(Right(Left(timeStr, 10), 2) & "-" & Mid(Left(timeStr, 10), 6, 2) & "-" & Left(Left(timeStr, 10), 4) & " " & Right(Trim(timeStr), 8))))

lvx.ListItems.Add I, , dwEnrollNumber
With lvx.ListItems(I)
.SubItems(1) = IIf(IsNull(timeStr), "", timeStr)
.SubItems(2) = IIf(IsNull(v1), "", v1)
.SubItems(3) = IIf(IsNull(dwVerifyMode), "", IIf(dwVerifyMode = 1, "Fingerprint", "Password"))
DoEvents
End With




lvx.Refresh
Wend
End If



If chkDownload.Value = 0 Then
cmdDownload01.Enabled = True
'MousePointer = vbDefault
End If
'adoMesinConnect.Close
End If
End Sub




Private Sub Command2_Click()
Dim dwEnrollNumber As Long
Dim dwVerifyMode As Long
Dim dwInOutMode As Long
Dim timeStr As String
Dim I As Long
Dim lAddNew As Boolean



If CZKEM1.ReadGeneralLogData(vMachinenumber) Then
    I = 1
    CZKEM1.ReadAllUserID (vMachinenumber)
    
    While CZKEM1.GetGeneralLogDataStr(vMachinenumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)
    
        listview1.ListItems.Add I, , dwEnrollNumber
        With listview1.ListItems(I)
            .SubItems(1) = IIf(IsNull(timeStr), "", timeStr)
            .SubItems(2) = IIf(IsNull(lngMachineNum), "", lngMachineNum)
            .SubItems(3) = IIf(IsNull(dwVerifyMode), "", IIf(dwVerifyMode = 1, "Fingerprint", "Password"))
            DoEvents
        End With

        Dim d_TimeStr As Date
      '  d_TimeStr = CDate(Left(Right(Left(timeStr, 10), 2) & " - " & Mid(Left(timeStr, 10), 6, 2) & " - " & Left(Left(timeStr, 10), 4) & " & " & Right(Trim(timeStr), 8), Len(Right(Left(timeStr, 10), 2) & " - " & Mid(Left(timeStr, 10), 6, 2) & " - " & Left(Left(timeStr, 10), 4) & " & " & Right(Trim(timeStr), 8))))

        listview1.ListItems.Add I, , dwEnrollNumber
        With listview1.ListItems(I)
            .SubItems(1) = IIf(IsNull(timeStr), "", timeStr)
            .SubItems(2) = IIf(IsNull(vMachinenumber), "", vMachinenumber)
            .SubItems(3) = IIf(IsNull(dwVerifyMode), "", IIf(dwVerifyMode = 1, "Fingerprint", "Password"))
            DoEvents
        End With

        listview1.Refresh
    Wend
End If
End Sub

Private Sub Command4_Click()
Dim x As String
x = Trim$(Text5.Text)

Picture1.Picture = LoadPicture("C:\Documents and Settings\Administrator\My Documents\My Pictures\" & x & ".gif")
MsgBox "ok"


End Sub

Private Sub CZKEM1_OnAttTransaction(ByVal EnrollNumber As Long, ByVal IsInValid As Long, ByVal AttState As Long, ByVal VerifyMethod As Long, ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long)
Text4.Text = EnrollNumber
End Sub

Private Sub CZKEM1_OnHIDNum(ByVal CardNumber As Long)
Text1.Text = CardNumber
End Sub

Private Sub Form_Load()
CZKEM1.BASE64 = 0

End Sub

Private Sub Photo1_OnPhotoSaving(Succeded As Boolean, Filename As String)

End Sub

Public Sub tmrTimer_Timer()

Dim devModel As String
Dim reader As Integer
Dim devNo As Integer
Dim ipAdd As String
Dim porNo As Long
Dim comm As Long
Dim commPort As Long
Dim baudRate As Long
Dim EnrollNumber As String
Dim log As Long
Dim IsInValid As Long
Dim Year As Long
Dim Month As Long
Dim Day As Long
Dim Hour As Long
Dim Minute As Long
Dim Second As Long
Dim WorkCode As Long
Dim dwMachineNumber As Long
Dim dwEnrollNumber As Long
Dim dwVerifyMode As Long
Dim dwInOutMode As Long
Dim dwYear As Long
Dim dwMonth As Long
Dim dwDay As Long
Dim dwHour As Long
Dim dwMinute As Long
Dim dwSecond As Long
Dim dwWorkCode As Long
Dim dwReserved As Long

'MYSQL CONNECTOR VARIABLE
dataconnect
 
'ADITIONAL VARIABLE
Dim date_current As String
Dim time_current As String
Dim yr_tmp As String
Dim mth_tmp As String
Dim day_tmp As String
Dim hr_tmp As String
Dim min_tmp As String
Dim sec_tmp As String
Dim id As String
Dim code As String
Dim verify As String
Dim verify_tmp As String
Dim id_tmp As String


    Dim bconn As Boolean
    Dim mint As Integer
   
   
'Open log file
fn = FreeFile
Open "d:\log\log.txt" For Append As #fn
   

   
    'Date
   date_current = Format(Now, "yyyymmdd")
   time_current = Format(Now, "hhmmss")
 

'READER KORAYA
lngComPort = CLng(Val(comport.Text))
lngMachineNum = CLng(Val(machineno.Text))
lngBaudRate = 115200 'default value of Baudrate

'devModel = "AC800+"
devNo = "1"
'ipAdd = "192.168.1.201"
portNo = "1"
comm = "1"
commPort = "1"
baudRate = "115200"
reader = "1"
       
   
   
    If portNo = "" Then Exit Sub
        bconn = CZKEM1.Connect_Com(lngComPort, lngMachineNum, lngBaudRate)
        CZKEM1.Beep (150)
        CZKEM1.EnableDevice lngMachineNum, False
        
        If CZKEM1.ReadGeneralLogData(CInt(devNo)) And time_current > 60500 And time_current < 200000 Then
 '           Write #fn, & #34;============================================================================
 '"
            Write #fn, "Date : " + date_current + "     Time : " + time_current
            Write #fn, "Id         Date         Time         WorkCode    Reader      IpAdress      Verify"
            Do While CZKEM1.GetGeneralExtLogData(CInt(devNo), dwEnrollNumber, dwVerifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkCode, dwReserved)

 'log date
                yr_tmp = (dwYear)
                mth_tmp = (dwMonth)
                day_tmp = (dwDay)
               
                If mth_tmp < 10 Then
                    mth_tmp = "0" & mth_tmp
                Else
                End If
               
                If day_tmp < 10 Then
                    day_tmp = "0" & day_tmp
                Else
                End If
               
                date_tmp = yr_tmp & mth_tmp & day_tmp
               
'log time
                hr_tmp = (dwHour)
                min_tmp = (dwMinute)
                sec_tmp = (dwSecond)
                   
                If hr_tmp < 10 Then
                    hr_tmp = "0" & hr_tmp
                Else
                End If
               
                If min_tmp < 10 Then
                    min_tmp = "0" & min_tmp
                Else
                End If
               
                If sec_tmp < 10 Then
                    sec_tmp = "0" & sec_tmp
                Else
                End If
               
                time_tmp = hr_tmp & min_tmp & sec_tmp
               
                date_time = date_tmp & time_tmp
'id convert
   
                id = (dwEnrollNumber)
               
                                               
                code = (dwWorkCode)
               
'verify type
                verify = (dwVerifyMode)
               
                If verify = 0 Then
                   verify_tmp = "password"
                Else
                End If
               
                If verify = 1 Then
                   verify_tmp = "fprint"
                Else
                End If
               
                If verify = 2 Then
                   verify_tmp = "card"
                Else
                End If
                   
               
           
               
     'insert log to mysql
               
                If ((code = 0) And (time_tmp > 53000 And time_tmp < 100000)) Then
                   code = 1
                Else
                End If
         
                Write #fn, (id) + "    " & (date_tmp) + "    " & (time_tmp) + "             " & (code) & "         " & (reader) & "       " & (ipAdd) & "   " & (verify_tmp)
                ac.Execute "insert into rfid (a1,a2,a3,a4,a5,a6,a7)  values ('" & reader & "',1,'" & date_time & "','" & code & "','" & id & "' ,'" & verify & "',Null)"
        Loop
   

   
     'erase reader log
       
        If (time_current > 81500 And time_current < 81800) Or (time_current > 130500 And time_current < 130800) Or (time_current > 144500 And time_current < 144800) Or (time_current > 180000 And time_current < 180300) Or (time_current > 164500 And time_current < 164800) Or (time_current > 235000 And time_current < 235300) Then
           If CZKEM1.ClearGLog(CInt(devNo)) Then
              Write #fn, "Erasing Reader 1 log"
            Else
            End If
        Else
        End If
        Else
        End If
       
     
     
 'CLOSING MYSQL & LOG FILE

Close #fn



'RENAMING LOG FILE
    If (time_current > 180000 And time_current < 180200) Then
        Name "e:\log\log.txt" As "e:\log\" & (date_current) & "-" & (time_current) & "-log.txt"
    Else
    End If
   

End Sub

Private Sub Timer1_Timer()
 
CZKEM1.RegEvent 1, 32767

End Sub

 _
 _
 _
 _
 _
 _
 _
 _
 _
 _
 _
 _

