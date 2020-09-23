VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Serial Port Send And Recieve"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbBaudRate 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   1080
      List            =   "frmMain.frx":0002
      TabIndex        =   25
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdateBaud 
      Caption         =   "Update Baud"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   22
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame fraPorts 
      Caption         =   "Comm Port"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton optPort 
         Caption         =   "1"
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton optPort 
         Caption         =   "1"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton optPort 
         Caption         =   "1"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton optPort 
         Caption         =   "1"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton optPort 
         Caption         =   "1"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton optPort 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdClearHistory 
      Caption         =   "Clear History"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ListBox lstHistory 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   5400
      TabIndex        =   15
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame fraBit 
      Caption         =   "Data Bits"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optDataBits 
         Caption         =   "8"
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optDataBits 
         Caption         =   "7"
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optDataBits 
         Caption         =   "6"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton optDataBits 
         Caption         =   "5"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame fraMode 
      Caption         =   "Mode"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2055
      Begin VB.OptionButton optString 
         Caption         =   "String"
         Height          =   195
         Left            =   960
         TabIndex        =   9
         Top             =   270
         Width           =   735
      End
      Begin VB.OptionButton optBinary 
         Caption         =   "Binary"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtRead 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2160
      Width           =   5175
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin MSCommLib.MSComm comSerial 
      Left            =   3480
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      BaudRate        =   38400
      InputMode       =   1
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame fraDataRead 
      Caption         =   "Data Recieved"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Frame fraSend 
      Caption         =   "Data To Send"
      Height          =   255
      Left            =   2280
      TabIndex        =   24
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lblBaud 
      Alignment       =   1  'Right Justify
      Caption         =   "Baud Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1470
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   File:
'       frmMain.frm
'   Author:
'       Tom DeWitt
'   Description:
'       This program was designed as a utility tool to allow the user to send and recieve data from a serial com port
'while monitory the byte data. If a loopback ciruit is made on the com port(jumper send to recieve DB9 pins 2 & 3)
'a developer can also monitor the way VB interprets the data that is being sent or recieved. The port settings can be
'configured as desired and the last setting are stored in the registry under the Key 'HKEY_LOCAL_MACHINE\Software\
'Damage Inc\Com Settings'. The Input Mode Property Defaults to binary. It was not designed to poll the com port or use
'handshaking with a device, it was intended to send data to a device and then read the devices response as a simple
'tool. The code maybe modified as desired and may be freely redistributed. I hold no responsiblity for the way it is
'used. It has been tested on Windows NT 4.0 SrvPk 6a and on Windows 2000 SrvPk 2. It was developed under VB6 SrvPk 5.
'As the author of the code I may not have tested all possible user misuse and abuse, as I only know what my intentions
'for the code were, not how it could possibly be used. It was not tested as a beta release.
'-----------------------------------------------------------------------------------------------------------------------
'   Revisions:
'       Original 2/7/2002
'-----------------------------------------------------------------------------------------------------------------------
'   Functions And Subroutines:
'   1.  BitOn(Number As Long, Bit As Long) As Boolean
'           Performs Bitwise Check on 'Number', Returns True if 'Bit' is On
'   2. VerifyPorts() Checks The Registry Entries For The ComPorts On The Current System
'   3. UpdateBaud() Changes The ComPort's Baud Rate And Calls The UpdateSettings() Sub To Update The Registry
'   4. VerifySettings() Check The Registry For The ComPorts Last Settings. If There is No Registry Entry It Creates One
'       And Places The Default(Com1 38400,n,8,1) Setting in The Registry. If There Is An Entry It Sets The ComPort.
'   5.UpdateSettings() Changes The Registry Entry When The User Chages The Com Port Or Settings. It Does Not Update The
'       InputMode which Defaults To Binary.
'-----------------------------------------------------------------------------------------------------------------------
'   Properties:
'-----------------------------------------------------------------------------------------------------------------------
'   Required Functions,Subroutines,Properties,Variables,Etc.:
'
'-----------------------------------------------------------------------------------------------------------------------
'   Variables:
'       Public:
'
'-----------------------------------------------------------------------------------------------------------------------
'       Private:
Private bLoaded As Boolean  'True After Form Is Loaded --> Enables Option Button Click Events
Private sDataBits As String
Private sMode As String
Private BaudRate(12) As String
Private Ports() As Variant
Private sBaudData As String
Private sSubKey As String
Private sKeyValue As String
Private sSettings As String
Private sPortNum As String
Private hnd As Long
'-----------------------------------------------------------------------------------------------------------------------
'       Constants:
Private Const lMainKey As Long = HKEY_LOCAL_MACHINE
Private Const lLength As Long = 1024
Private Const sSettingsKey As String = "Settings"
Private Const sPortKey As String = "Port"
'-----------------------------------------------------------------------------------------------------------------------
'       Special Notes:
'           Printing Line Length is 120 Characters
'-----------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------
'       Enumeration Constants:
'-----------------------------------------------------------------------------------------------------------------------
'Initialize The BaudRate Array As Set Values ie Make Shift Constants
'---------------------------------------------------------START---------------------------------------------------------
Private Sub Form_Initialize()
    BaudRate(0) = "110"
    BaudRate(1) = "300"
    BaudRate(2) = "600"
    BaudRate(3) = "1200"
    BaudRate(4) = "2400"
    BaudRate(5) = "9600"
    BaudRate(6) = "14400"
    BaudRate(7) = "19200"
    BaudRate(8) = "28800"
    BaudRate(9) = "38400"
    BaudRate(10) = "56000"
    BaudRate(11) = "128000"
    BaudRate(12) = "256000"
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbBaudRate_Click()
        sBaudData = ""
        cmdUpdateBaud.Enabled = True
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Monitor Keyboard Input While Editing Baud Rate
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbBaudRate_KeyPress(KeyAscii As Integer)
        Select Case KeyAscii
            Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57                         'Allow Only Numbers
                sBaudData = sBaudData & Chr$(KeyAscii)
            Case 13                                                             'Enter Pressed
                sBaudData = ""
                UpdateBaud
            Case 127                                                            'Delete Key Pressed
                sBaudData = ""
            Case 8                                                              'Backspace Key Pressed
                If sBaudData <> "" Then
                    sBaudData = Left$(sBaudData, (Len(sBaudData) - 1))
                End If
            Case Else
                sBaudData = ""
        End Select
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Auto Fill Allowable Baud Rates
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbBaudRate_Change()
    Dim iX As Long
    Dim iL As Long
    Dim sCurrent As String

        If bLoaded Then
            cmdUpdateBaud.Enabled = True
            sCurrent = sBaudData
            iL = Len(sCurrent)
                For iX = 0 To 12
                    If sCurrent = Left$(BaudRate(iX), iL) Then
                        cmbBaudRate.Text = BaudRate(iX)
                        cmbBaudRate.SelLength = Len(cmbBaudRate.Text)
                        Exit Sub
                    End If
                Next
        End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdUpdateBaud_Click()
        UpdateBaud
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Check If The Requested Bit Is 'On' In The Given Number
'---------------------------------------------------------START---------------------------------------------------------
Function BitOn(Number As Long, Bit As Long) As Boolean
    Dim iX As Long
    Dim iY As Long

        iY = 1
        For iX = 1 To Bit - 1
            iY = iY * 2
        Next
        If Number And iY Then BitOn = True Else BitOn = False
End Function
'----------------------------------------------------------END----------------------------------------------------------
'Open The Local Machine Registry And Get The Serial Ports Available On The Local Machine, Validate Selected Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub VerifyPorts()
    Dim sPort As String
    Dim iX As Long
    Dim iY As Long
    Dim lngType As Long
    Dim lngValue As Long
    Dim sName As String
    Dim sSwap As String
    ReDim varResult(0 To 1, 0 To 100) As Variant
    Const lNameLen As Long = 260
    Const lDataLen As Long = 4096

        sSubKey = "Hardware\Devicemap\SerialComm"
        If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_READ, hnd) Then Exit Sub
            For iX = 0 To 999999
                If iX > UBound(varResult, 2) Then
                    ReDim Preserve varResult(0 To 1, iX + 99)
                End If
                sName = Space$(lNameLen)
                ReDim binValue(0 To lDataLen - 1) As Byte
                If RegEnumValue(hnd, iX, sName, lNameLen, ByVal 0&, lngType, binValue(0), lDataLen) Then Exit For
                    varResult(0, iX) = Left$(sName, lNameLen)
                    
                    Select Case lngType
                        Case REG_DWORD
                            CopyMemory lngValue, binValue(0), 4
                            varResult(1, iX) = lngValue
                        Case REG_SZ
                            varResult(1, iX) = Left$(StrConv(binValue(), vbUnicode), lDataLen - 1)
                        Case Else
                            ReDim Preserve binValue(0 To lDataLen - 1) As Byte
                            varResult(1, iX) = binValue()
                    End Select
            Next
        If hnd Then RegCloseKey hnd                                             'Close The Registry Key
        ReDim Preserve varResult(0 To 1, iX - 1) As Variant
        ReDim Ports(iX - 1)
        For iX = 0 To UBound(varResult, 2)                                      'Trim 'Port' To Get Just The Number
            sPort = Mid$(varResult(1, iX), 4, 1)
            Ports(iX) = sPort
        Next

        iY = UBound(Ports)                                                       'Arrange The Ports Numbers Low To High
        For iX = 0 To (iY - 1)
            If Ports(iX + 1) < Ports(iX) Then
                sSwap = Ports(iX + 1)
                Ports(iX + 1) = Ports(iX)
                Ports(iX) = sSwap
                iX = -1
            End If
        Next

End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Changes The ComPorts Baud Rate
'---------------------------------------------------------START---------------------------------------------------------
Private Sub UpdateBaud()
Attribute UpdateBaud.VB_Description = "Changes the baud rate of the serial port"
    Dim sNewBaud As String
    Dim sOldBaud As String
    Dim sTmp As String
    Dim iX As Long
    
On Error GoTo ErrTrap

        sNewBaud = cmbBaudRate.Text
        For iX = 0 To 12
            If BaudRate(iX) = sNewBaud Then
                Exit For
            Else
                If iX = 12 Then
                    MsgBox "Invalid Baud Rate, Please Try Again !", vbInformation, "Data Entry Error !"
                    sBaudData = ""
                    cmbBaudRate.Text = ""
                    cmdUpdateBaud.Enabled = False
                    Exit Sub
                End If
            End If
        Next
        sTmp = comSerial.Settings
        sOldBaud = Left$(sTmp, (InStr(1, sTmp, ",", vbBinaryCompare) - 1))
        sTmp = Replace(sTmp, sOldBaud, sNewBaud, , , vbBinaryCompare)
        comSerial.Settings = sTmp
        cmdUpdateBaud.Enabled = False
        sBaudData = ""
        UpdateSettings
    Exit Sub

ErrTrap:
        MsgBox Err.Number & " " & Err.Description & vbCr & " Error Generated By " & Err.Source, vbCritical, _
"System Error Trap !"
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
    Set frmMain = Nothing
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Read Input Data Then Display In The txtRead Box
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdRead_Click()
    Dim bytInput() As Byte
    Dim bytElement As Byte
    Dim iX As Long
    Dim iY As Long
    Dim iL As Long
    Dim iP As Long
    Dim sResult As String
    Dim sHistory As String
    Dim sData As String
    Dim sSpace As String

On Error GoTo ErrTrap

        If comSerial.PortOpen = False Then
            comSerial.PortOpen = True
        End If

        bytInput = comSerial.Input                                              'Read The Input and Get Its Size
        iX = UBound(bytInput(), 1)
        For iY = 0 To iX
            If sResult <> "" Then                                               'Ignore Padding on First Byte
                If iY Mod 4 Then                                                'Pad Between Bytes
                    sResult = "    " & sResult
                Else
                    sResult = vbCrLf & sResult                                  'Start New Line After 4 Bytes
                End If
            End If
            bytElement = bytInput(iY)                                           'Get Single Byte Element
            sData = Chr$(bytElement)                                            'and Its Character
            For iL = 1 To 8                                                     'Iterate Each Bit of the Byte
                Select Case iL
                    Case 4                                                      'Comma Deliminate Each Digit
                        sSpace = " , "
                    Case Else
                        sSpace = ""
                End Select
                sResult = sSpace & Abs(CInt(BitOn(CLng(bytElement), iL))) & sResult
            Next
            If sResult <> "" Then
                If Asc(sData) = 0 Then                                          'Check and Replace Null
                    sData = "~"                                                 '~ Replaces Null, Change If Desired
                End If
                sResult = "(" & sData & ")> " & sResult
            End If
        Next
        txtRead.Text = sResult & vbCrLf
        cmdRead.Enabled = False
        lstHistory.AddItem ("Read " & sDataBits & " Bits" & " As " & sMode)     'Write Line To The History List
        Do While Len(sResult)                                                   'Parse Thru Result And Create History
            iP = InStrRev(sResult, "(", , vbBinaryCompare)
            sHistory = Replace(Trim(Mid$(sResult, iP)), vbCrLf, "", , , vbBinaryCompare)
            sResult = Left(sResult, (iP - 1))
            lstHistory.AddItem (sHistory & " :ASCII " & CStr(Asc(Mid$(sHistory, 2, 1))))
        Loop
        txtSend.SetFocus                                                        'Select Text In txtSnd Box
        txtSend.SelStart = 0
        txtSend.SelLength = Len(txtSend.Text)
        cmdClearHistory.Enabled = True                                          'Enable Clear History Button
    Exit Sub

ErrTrap:
        MsgBox Err.Number & " " & Err.Description & vbCr & " Error Generated By " & Err.Source, vbCritical, _
"System Error Trap !"
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Send Data That Is Displayed In The txtSend Box
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdSend_Click()

On Error GoTo ErrTrap

        If comSerial.PortOpen = False Then
            comSerial.PortOpen = True
        End If
        comSerial.Output = txtSend.Text                                         'Write Line To The History List
        cmdRead.Enabled = True
        lstHistory.AddItem ("Send " & sDataBits & " Bits" & " As " & sMode)
        lstHistory.AddItem txtSend.Text
    Exit Sub

ErrTrap:
        MsgBox Err.Number & " " & Err.Description & vbCr & " Error Generated By " & Err.Source, vbCritical, _
"System Error Trap !"
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Get & Display Port Settings, Enable Option Button Click After Loaded(bLoaded = True)
'---------------------------------------------------------START---------------------------------------------------------
Private Sub Form_Load()
    Dim iX As Long
    Dim iY As Long
    Dim sTmp As String
    Dim sPort As String
    Dim sSelectedPort As String
    Dim bFlag As Boolean
    Dim opt As OptionButton

        VerifyPorts
        VerifySettings
        sSettings = comSerial.Settings
        sSelectedPort = comSerial.CommPort
        Select Case comSerial.InputMode
            Case comInputModeBinary
                optBinary.Value = True
                sMode = "Binary"
            Case comInputModeText
                optString.Value = True
                sMode = "String"
        End Select
        For iX = 0 To UBound(BaudRate())
            cmbBaudRate.AddItem BaudRate(iX)
        Next
        sTmp = Left$(sSettings, (InStr(1, sSettings, ",", vbBinaryCompare) - 1))
        sDataBits = Left$(Right$(sSettings, 3), 1)
        optDataBits(CInt(sDataBits)).Value = True
        cmbBaudRate.Text = sTmp
        
        iY = UBound(Ports)
        For iX = 0 To iY                                                        'Enable The Approriate Option Buttons
            sPort = Ports(iX)
            optPort(iX).Visible = True
            optPort(iX).Caption = sPort
            If sPort = sSelectedPort Then
                bFlag = True
                optPort(iX).Value = True
            End If
        Next
        If Not bFlag Then                                                       ' If Port Doesn't Exist Use 1st One
            comSerial.CommPort = CInt(optPort(0).Caption)
            optPort(0).Value = True
        End If
        bLoaded = True
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Switch Port Mode To Binary
'---------------------------------------------------------START---------------------------------------------------------
Private Sub optBinary_Click()
        If bLoaded Then
            comSerial.InputMode = comInputModeBinary
            sMode = "Binary"
        End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Switch Port Data Bits To Selected Option
'---------------------------------------------------------START---------------------------------------------------------
Private Sub optDataBits_Click(Index As Integer)
    Dim sTmp As String

On Error GoTo ErrTrap

        If bLoaded Then
            sTmp = comSerial.Settings
            Mid(sTmp, (Len(sTmp) - 2), 1) = CStr(Index)
            sDataBits = CStr(Index)
            comSerial.Settings = sTmp
            UpdateSettings
        End If
    Exit Sub

ErrTrap:
        MsgBox Err.Number & " " & Err.Description & vbCr & " Error Generated By " & Err.Source, vbCritical, _
"System Error Trap !"
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Change The Comm Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub optPort_Click(Index As Integer)
        If bLoaded Then
            comSerial.CommPort = CInt(optPort(Index).Caption)
            UpdateSettings
        End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Switch Port Mode To String
'---------------------------------------------------------START---------------------------------------------------------
Private Sub optString_Click()
        If bLoaded Then
            comSerial.InputMode = comInputModeText
            sMode = "String"
        End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Disable cmdSend Button When The txtSend Box Is Empty
'---------------------------------------------------------START---------------------------------------------------------
Private Sub txtSend_Change()
    If txtSend.Text <> "" Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Clear History List Box
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdClearHistory_Click()
    lstHistory.Clear
    cmdClearHistory.Enabled = False
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Select All Text in txtSend Box
'---------------------------------------------------------START---------------------------------------------------------
Private Sub txtSend_GotFocus()
        txtSend.SelStart = 0
        txtSend.SelLength = Len(txtSend.Text)
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Check The Registy For The Last Used Settings And Sets The MSComm Object Properties. If There Is No Entry It Creates
'One With The Default Setting(Com1 38400n,8,1)
'---------------------------------------------------------START---------------------------------------------------------
Private Sub VerifySettings()
Attribute VerifySettings.VB_Description = "Checks the registry for the last com port settings"
    Dim disposition As Long
    Dim sTmp As String

On Error GoTo ErrTrap

        sSettings = comSerial.Settings
        sPortNum = comSerial.CommPort
        sSubKey = "Software\Damage Inc\Com Settings"
        If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_READ, hnd) Then
            If RegCreateKeyEx(lMainKey, sSubKey, 0, 0, 0, 0, 0, hnd, disposition) Then
                Err.Raise 1001, "VerifySettings() Sub", "Could Not Create Registry Key"
            End If
        End If

'The Key Has Been Found/or Created, Now Check To See If Previous Settings Are Present

'Check For The Settings Subkey and Retrieve Value If Present, Then Set ComPort 'Settings' Property

        sKeyValue = Space$(lLength)                                             'Pad The sKeyValue Variable
        If RegQueryValueEx(hnd, sSettingsKey, 0, REG_SZ, ByVal sKeyValue, lLength) Then     '0 Return if Successful
            If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_WRITE, hnd) Then                      '0 Return if Successful
                Err.Raise 1001, "VerifySettings() Sub", "Could Not Open Registry Key"
            Else        'The Value Was Not Present, Set To Default Port 'Settings' Property
                If RegSetValueEx(hnd, sSettingsKey, 0, REG_SZ, ByVal sSettings, Len(sSettings)) Then
                    Err.Raise 1001, "VerifySettings() Sub", "Could Not Set Registry Key Settings Value"
                End If
            End If
        Else            'Read Value From Key And Set The Port 'Settings' Property To The Value In The Registry
            comSerial.Settings = sKeyValue
        End If

'Check For The Port Subkey and Retrieve Value If Present, Then Set ComPort 'Port' Property

        sKeyValue = Space$(lLength)                                             'Pad The sKeyValue Variable
        If RegQueryValueEx(hnd, sPortKey, 0, REG_SZ, ByVal sKeyValue, lLength) Then         '0 Return if Successful
            If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_WRITE, hnd) Then                      '0 Return if Successful
                Err.Raise 1001, "VerifySettings() Sub", "Could Not Open Registry Key"
            Else        'The Value Was Not Present, Set To Default Port 'Port' Property
                If RegSetValueEx(hnd, sPortKey, 0, REG_SZ, ByVal sPortNum, Len(sPortNum)) Then
                    Err.Raise 1001, "VerifySettings() Sub", "Could Not Set Registry Key Port Value"
                End If
            End If
        Else            'Read Value From Key And Set The Port 'Port' Property To The Value In The Registry
            comSerial.CommPort = sKeyValue
        End If

        RegCloseKey hnd
    Exit Sub

ErrTrap:
        MsgBox Err.Number & " " & Err.Description & vbCr & " Error Generated By " & Err.Source, vbCritical, _
"System Error Trap !"
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Changes The Registry Entries When The User Changes Port Settings
'---------------------------------------------------------START---------------------------------------------------------
Private Sub UpdateSettings()
Attribute UpdateSettings.VB_Description = "Updates the registry entry to the current com port settings"

On Error GoTo ErrTrap

        sSettings = comSerial.Settings
        sPortNum = comSerial.CommPort
        sSubKey = "Software\Damage Inc\Com Settings"

            If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_WRITE, hnd) Then                      '0 Return if Successful
                Err.Raise 1001, "VerifySettings() Sub", "Could Not Open Registry Key"
            Else        'The Value Was Not Present, Set To Default Port 'Settings' Property
                If RegSetValueEx(hnd, sSettingsKey, 0, REG_SZ, ByVal sSettings, Len(sSettings)) Then
                    Err.Raise 1001, "VerifySettings() Sub", "Could Not Set Registry Key Settings Value"
                End If
            End If

            If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_WRITE, hnd) Then                      '0 Return if Successful
                Err.Raise 1001, "VerifySettings() Sub", "Could Not Open Registry Key"
            Else        'The Value Was Not Present, Set To Default Port 'Port' Property
                If RegSetValueEx(hnd, sPortKey, 0, REG_SZ, ByVal sPortNum, Len(sPortNum)) Then
                    Err.Raise 1001, "VerifySettings() Sub", "Could Not Set Registry Key Port Value"
                End If
            End If

    Exit Sub

ErrTrap:
        MsgBox Err.Number & " " & Err.Description & vbCr & " Error Generated By " & Err.Source, vbCritical, _
"System Error Trap !"
End Sub
'----------------------------------------------------------END----------------------------------------------------------
