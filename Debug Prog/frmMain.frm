VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Number Pad"
      Height          =   2175
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   2535
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&D"
         Height          =   375
         Index           =   68
         Left            =   1920
         TabIndex        =   28
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&C"
         Height          =   375
         Index           =   67
         Left            =   1920
         TabIndex        =   27
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&B"
         Height          =   375
         Index           =   66
         Left            =   1920
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&A"
         Height          =   375
         Index           =   65
         Left            =   1920
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&*"
         Height          =   375
         Index           =   42
         Left            =   1320
         TabIndex        =   24
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&0"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   23
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&#"
         Height          =   375
         Index           =   35
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&9"
         Height          =   375
         Index           =   9
         Left            =   1320
         TabIndex        =   21
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&8"
         Height          =   375
         Index           =   8
         Left            =   720
         TabIndex        =   20
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&7"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&6"
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&5"
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&4"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&3"
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&2"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdDialPad 
         Caption         =   "&1"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4215
      Begin VB.CommandButton CmdDial 
         Caption         =   "&Dial"
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdHangUp 
         Caption         =   "Hang &Up"
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtTelephoneNumber 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton cmdSecureCall 
         Caption         =   "&Secure Call"
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdHold 
         Caption         =   "&Hold"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "&Accept"
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdAnswer 
         Caption         =   "A&nswer"
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdRedirect 
         Caption         =   "&Redirect"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   2160
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2
         Left            =   0
         Top             =   120
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Device:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CallObj As Object
Private WithEvents CallObj As TAPI.APPLICATION
Attribute CallObj.VB_VarHelpID = -1
Dim currentDevice As Integer

Private Sub CallObj_ACCEPTED(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_BUSY(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_Connected(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_DIALING(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_DIALTONE(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_DISCONNECTED(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_INCOMMINGMESSAGE(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_LINEREPLY(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_OFFERING(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_APPNEWCALL(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_REQUESTMAKECALL(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3

    Debug.Print CallObj.Request.ApplicationName, CallObj.Request.CalledParty, CallObj.Request.Comment, CallObj.Request.DestinationAddress
End Sub

Private Sub CallObj_CallInfo(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CallObj_IDLE(ByVal DeviceIndex As Long, ByVal hDevice As Long, ByVal dwMsg As Long, ByVal dwCallbackInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    Debug.Print CallObj.devices.Item(DeviceIndex).Last_Event, CallObj.devices.Item(DeviceIndex).Last_Function, DeviceIndex, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3
End Sub

Private Sub CmdDialPad_Click(Index As Integer)
    If CallObj.devices.Item(0).Calls.Count > 0 Then
        Select Case Index
            Case 0 To 9: CallObj.devices.Item(0).Calls.Item(0).Func_lineGenerateDigits &H2, Index, 2
            Case Else: CallObj.devices.Item(0).Calls.Item(0).Func_lineGenerateDigits &H2, Chr(Index), 2
        End Select
    Else
        Select Case Index
            Case 0 To 9: Me.txtTelephoneNumber.Text = txtTelephoneNumber.Text & Trim(Str(Index))
        End Select
    End If
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = currentDevice Then Exit Sub
    Call CallObj.devices.Item(currentDevice).Func_lineClose
    currentDevice = Combo1.ListIndex
    Call CallObj.devices.Item(Combo1.ListIndex).Func_lineOpen(&H4& Or &H2&, &H10&)
End Sub

Private Sub Form_Load()
    Set CallObj = New TAPI.APPLICATION
    'Set CallObj = CreateObject("TAPI.APPLICATION")
    Let CallObj.Debugging = True
    If CallObj.devices.Count <= 0 Then End
    
    currentDevice = 0
    
    Call CallObj.devices.Item(currentDevice).Func_lineOpen(&H4& Or &H2&, &H10&)
    Call CallObj.Func_lineRegisterRequestRecipient
    
    Dim i As Integer
    
    For i = 0 To CallObj.devices.Count - 1
        Combo1.AddItem CallObj.devices.Item(i).DeviceName, i
    Next i
    
    Let Combo1.ListIndex = 0
    Let CallObj.Priority.DataModem = True
    Let CallObj.Priority.InteractiveVoice = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Let CallObj.Priority.DataModem = False
    Let CallObj.Priority.InteractiveVoice = False
    
    Call CallObj.Func_lineUnregisterRequestRecipient
    Call CallObj.devices.Item(Combo1.ListIndex).Func_lineClose
    
    Set CallObj = Nothing
End Sub

Private Sub cmdAccept_Click()
    Call CallObj.devices.Item(Combo1.ListIndex).Calls.Item(0).Func_lineAccept("")
End Sub

Private Sub cmdAnswer_Click()
    Call CallObj.devices.Item(Combo1.ListIndex).Calls.Item(0).Func_lineAnswer("")
End Sub

Private Sub CmdDial_Click()
    Call CallObj.devices.Item(Combo1.ListIndex).Func_lineMakeCall(txtTelephoneNumber.Text, CallObj.CountryCode)
End Sub

Private Sub CmdHangUp_Click()
    Call CallObj.devices.Item(Combo1.ListIndex).Calls.Item(0).Func_LineDrop("")
End Sub

Private Sub cmdSecureCall_Click()
    Call CallObj.devices.Item(Combo1.ListIndex).Calls.Item(0).Func_lineSecureCall
End Sub

Private Sub txtTelephoneNumber_DblClick()
    txtTelephoneNumber.Text = ""
End Sub
