VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APPLICATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"AppPriorityObj"
Attribute VB_Ext_KEY = "Member1" ,"Devices"
Attribute VB_Ext_KEY = "Member2" ,"RequestObj"
Option Explicit

Private Const TAPI_LOWAPI = &H10004
Private Const TAPI_HIGHAPI = &H20001

Private Const LINECALLSTATE_IDLE = &H1&
Private Const LINECALLSTATE_OFFERING = &H2&
Private Const LINECALLSTATE_ACCEPTED = &H4&
Private Const LINECALLSTATE_DIALTONE = &H8&
Private Const LINECALLSTATE_DIALING = &H10&
Private Const LINECALLSTATE_RINGBACK = &H20&
Private Const LINECALLSTATE_BUSY = &H40&
Private Const LINECALLSTATE_SPECIALINFO = &H80&
Private Const LINECALLSTATE_CONNECTED = &H100&
Private Const LINECALLSTATE_PROCEEDING = &H200&
Private Const LINECALLSTATE_ONHOLD = &H400&
Private Const LINECALLSTATE_CONFERENCED = &H800&
Private Const LINECALLSTATE_ONHOLDPENDCONF = &H1000&
Private Const LINECALLSTATE_ONHOLDPENDTRANSFER = &H2000&
Private Const LINECALLSTATE_DISCONNECTED = &H4000&
Private Const LINECALLSTATE_UNKNOWN = &H8000&

Private Const LINE_ADDRESSSTATE = 0&
Private Const LINE_CALLINFO = 1&
Private Const LINE_CALLSTATE = 2&
Private Const LINE_CLOSE = 3&
Private Const LINE_DEVSPECIFIC = 4&
Private Const LINE_DEVSPECIFICFEATURE = 5&
Private Const LINE_GATHERDIGITS = 6&
Private Const LINE_GENERATE = 7&
Private Const LINE_LINEDEVSTATE = 8&
Private Const LINE_MONITORDIGITS = 9&
Private Const LINE_MONITORMEDIA = 10&
Private Const LINE_MONITORTONE = 11&
Private Const LINE_REPLY = 12&
Private Const LINE_REQUEST = 13&
Private Const LINE_CREATE = 19&             ' TAPI v1.4
Private Const LINE_AGENTSPECIFIC = 21&      ' TAPI v2.0
Private Const LINE_AGENTSTATUS = 22&        ' TAPI v2.0
Private Const LINE_APPNEWCALL = 23&         ' TAPI v2.0
Private Const LINE_PROXYREQUEST = 24&       ' TAPI v2.0
Private Const LINE_REMOVE = 25&             ' TAPI v2.0

Private Const LINEREQUESTMODE_MAKECALL = &H1&
Private Const LINEREQUESTMODE_MEDIACALL = &H2&
Private Const LINEREQUESTMODE_DROP = &H4&

Private Const TAPIMAXDESTADDRESSSIZE = 80&
Private Const TAPIMAXAPPNAMESIZE = 40&
Private Const TAPIMAXCALLEDPARTYSIZE = 40&
Private Const TAPIMAXCOMMENTSIZE = 80&
Private Const TAPIMAXDEVICECLASSSIZE = 40&
Private Const TAPIMAXDEVICEIDSIZE = 40&

Private Type LINEREQMAKECALL
    szDestAddress As String * TAPIMAXDESTADDRESSSIZE
    szAppName As String * TAPIMAXAPPNAMESIZE
    szCalledParty As String * TAPIMAXCALLEDPARTYSIZE
    szComment As String * TAPIMAXCOMMENTSIZE
End Type

Private Type LINEINITIALIZEEXPARAMS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    dwOptions As Long
    hEvent As Long 'union hEvent and Completion port
    dwCompletionKey As Long
End Type

Private Type LINEREQMEDIACALL
    hWnd As Integer
    wRequestID As Integer
    szDeviceClass As String * TAPIMAXDEVICECLASSSIZE
    ucDeviceID As String * TAPIMAXDEVICEIDSIZE
    dwSize As Long
    dwSecure As Long
    szDestAddress As String * TAPIMAXDESTADDRESSSIZE
    szAppName As String * TAPIMAXAPPNAMESIZE
    szCalledParty As String * TAPIMAXCALLEDPARTYSIZE
    szComment As String * TAPIMAXCOMMENTSIZE
End Type



Private Declare Function tapiGetLocationInfo Lib "TAPI32.DLL" (ByVal lpszCountryCode As String, ByVal lpszCityCode As String) As Long

Private Declare Function lineInitializeEx Lib "TAPI32.DLL" Alias "lineInitializeExA" (ByRef lphLineApp As Long, ByVal hInstance As Long, ByVal lpfnCallback As Long, ByVal lpszFriendlyAppName As String, ByRef lpdwNumDevs As Long, ByRef lpdwAPIVersion As Long, ByRef lpLineInitializeExParams As LINEINITIALIZEEXPARAMS) As Long
Private Declare Function lineShutdown Lib "TAPI32.DLL" (ByVal hLineApp As Long) As Long
Private Declare Function lineGetID Lib "tapi.dll" (ByVal hLine As Integer, ByVal dwAddressID As Long, ByVal hCall As Integer, ByVal dwSelect As Long, lpDeviceID As String, ByVal lpszDeviceClass As String)
Private Declare Function lineRegisterRequestRecipient Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwRegistrationInstance As Long, ByVal dwRequestMode As Long, ByVal bEnable As Long) As Long

Public Event INCOMMINGMESSAGE( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event LINEREPLY( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event IDLE( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event OFFERING( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event ACCEPTED( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event DIALTONE( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event DIALING( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event RINGBACK( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event BUSY( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event SpecialInfo( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event Connected( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event PROCEEDING( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event ONHOLD( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event CONFERENCED( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event ONHOLDPENDCONF( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event ONHOLDPENDTRANSFER( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event DISCONNECTED( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event Unknown( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event AddressState( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event CallInfo( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event lineClose( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event CREATE( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event Remove( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event devSPECIFIC( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event DEVSTATE( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event REQUESTMAKECALL( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event REQUESTDROP( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event REQUESTMEDIACALL( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event AGENTSPECIFIC( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event AGENTSTATUS( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event APPNEWCALL( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
Public Event PROXYREQUEST( _
    ByVal DeviceIndex As Long, _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)

Private numDevices As Long 'This is used to store the amount of devices in the system
Private mvarDevices As Devices
Private mvarAppPriorityObj As AppPriorityObj
Private mvarRequestObj As RequestObj

Public Property Get Request() As RequestObj
Attribute Request.VB_Description = "This returns a RequestObj."
    If mvarRequestObj Is Nothing Then Set mvarRequestObj = New RequestObj
    Set Request = mvarRequestObj
End Property

Public Property Get Priority() As AppPriorityObj
Attribute Priority.VB_Description = "This returns a AppPriorityObj."
    If mvarAppPriorityObj Is Nothing Then Set mvarAppPriorityObj = New AppPriorityObj
    Set Priority = mvarAppPriorityObj
End Property

Public Property Get Devices() As Devices
Attribute Devices.VB_Description = "This returns a collection of DeviceItemObj."
Attribute Devices.VB_UserMemId = 0
    If mvarDevices Is Nothing Then Set mvarDevices = New Devices
    Set Devices = mvarDevices
End Property

Private Sub Class_Initialize()
    Let Module1.AppObjPtr = ObjPtr(Me) ' Setting the applications object pointer to a global variable.
    
    Call Func_lineInitializeEx ' Initializing TAPI
    
    ReDim Module1.DeviceItemHndCall(0 To numDevices, 0 To 100) As Long
    ReDim Module1.LastEvent(0 To numDevices) As String
    ReDim Module1.LastFunction(0 To numDevices) As String
    
    Set mvarDevices = New Devices
    
    For DeviceItem = 0 To numDevices - 1
        mvarDevices.Add New DeviceItemObj, DeviceItem
    Next DeviceItem
End Sub

Private Sub Class_Terminate()
    Set mvarRequestObj = Nothing
    Set mvarAppPriorityObj = Nothing
    Set mvarDevices = Nothing
    
    Erase Module1.DeviceItemHndCall
    
    Call Func_lineShutdown  ' Shutting down TAPI
End Sub

Private Function Func_lineShutdown() As Long
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineShutdown"
        Close #FreeFileNo
    End If
    
    Func_lineShutdown = lineShutdown(Module1.hndLine)
End Function

Private Function Func_lineInitializeEx() As Long
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineInitializeEx"
        Close #FreeFileNo
    End If
    
    Dim LineInitializeExParam  As LINEINITIALIZEEXPARAMS
    
    LineInitializeExParam.dwTotalSize = Len(LineInitializeExParam)
    
    Func_lineInitializeEx = lineInitializeEx( _
        hndLine, _
        App.hInstance, _
        CLng(Val(AddressOf LINECALLBACK)), _
        App.EXEName & vbNullString, _
        numDevices, _
        TAPI_HIGHAPI, _
        LineInitializeExParam _
    )
End Function

Public Property Get CountryCode() As Long
Attribute CountryCode.VB_Description = "Local Country Code.  For Example, Ireland = (353)"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Debug print the call back message to file '''''''''''''''''''''''''''''
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: tapiGetLocationInfo"
        Close #FreeFileNo
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Dim strCountryCode$, strcityCode$, retVal&
    
    Let strCountryCode = String$(128, Chr(0))
    Let strcityCode = String$(128, Chr(0))
    Let retVal = tapiGetLocationInfo(strCountryCode, strcityCode)
    
    If retVal = 0 Then CountryCode = CLng(Val(strCountryCode))
End Property

Public Property Get CityCode() As Long
Attribute CityCode.VB_Description = "Local City Code.  For example,  Dublin = (01)"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Debug print the call back message to file '''''''''''''''''''''''''''''
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: tapiGetLocationInfo"
        Close #FreeFileNo
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Dim strCountryCode$, strcityCode$, retVal&
    
    Let strCountryCode = String$(128, Chr(0))
    Let strcityCode = String$(128, Chr(0))
    Let retVal = tapiGetLocationInfo(strCountryCode, strcityCode)
    
    If retVal = 0 Then CityCode = CLng(Val(strcityCode))
End Property

Friend Sub RecieveCallBack( _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
    On Error Resume Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Index number from call back instance number ''''''''''''''''''''''
    Dim Index%
    
    For Index = 0 To Devices.Count - 1
        If ObjPtr(Devices.Item(Index)) = dwCallbackInstance Then Exit For
    Next Index
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Debug print the call back messag to file '''''''''''''''''''''''''''''
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Debug.Print hDevice; Tab(25); LoadResString(dwMsg); Tab(45); dwParam1; Tab(55); dwParam2; Tab(65); dwParam3
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, hDevice, LoadResString(dwMsg), dwParam1, dwParam2, dwParam3
        Close #FreeFileNo
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filter the call back message to an event '''''''''''''''''''''''''''''
    Select Case dwMsg
        Case LINE_ADDRESSSTATE
            RaiseEvent AddressState(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "AddressState" ' Set last Event sent
        Case LINE_CALLINFO
            RaiseEvent CallInfo(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "CallInfo" ' Set last Event sent
        Case LINE_CALLSTATE
            Select Case dwParam1
                Case LINECALLSTATE_IDLE
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Deallocate and remove the call now that it is finished'''''''''''''''''''''''''''''''''''''''''''''''''
                    Dim currentCallCounter&
                    
                    For currentCallCounter = 0 To Me.Devices.Item(Index).Calls.Count ' Cycle through call objects
                        If Module1.DeviceItemHndCall(Index, currentCallCounter) = hDevice Then ' Check which call object
                            Call Me.Devices.Item(Index).Calls.Item(currentCallCounter).Func_lineDealloc ' Deallocate call
                            Call DestroyCallObject(Index, currentCallCounter) ' Destroy call object
                            
                            Exit For ' Found the call object so no need to check the rest so exit the for loop
                        End If
                    Next ' Next call object in the collection
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    RaiseEvent IDLE(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3) ' Tell Parent Application
                    
                    Module1.LastEvent(Index) = "Idle" ' Set last Event sent
                Case LINECALLSTATE_OFFERING
                    Call CreateCallObject(Index, dwParam2)
                    
                    RaiseEvent OFFERING(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3) ' Tell Parent Application
                    
                    Module1.LastEvent(Index) = "Offering" ' Set last Event sent property
                Case LINECALLSTATE_ACCEPTED
                    RaiseEvent ACCEPTED(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Accepted" ' Set last Event sent
                Case LINECALLSTATE_DIALTONE
                    RaiseEvent DIALTONE(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Dialtone" ' Set last Event sent
                Case LINECALLSTATE_DIALING
                    RaiseEvent DIALING(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Dialing" ' Set last Event sent
                Case LINECALLSTATE_RINGBACK
                    RaiseEvent RINGBACK(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Ringback" ' Set last Event sent
                Case LINECALLSTATE_BUSY
                    RaiseEvent BUSY(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Buzy" ' Set last Event sent
                Case LINECALLSTATE_SPECIALINFO
                    RaiseEvent SpecialInfo(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Specialinfo" ' Set last Event sent
                Case LINECALLSTATE_CONNECTED
                    RaiseEvent Connected(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Connected" ' Set last Event sent
                Case LINECALLSTATE_PROCEEDING
                    RaiseEvent PROCEEDING(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Proceeding" ' Set last Event sent
                Case LINECALLSTATE_ONHOLD
                    RaiseEvent ONHOLD(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Onhold" ' Set last Event sent
                Case LINECALLSTATE_CONFERENCED
                    RaiseEvent CONFERENCED(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Conferenced" ' Set last Event sent
                Case LINECALLSTATE_ONHOLDPENDCONF
                    RaiseEvent ONHOLDPENDCONF(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Onholdpendconf"
                Case LINECALLSTATE_ONHOLDPENDTRANSFER
                    RaiseEvent ONHOLDPENDTRANSFER(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Onholdpendtransfer" ' Set last Event sent
                Case LINECALLSTATE_DISCONNECTED
                    RaiseEvent DISCONNECTED(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Disconnected" ' Set last Event sent
                Case LINECALLSTATE_UNKNOWN
                    RaiseEvent Unknown(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                    
                    Module1.LastEvent(Index) = "Unknown" ' Set last Event sent
            End Select
        Case LINE_REPLY
            If dwParam2 = 0 And Module1.LastFunction(Index) = "lineMakeCall" Then Call CreateCallObject(Index, dwParam3)
            
            RaiseEvent LINEREPLY(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "Reply" ' Set last Event sent
        Case LINE_CLOSE
            RaiseEvent lineClose(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "Close" ' Set last Event sent
        Case LINE_CREATE
            RaiseEvent CREATE(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "Create" ' Set last Event sent
        Case LINE_AGENTSPECIFIC
            RaiseEvent AGENTSPECIFIC(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "AgentSpecific" ' Set last Event sent
        Case LINE_AGENTSTATUS
            RaiseEvent AGENTSTATUS(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "AgentStatus" ' Set last Event sent
        Case LINE_APPNEWCALL
            Call CreateCallObject(Index, dwParam2)
            
            RaiseEvent APPNEWCALL(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "AppNewCall" ' Set last Event sent
        Case LINE_PROXYREQUEST
            RaiseEvent PROXYREQUEST(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "ProxyRequest" ' Set last Event sent
        Case LINE_REMOVE
            RaiseEvent Remove(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "Remove" ' Set last Event sent
        Case LINE_DEVSPECIFIC
            RaiseEvent devSPECIFIC(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "DevSpecific" ' Set last Event sent
        Case LINE_LINEDEVSTATE
            RaiseEvent DEVSTATE(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            
            Module1.LastEvent(Index) = "LineDevState" ' Set last Event sent
        Case LINE_REQUEST
            Module1.RequestMode = dwParam1 ' Set the request mode as global for use in the request object
            
            Set Request = Nothing ' Delete the last request object made if there is one
            Set Request = New RequestObj ' Create new request object
            
            Select Case dwParam1 ' Select request mode
                Case LINEREQUESTMODE_MAKECALL: RaiseEvent REQUESTMAKECALL(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                Case LINEREQUESTMODE_MEDIACALL: RaiseEvent REQUESTMEDIACALL(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
                Case LINEREQUESTMODE_DROP: RaiseEvent REQUESTDROP(Index, hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
            End Select
            
            Module1.LastEvent(Index) = "Request" ' Set last Event sent
    End Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub CreateCallObject(Index, CallNumber)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Create call object '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Let Module1.DeviceItem = Index ' Set the device item for creating the call object
    Let Module1.CallItem = Me.Devices.Item(Index).Calls.Count ' Get call item index
    Let Module1.DeviceItemHndCall(Index, Module1.CallItem) = CallNumber ' Set the call number for the new call object
    
    Call Me.Devices.Item(Index).Calls.Add(New CallObj, Me.Devices.Item(Index).Calls.Count) ' Add a new call object to the collection
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Debug to immediate window and to file ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Debug.Print "Created Call Object"; Tab(25); Index; Tab(45); CallNumber; Tab(55); Me.Devices.Item(Index).Calls.Count
    
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Created Call Object", Index, CallNumber, Me.Devices.Item(Index).Calls.Count
        Close #FreeFileNo
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub DestroyCallObject(Index, CallItem)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Destroy call object ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call Me.Devices.Item(Index).Calls.Remove(CallItem) ' Remove call object
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Debug to immediate window and to file ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Debug.Print "Removed Call Object"; Tab(25); Index; Tab(45); CallItem; Tab(55); Me.Devices.Item(Index).Calls.Count
    
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile ' Get file IO number
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo ' Open file
            Print #FreeFileNo, Now, "Removed Call Object", Index, CallItem, Me.Devices.Item(Index).Calls.Count ' Output to file
        Close #FreeFileNo ' Close file
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Public Function Func_lineRegisterRequestRecipient() As Long
Attribute Func_lineRegisterRequestRecipient.VB_Description = "Register to receive telephony requests."
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineRegisterRequestRecipient - Register"
        Close #FreeFileNo
    End If
    
    Func_lineRegisterRequestRecipient = lineRegisterRequestRecipient( _
        hndLine, _
        App.hInstance, _
        LINEREQUESTMODE_MAKECALL, _
        True _
    )
End Function

Public Function Func_lineUnregisterRequestRecipient() As Long
Attribute Func_lineUnregisterRequestRecipient.VB_Description = "Unregister to receive telephony requests."
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineRegisterRequestRecipient - Unregister"
        Close #FreeFileNo
    End If
    
    Func_lineUnregisterRequestRecipient = lineRegisterRequestRecipient( _
        hndLine, _
        App.hInstance, _
        LINEREQUESTMODE_MAKECALL, _
        False _
    )
End Function

Public Property Get Debugging() As Boolean
    Debugging = Module1.DebugMode
End Property

Public Property Let Debugging(x As Boolean)
    Module1.DebugMode = x
End Property
