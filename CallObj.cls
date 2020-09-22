VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is a call on the line device."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"CalledObj"
Attribute VB_Ext_KEY = "Member1" ,"CallerObj"
Attribute VB_Ext_KEY = "Member2" ,"ConnectedObj"
Attribute VB_Ext_KEY = "Member3" ,"RedirectingObj"
Attribute VB_Ext_KEY = "Member4" ,"RedirectionObj"
Option Explicit

Private Const LINEDIGITMODE_PULSE = &H1&
Private Const LINEDIGITMODE_DTMF = &H2&
Private Const LINEDIGITMODE_DTMFEND = &H4&

Private Const LINECALLINFO_FIXEDSIZE = 296

Private Type LINEDIALPARAMS
    dwDialPause As Long
    dwDialSpeed As Long
    dwDigitDuration As Long
    dwWaitForDialtone As Long
End Type

Private Type lineCallInfo
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    hLine As Long
    dwLineDeviceID As Long
    dwAddressID As Long

    dwBearerMode As Long
    dwRate As Long
    dwMediaMode As Long

    dwAppSpecific As Long
    dwCallID As Long
    dwRelatedCallID As Long
    dwCallParamFlags As Long
    dwCallStates As Long

    dwMonitorDigitModes As Long
    dwMonitorMediaModes As Long
    DialParams As LINEDIALPARAMS

    dwOrigin As Long
    dwReason As Long
    dwCompletionID As Long
    dwNumOwners As Long
    dwNumMonitors As Long

    dwCountryCode As Long
    dwTrunk As Long

    dwCallerIDFlags As Long
    dwCallerIDSize As Long
    dwCallerIDOffset As Long
    dwCallerIDNameSize As Long
    dwCallerIDNameOffset As Long

    dwCalledIDFlags As Long
    dwCalledIDSize As Long
    dwCalledIDOffset As Long
    dwCalledIDNameSize As Long
    dwCalledIDNameOffset As Long

    dwConnectedIDFlags As Long
    dwConnectedIDSize As Long
    dwConnectedIDOffset As Long
    dwConnectedIDNameSize As Long
    dwConnectedIDNameOffset As Long

    dwRedirectionIDFlags As Long
    dwRedirectionIDSize As Long
    dwRedirectionIDOffset As Long
    dwRedirectionIDNameSize As Long
    dwRedirectionIDNameOffset As Long

    dwRedirectingIDFlags As Long
    dwRedirectingIDSize As Long
    dwRedirectingIDOffset As Long
    dwRedirectingIDNameSize As Long
    dwRedirectingIDNameOffset As Long

    dwAppNameSize As Long
    dwAppNameOffset As Long

    dwDisplayableAddressSize As Long
    dwDisplayableAddressOffset As Long

    dwCalledPartySize As Long
    dwCalledPartyOffset As Long

    dwCommentSize As Long
    dwCommentOffset As Long

    dwDisplaySize As Long
    dwDisplayOffset As Long

    dwUserUserInfoSize As Long
    dwUserUserInfoOffset As Long

    dwHighLevelCompSize As Long
    dwHighLevelCompOffset As Long

    dwLowLevelCompSize As Long
    dwLowLevelCompOffset As Long

    dwChargingInfoSize As Long
    dwChargingInfoOffset As Long

    dwTerminalModesSize As Long
    dwTerminalModesOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    
    mem As String * 2048 ' added by mca
End Type

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

Private Type LINECALLSTATUS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwCallState As Long
    dwCallStateMode As Long
    dwCallPrivilege As Long
    dwCallFeatures As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long

    mem As String * 2048
End Type

Private Const LINEDEVCAPS_FIXEDSIZE = 236
Private Const LINECALLSTATUS_FIXEDSIZE = 36

Private Declare Function lineGetCallStatus Lib "TAPI32.DLL" (ByVal hCall As Long, lpCallStatus As Any) As Long
Private Declare Function lineGetCallInfo Lib "TAPI32.DLL" (ByVal hCall As Long, ByRef lpCallInfo As lineCallInfo) As Long
Private Declare Function lineSecureCall Lib "TAPI32.DLL" (ByVal hCall As Long) As Long
Private Declare Function lineRedirect Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal lpszDestAddress As String, ByVal dwCountryCode As Long) As Long
Private Declare Function lineHold Lib "TAPI32.DLL" (ByVal hCall As Long) As Long
Private Declare Function lineUnhold Lib "TAPI32.DLL" (ByVal hCall As Long) As Long
Private Declare Function lineGenerateDigits Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwDigitMode As Long, ByVal lpszDigits As String, ByVal dwDuration As Long) As Long
Private Declare Function lineAnswer Lib "TAPI32.DLL" (ByVal hCall As Long, lpsUserUserInfo As Any, ByVal dwSize As Long) As Long
Private Declare Function lineDeallocateCall Lib "TAPI32.DLL" (ByVal hCall As Long) As Long
Private Declare Function lineDrop Lib "TAPI32.DLL" (ByVal hCall As Long, lpsUserUserInfo As Any, ByVal dwSize As Long) As Long
Private Declare Function lineAccept Lib "TAPI32.DLL" (ByVal hCall As Long, lpsUserUserInfo As Any, ByVal dwSize As Long) As Long
Private Declare Function lineMonitorDigits Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwDigitModes As Long) As Long
Private Declare Function lineMonitorMedia Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwMediaModes As Long) As Long

' To implement
Private Declare Function lineSwapHold Lib "TAPI32.DLL" (ByVal hActiveCall As Long, ByVal hHeldCall As Long) As Long
Private Declare Function linePickup Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, lphCall As Long, ByVal lpszDestAddress As String, ByVal lpszGroupID As String) As Long
Private Declare Function linePark Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwParkMode As Long, ByVal lpszDirAddress As String, lpNonDirAddress As Any) As Long
Private Declare Function lineUnpark Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, lphCall As Long, ByVal lpszDestAddress As String) As Long
Private Declare Function linePrepareAddToConference Lib "TAPI32.DLL" (ByVal hConfCall As Long, lphConsultCall As Long, lpCallParams As Any) As Long
Private Declare Function lineRemoveFromConference Lib "TAPI32.DLL" (ByVal hCall As Long) As Long
Private Declare Function lineGetConfRelatedCalls Lib "TAPI32.DLL" (ByVal hCall As Long, lpCallList As Any) As Long
Private Declare Function lineAddToConference Lib "TAPI32.DLL" (ByVal hConfCall As Long, ByVal hConsultCall As Long) As Long
Private Declare Function lineSetupConference Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal hLine As Long, lphConfCall As Long, lphConsultCall As Long, ByVal dwNumParties As Long, lpCallParams As Any) As Long
Private Declare Function lineUncompleteCall Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwCompletionID As Long) As Long
Private Declare Function lineCompleteCall Lib "TAPI32.DLL" (ByVal hCall As Long, lpdwCompletionID As Long, ByVal dwCompletionMode As Long, ByVal dwMessageID As Long) As Long
Private Declare Function lineSetTerminal Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal hCall As Long, ByVal dwSelect As Long, ByVal dwTerminalModes As Long, ByVal dwTerminalID As Long, ByVal bEnable As Long) As Long
Private Declare Function lineSetMediaControl Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal hCall As Long, ByVal dwSelect As Long, lpDigitList As Any, ByVal dwDigitNumEntries As Long, lpMediaList As Any, ByVal dwMediaNumEntries As Long, lpToneList As Any, ByVal dwToneNumEntries As Long, lpCallStateList As Any, ByVal dwCallStateNumEntries As Long) As Long
Private Declare Function lineSetMediaMode Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwMediaModes As Long) As Long
Private Declare Function lineSetAppSpecific Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwAppSpecific As Long) As Long
Private Declare Function lineSetCallParams Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwBearerMode As Long, ByVal dwMinRate As Long, ByVal dwMaxRate As Long, lpDialParams As LINEDIALPARAMS) As Long
Private Declare Function lineSetCallPrivilege Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwCallPrivilege As Long) As Long
Private Declare Function lineHandoff Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal lpszFileName As String, ByVal dwMediaMode As Long) As Long
Private Declare Function lineForward Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal bAllAddresses As Long, ByVal dwAddressID As Long, lpForwardList As Any, ByVal dwNumRingsNoAnswer As Long, lphConsultCall As Long, lpCallParams As Any) As Long
Private Declare Function lineBlindTransfer Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal lpszDestAddress As String, ByVal dwCountryCode As Long) As Long
Private Declare Function lineSetupTransfer Lib "TAPI32.DLL" (ByVal hCall As Long, lphConsultCall As Long, lpCallParams As Any) As Long
Private Declare Function lineCompleteTransfer Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal hConsultCall As Long, lphConfCall As Long, ByVal dwTransferMode As Long) As Long

Dim Index As Long
Dim CallIndex As Long

Private mvarRedirectionObj As RedirectionObj
Private mvarRedirectingObj As RedirectingObj
Private mvarConnectedObj As ConnectedObj
Private mvarCalledObj As CalledObj
Private mvarCallerObj As CallerObj

Public Property Get Caller() As CallerObj
    If mvarCallerObj Is Nothing Then Set mvarCallerObj = New CallerObj
    Set Caller = mvarCallerObj
End Property

Public Property Get Called() As CalledObj
    If mvarCalledObj Is Nothing Then Set mvarCalledObj = New CalledObj
    Set Called = mvarCalledObj
End Property

Public Property Get Connected() As ConnectedObj
    If mvarConnectedObj Is Nothing Then Set mvarConnectedObj = New ConnectedObj
    Set Connected = mvarConnectedObj
End Property

Public Property Get Redirecting() As RedirectingObj
    If mvarRedirectingObj Is Nothing Then Set mvarRedirectingObj = New RedirectingObj
    Set Redirecting = mvarRedirectingObj
End Property

Public Property Get Redirection() As RedirectionObj
    If mvarRedirectionObj Is Nothing Then Set mvarRedirectionObj = New RedirectionObj
    Set Redirection = mvarRedirectionObj
End Property

Private Sub Class_Terminate()
    Set mvarCallerObj = Nothing
    Set mvarCalledObj = Nothing
    Set mvarConnectedObj = Nothing
    Set mvarRedirectingObj = Nothing
    Set mvarRedirectionObj = Nothing
End Sub

Private Sub Class_Initialize()
    Index = DeviceItem
    CallIndex = CallItem
    Set mvarCallerObj = New CallerObj
    Set mvarCalledObj = New CalledObj
    Set mvarConnectedObj = New ConnectedObj
    Set mvarRedirectingObj = New RedirectingObj
    Set mvarRedirectionObj = New RedirectionObj
End Sub

Public Property Get State() As String
    Dim lineCallState As LINECALLSTATUS, retVal As Long

    lineCallState.dwTotalSize = Len(lineCallState)
    retVal = lineGetCallStatus(DeviceItemHndCall(Index, CallIndex), lineCallState)

    State = TAPI_GETCALLSTATE(lineCallState.dwCallState)
End Property

Public Property Get StateMode() As Long
    Dim lineCallState As LINECALLSTATUS, retVal As Long

    lineCallState.dwTotalSize = Len(lineCallState)
    retVal = lineGetCallStatus(DeviceItemHndCall(Index, CallIndex), lineCallState)

    StateMode = lineCallState.dwCallStateMode
End Property

Public Property Get Privilege() As Long
    Dim lineCallState As LINECALLSTATUS, retVal As Long

    lineCallState.dwTotalSize = Len(lineCallState)
    retVal = lineGetCallStatus(DeviceItemHndCall(Index, CallIndex), lineCallState)

    Privilege = lineCallState.dwCallPrivilege
End Property

Public Property Get Features() As Long
    Dim lineCallState As LINECALLSTATUS, retVal As Long

    lineCallState.dwTotalSize = Len(lineCallState)
    retVal = lineGetCallStatus(DeviceItemHndCall(Index, CallIndex), lineCallState)

    Features = lineCallState.dwCallFeatures
End Property

Public Property Get CallStatusSPECIFIC() As String
    Dim lineCallState As LINECALLSTATUS, retVal As Long

    lineCallState.dwTotalSize = Len(lineCallState)
    retVal = lineGetCallStatus(DeviceItemHndCall(Index, CallIndex), lineCallState)

    If lineCallState.dwDevSpecificSize <> 0 Then
        CallStatusSPECIFIC = GetCallStatusString(lineCallState.mem, lineCallState.dwDevSpecificOffset, lineCallState.dwDevSpecificSize)
    End If
End Property

Private Function GetCallStatusString(mem As String, offset As Long, size As Long) As String
    GetCallStatusString = Trim(Replace(Replace(Mid(mem, offset + 1 - LINECALLSTATUS_FIXEDSIZE, size - 1), Chr(0), " "), "|", " "))
End Function

Private Function TAPI_GETCALLSTATE(CALL_STATE As Long) As String
    Select Case CALL_STATE
        Case LINECALLSTATE_IDLE: TAPI_GETCALLSTATE = "IDLE"
        Case LINECALLSTATE_OFFERING: TAPI_GETCALLSTATE = "OFFERING"
        Case LINECALLSTATE_ACCEPTED: TAPI_GETCALLSTATE = "ACCEPTED"
        Case LINECALLSTATE_DIALTONE: TAPI_GETCALLSTATE = "DIALTONE"
        Case LINECALLSTATE_DIALING: TAPI_GETCALLSTATE = "DIALING"
        Case LINECALLSTATE_RINGBACK: TAPI_GETCALLSTATE = "RINGBACK"
        Case LINECALLSTATE_BUSY: TAPI_GETCALLSTATE = "BUSY"
        Case LINECALLSTATE_SPECIALINFO: TAPI_GETCALLSTATE = "SPECIALINFO"
        Case LINECALLSTATE_CONNECTED: TAPI_GETCALLSTATE = "CONNECTED"
        Case LINECALLSTATE_PROCEEDING: TAPI_GETCALLSTATE = "PROCEDING"
        Case LINECALLSTATE_ONHOLD: TAPI_GETCALLSTATE = "ON HOLD"
        Case LINECALLSTATE_CONFERENCED: TAPI_GETCALLSTATE = "CONFERENCED"
        Case LINECALLSTATE_ONHOLDPENDCONF: TAPI_GETCALLSTATE = "ON HOLD PEND CONF"
        Case LINECALLSTATE_ONHOLDPENDTRANSFER: TAPI_GETCALLSTATE = "ON HOLD PENDING TRANSFER"
        Case LINECALLSTATE_DISCONNECTED: TAPI_GETCALLSTATE = "DISCONNECTED"
        Case Else: TAPI_GETCALLSTATE = "UNKNOWN"
    End Select
End Function

Public Function Func_lineAccept(Optional ByVal UserUserInfo As String) As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineAccept"
        Close #FreeFileNo
    End If
    
    Func_lineAccept = lineAccept(DeviceItemHndCall(Index, CallIndex), UserUserInfo, Len(UserUserInfo))
    LastFunction(Index) = "lineAccept"
End Function

Public Function Func_lineRedirect(ByVal DestAddress As String, CountryCode As Long)
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineRedirect"
        Close #FreeFileNo
    End If
    
    Func_lineRedirect = lineRedirect(DeviceItemHndCall(Index, CallIndex), DestAddress, CountryCode)
    LastFunction(Index) = "lineRedirect"
End Function

Public Function Func_lineSecureCall() As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineSecureCall"
        Close #FreeFileNo
    End If
    
    Func_lineSecureCall = lineSecureCall(DeviceItemHndCall(Index, CallIndex))
    LastFunction(Index) = "lineSecureCall"
End Function

Public Function Func_lineHold() As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineHold"
        Close #FreeFileNo
    End If
    
    Func_lineHold = lineHold(DeviceItemHndCall(Index, CallIndex))
    LastFunction(Index) = "lineHold"
End Function

Public Function Func_lineUnhold() As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineUnhold"
        Close #FreeFileNo
    End If
    
    Func_lineUnhold = lineUnhold(DeviceItemHndCall(Index))
    LastFunction(Index) = "lineUnhold"
End Function

Public Function Func_lineGenerateDigits(ByVal Mode As Long, ByVal digit As String, ByRef duration As Long) As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineGenerateDigits"
        Close #FreeFileNo
    End If
    
    Func_lineGenerateDigits = lineGenerateDigits(DeviceItemHndCall(Index, CallIndex), Mode, digit, duration)
    LastFunction(Index) = "lineGenerateDigits"
End Function

Public Function Func_lineAnswer(Optional ByVal lpsUserUserInfo As String) As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineAnswer"
        Close #FreeFileNo
    End If
    
    Func_lineAnswer = lineAnswer( _
        DeviceItemHndCall(Index, CallIndex), _
        lpsUserUserInfo, _
        Len(lpsUserUserInfo) _
    )
    
    LastFunction(Index) = "lineAnswer"
End Function

Public Function Func_lineDealloc() As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineDeallocateCall"
        Close #FreeFileNo
    End If
    
    Func_lineDealloc = lineDeallocateCall( _
        DeviceItemHndCall(Index, CallIndex) _
    )
    
    LastFunction(Index) = "lineDeallocateCall"
End Function

Public Function Func_LineDrop(Optional ByVal lpsUserUserInfo As String) As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineDrop"
        Close #FreeFileNo
    End If
    
    Func_LineDrop = lineDrop( _
        DeviceItemHndCall(Index, CallIndex), _
        lpsUserUserInfo, _
        Len(lpsUserUserInfo) _
    )
    
    LastFunction(Index) = "lineDrop"
End Function

Public Property Get CallParamFlags() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    CallParamFlags = TAPI_GETCALLPARAM(lineCallInformation.dwCallParamFlags)
End Property

Public Property Get RelatedCallID() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    RelatedCallID = lineCallInformation.dwRelatedCallID
End Property

Public Property Get CallID() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    CallID = lineCallInformation.dwCallID
End Property

Public Property Get Origin() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    Origin = TAPI_GETCALLORIGIN(lineCallInformation.dwOrigin)
End Property

Public Property Get DeviceID() As Long
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    DeviceID = lineCallInformation.dwLineDeviceID
End Property

Public Property Get AddressID() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)
    AddressID = lineCallInformation.dwAddressID
End Property

Public Property Get Rate() As Long
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    Rate = lineCallInformation.dwRate
End Property

Public Property Get ID() As Long
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)
    ID = lineCallInformation.dwCallID
End Property

Public Property Get CountryCode() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    CountryCode = lineCallInformation.dwCountryCode
End Property

Public Property Get Trunk() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    Trunk = lineCallInformation.dwTrunk
End Property

Public Property Get AppName() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwAppNameSize <> 0 Then
        AppName = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwAppNameOffset, lineCallInformation.dwAppNameSize)
    End If
End Property

Public Property Get DisplayableAddress() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwDisplayableAddressSize <> 0 Then
        DisplayableAddress = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwDisplayableAddressOffset, lineCallInformation.dwDisplayableAddressSize)
    End If
End Property

Public Property Get CalledParty() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwCalledPartySize <> 0 Then
        CalledParty = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwCalledPartyOffset, lineCallInformation.dwCalledPartySize)
    End If
End Property

Public Property Get Comment() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwCommentSize <> 0 Then
        Comment = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwCommentOffset, lineCallInformation.dwCommentSize)
    End If
End Property

Public Property Get UserUserInfo() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwUserUserInfoSize <> 0 Then
        UserUserInfo = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwUserUserInfoOffset, lineCallInformation.dwUserUserInfoSize)
    End If
End Property

Public Property Get HighLevelComp() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwHighLevelCompSize <> 0 Then
        HighLevelComp = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwHighLevelCompOffset, lineCallInformation.dwHighLevelCompSize)
    End If
End Property

Public Property Get LowLevelComp() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwLowLevelCompSize <> 0 Then
        LowLevelComp = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwLowLevelCompOffset, lineCallInformation.dwLowLevelCompSize)
    End If
End Property

Public Property Get ChargingInfo() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwChargingInfoSize <> 0 Then
        ChargingInfo = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwChargingInfoOffset, lineCallInformation.dwChargingInfoSize)
    End If
End Property

Public Property Get TerminalModes() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwTerminalModesSize <> 0 Then
        TerminalModes = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwTerminalModesOffset, lineCallInformation.dwTerminalModesSize)
    End If
End Property

Public Property Get CallInfoSPECIFIC() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwDevSpecificSize <> 0 Then
        CallInfoSPECIFIC = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwDevSpecificOffset, lineCallInformation.dwDevSpecificSize)
    End If
End Property

Private Function GetCallInfoString(mem As String, offset As Long, size As Long) As String
    GetCallInfoString = Trim(Replace(Replace(Mid(mem, offset + 1 - LINECALLINFO_FIXEDSIZE, size - 1), Chr(0), " "), "|", " "))
End Function

Public Function Func_lineMonitorDigits() As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineMonitorDigits"
        Close #FreeFileNo
    End If
    
    Func_lineMonitorDigits = lineMonitorDigits(DeviceItemHndCall(Index, CallIndex), LINEDIGITMODE_DTMF)

    Let Module1.LastFunction(Index) = "lineMonitorDigits"
End Function

Public Function Func_lineMonitorMedia(ByVal Mode As Long) As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineMonitorMedia"
        Close #FreeFileNo
    End If
    
    Func_lineMonitorMedia = lineMonitorMedia(DeviceItemHndCall(Index, CallIndex), Mode)
    
    Let Module1.LastFunction(Index) = "lineMonitorMedia"
End Function
