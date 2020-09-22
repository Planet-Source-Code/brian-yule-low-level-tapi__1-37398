VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceItemObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is a  line device"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Addresses"
Attribute VB_Ext_KEY = "Member1" ,"CallsObj"
Attribute VB_Ext_KEY = "Member2" ,"DeviceItemConfigObj"
Option Explicit

Private Const LINEDEVCAPS_FIXEDSIZE = 236
Private Const LINEDEVSTATUS_FIXEDSIZE = 76

Private Const TAPI_LOWAPI = &H10003
Private Const TAPI_HIGHAPI = &H20001

Private Const LINEADDRESSMODE_ADDRESSID = &H1&
Private Const LINEADDRESSMODE_DIALABLEADDR = &H2&

Private Const STRINGFORMAT_ASCII = 1&
Private Const STRINGFORMAT_DBCS = 2&
Private Const STRINGFORMAT_UNICODE = 3&
Private Const STRINGFORMAT_BINARY = 4&

Private Const LINEDEVCAPFLAGS_CROSSADDRCONF = &H1
Private Const LINEDEVCAPFLAGS_HIGHLEVCOMP = &H2
Private Const LINEDEVCAPFLAGS_LOWLEVCOMP = &H4
Private Const LINEDEVCAPFLAGS_MEDIACONTROL = &H8
Private Const LINEDEVCAPFLAGS_MULTIPLEADDR = &H10
Private Const LINEDEVCAPFLAGS_CLOSEDROP = &H20
Private Const LINEDEVCAPFLAGS_DIALBILLING = &H40
Private Const LINEDEVCAPFLAGS_DIALQUIET = &H80
Private Const LINEDEVCAPFLAGS_DIALDIALTONE = &H100

Private Const LINEBEARERMODE_VOICE = &H1&
Private Const LINEBEARERMODE_SPEECH = &H2&
Private Const LINEBEARERMODE_MULTIUSE = &H4&
Private Const LINEBEARERMODE_DATA = &H8&
Private Const LINEBEARERMODE_ALTSPEECHDATA = &H10&
Private Const LINEBEARERMODE_NONCALLSIGNALING = &H20&

Private Const LINEMEDIAMODE_UNKNOWN = &H2&
Private Const LINEMEDIAMODE_INTERACTIVEVOICE = &H4&
Private Const LINEMEDIAMODE_AUTOMATEDVOICE = &H8&
Private Const LINEMEDIAMODE_DATAMODEM = &H10&
Private Const LINEMEDIAMODE_G3FAX = &H20&
Private Const LINEMEDIAMODE_TDD = &H40&
Private Const LINEMEDIAMODE_G4FAX = &H80&
Private Const LINEMEDIAMODE_DIGITALDATA = &H100&
Private Const LINEMEDIAMODE_TELETEX = &H200&
Private Const LINEMEDIAMODE_VIDEOTEX = &H400&
Private Const LINEMEDIAMODE_TELEX = &H800&
Private Const LINEMEDIAMODE_MIXED = &H1000&
Private Const LINEMEDIAMODE_ADSI = &H2000&

Private Const LINEDIGITMODE_PULSE = &H1&
Private Const LINEDIGITMODE_DTMF = &H2&
Private Const LINEDIGITMODE_DTMFEND = &H4&

Private Const LINEANSWERMODE_NONE = &H1&
Private Const LINEANSWERMODE_DROP = &H2&
Private Const LINEANSWERMODE_HOLD = &H4&

Private Const LINETONEMODE_CUSTOM = &H1&
Private Const LINETONEMODE_RINGBACK = &H2&
Private Const LINETONEMODE_BUSY = &H4&
Private Const LINETONEMODE_BEEP = &H8&
Private Const LINETONEMODE_BILLING = &H10&

Private Type LINEMONITORTONE
    dwAppSpecific As Long
    dwDuration As Long
    dwFrequency1 As Long
    dwFrequency2 As Long
    dwFrequency3 As Long
End Type

Private Type LINEDIALPARAMS
    dwDialPause As Long
    dwDialSpeed As Long
    dwDigitDuration As Long
    dwWaitForDialtone As Long
End Type

Private Type LINECALLPARAMS
    dwTotalSize As Long

    dwBearerMode As Long
    dwMinRate As Long
    dwMaxRate As Long
    dwMediaMode As Long

    dwCallParamFlags As Long
    dwAddressMode As Long
    dwAddressID As Long

    DialParams As LINEDIALPARAMS

    dwOrigAddressSize As Long
    dwOrigAddressOffset As Long

    dwDisplayableAddressSize As Long
    dwDisplayableAddressOffset As Long

    dwCalledPartySize As Long
    dwCalledPartyOffset As Long

    dwCommentSize As Long
    dwCommentOffset As Long

    dwUserUserInfoSize As Long
    dwUserUserInfoOffset As Long

    dwHighLevelCompSize As Long
    dwHighLevelCompOffset As Long

    dwLowLevelCompSize As Long
    dwLowLevelCompOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    '
    mem As String * 2048 ' added by mca
End Type

Private Type LINEDEVCAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwProviderInfoSize As Long
    dwProviderInfoOffset As Long

    dwSwitchInfoSize As Long
    dwSwitchInfoOffset As Long

    dwPermanentLineID As Long
    dwLineNameSize As Long
    dwLineNameOffset As Long
    dwStringFormat As Long

    dwAddressModes As Long
    dwNumAddresses As Long
    dwBearerModes As Long
    dwMaxRate As Long
    dwMediaModes As Long

    dwGenerateToneModes As Long
    dwGenerateToneMaxNumFreq As Long
    dwGenerateDigitModes As Long
    dwMonitorToneMaxNumFreq As Long
    dwMonitorToneMaxNumEntries As Long
    dwMonitorDigitModes As Long
    dwGatherDigitsMinTimeout As Long
    dwGatherDigitsMaxTimeout As Long

    dwMedCtlDigitMaxListSize As Long
    dwMedCtlMediaMaxListSize As Long
    dwMedCtlToneMaxListSize As Long
    dwMedCtlCallStateMaxListSize As Long

    dwDevCapFlags As Long
    dwMaxNumActiveCalls As Long
    dwAnswerMode As Long
    dwRingModes As Long
    dwLineStates As Long

    dwUUIAcceptSize As Long
    dwUUIAnswerSize As Long
    dwUUIMakeCallSize As Long
    dwUUIDropSize As Long
    dwUUISendUserInfoSize As Long
    dwUUICallInfoSize As Long

    MinDialParams As LINEDIALPARAMS
    MaxDialParams As LINEDIALPARAMS
    DefaultDialParams As LINEDIALPARAMS

    dwNumTerminals As Long
    dwTerminalCapsSize As Long
    dwTerminalCapsOffset As Long
    dwTerminalTextEntrySize As Long
    dwTerminalTextSize As Long
    dwTerminalTextOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    
    mem As String * 2048
End Type

Private Type LPPROVIDERLIST
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    
    dwNumProviders As Long
    dwProviderListSize As Long
    dwProviderListOffset As Long

    mem As String * 2048
End Type

Private Type LINEADDRESSCAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwLineDeviceID As Long

    dwAddressSize As Long
    dwAddressOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long

    dwAddressSharing As Long
    dwAddressStates As Long
    dwCallInfoStates As Long
    dwCallerIDFlags As Long
    dwCalledIDFlags As Long
    dwConnectedIDFlags As Long
    dwRedirectionIDFlags As Long
    dwRedirectingIDFlags As Long
    dwCallStates As Long
    dwDialToneModes As Long
    dwBusyModes As Long
    dwSpecialInfo As Long
    dwDisconnectModes As Long

    dwMaxNumActiveCalls As Long
    dwMaxNumOnHoldCalls As Long
    dwMaxNumOnHoldPendingCalls As Long
    dwMaxNumConference As Long
    dwMaxNumTransConf As Long

    dwAddrCapFlags As Long
    dwCallFeatures As Long
    dwRemoveFromConfCaps As Long
    dwRemoveFromConfState As Long
    dwTransferModes As Long
    dwParkModes As Long

    dwForwardModes As Long
    dwMaxForwardEntries As Long
    dwMaxSpecificEntries As Long
    dwMinFwdNumRings As Long
    dwMaxFwdNumRings As Long

    dwMaxCallCompletions As Long
    dwCallCompletionConds As Long
    dwCallCompletionModes As Long
    dwNumCompletionMessages As Long
    dwCompletionMsgTextEntrySize As Long
    dwCompletionMsgTextSize As Long
    dwCompletionMsgTextOffset As Long
    
    dwPredictiveAutoTransferStates As Long
    dwNumCallTreatments As Long
    dwCallTreatmentListSize As Long
    dwCallTreatmentListOffset As Long
    dwDeviceClassesSize As Long
    dwDeviceClassesOffset As Long
    dwMaxCallDataSize As Long
    dwCallFeatures2 As Long
    dwMaxNoAnswerTimeout As Long
    dwConnectedModes As Long
    dwOfferingModes As Long
    dwAvailableMediaModes As Long

    mem As String * 2048
End Type

Private Type LINEEXTENSIONID
    dwExtensionID0 As Long
    dwExtensionID1 As Long
    dwExtensionID2 As Long
    dwExtensionID3 As Long
End Type

Private Type LINETRANSLATECAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwNumLocations As Long
    dwLocationListSize As Long
    dwLocationListOffset As Long

    dwCurrentLocationID As Long

    dwNumCards As Long
    dwCardListSize As Long
    dwCardListOffset As Long

    dwCurrentPreferredCardID As Long

    mem As String * 2048
    
End Type


Private Declare Function lineGetDevCaps Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion As Long, lpLineDevCaps As LINEDEVCAPS) As Long
Private Declare Function LineOpen Lib "TAPI32.DLL" Alias "lineOpen" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByRef lphLine As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion As Long, ByVal dwCallbackInstance As Long, ByVal dwPrivileges As Long, ByVal dwMediaModes As Long, lpCallParams As Any) As Long
Private Declare Function lineClose Lib "TAPI32.DLL" (ByVal hLine As Long) As Long
Private Declare Function lineMakeCall Lib "TAPI32.DLL" (ByVal hLine As Long, ByRef lphCall As Long, ByVal lpszDestAddress As String, ByVal dwCountryCode As Long, lpCallParams As Any) As Long

Private Declare Function lineGetStatusMessages Lib "TAPI32.DLL" (ByVal hLine As Long, lpdwLineStates As Long, lpdwAddressStates As Long) As Long

Private Declare Function lineNegotiateAPIVersion Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPILowVersion As Long, ByVal dwAPIHighVersion As Long, lpdwAPIVersion As Long, lpExtensionID As LINEEXTENSIONID) As Long
Private Declare Function lineGetTranslateCaps Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwAPIVersion As Long, lpTranslateCaps As Any) As Long
'Private Declare Function lineGetProviderList Lib "TAPI32.DLL" (dwAPIVersion As Long, LPLINEPROVIDERLIST As LPPROVIDERLIST) As Long
'
'Private Declare Function lineMonitorTones Lib "TAPI32.DLL" (ByVal hCall As Long, lpToneList As LINEMONITORTONE, ByVal dwNumEntries As Long) As Long

Private Declare Function lineGetRequest Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwRequestMode As Long, lpRequestBuffer As Any) As Long

Private Type LINEDEVSTATUS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwNumOpens As Long
    dwOpenMediaModes As Long
    dwNumActiveCalls As Long
    dwNumOnHoldCalls As Long
    dwNumOnHoldPendCalls As Long
    dwLineFeatures As Long
    dwNumCallCompletions As Long
    dwRingMode As Long
    dwSignalLevel As Long
    dwBatteryLevel As Long
    dwRoamMode As Long

    dwDevStatusFlags As Long

    dwTerminalModesSize As Long
    dwTerminalModesOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    
    mem As String * 2048
End Type

Private Declare Function lineGetLineDevStatus Lib "TAPI32.DLL" (ByVal hLine As Long, lpLineDevStatus As LINEDEVSTATUS) As Long

Private Const LINEDEVSTATE_OTHER = &H1&
Private Const LINEDEVSTATE_RINGING = &H2&
Private Const LINEDEVSTATE_CONNECTED = &H4&
Private Const LINEDEVSTATE_DISCONNECTED = &H8&
Private Const LINEDEVSTATE_MSGWAITON = &H10&
Private Const LINEDEVSTATE_MSGWAITOFF = &H20&
Private Const LINEDEVSTATE_INSERVICE = &H40&
Private Const LINEDEVSTATE_OUTOFSERVICE = &H80&
Private Const LINEDEVSTATE_MAINTENANCE = &H100&
Private Const LINEDEVSTATE_OPEN = &H200&
Private Const LINEDEVSTATE_CLOSE = &H400&
Private Const LINEDEVSTATE_NUMCALLS = &H800&
Private Const LINEDEVSTATE_NUMCOMPLETIONS = &H1000&
Private Const LINEDEVSTATE_TERMINALS = &H2000&
Private Const LINEDEVSTATE_ROAMMODE = &H4000&
Private Const LINEDEVSTATE_BATTERY = &H8000&
Private Const LINEDEVSTATE_SIGNAL = &H10000
Private Const LINEDEVSTATE_DEVSPECIFIC = &H20000
Private Const LINEDEVSTATE_REINIT = &H40000
Private Const LINEDEVSTATE_LOCK = &H80000
Private Const LINEDEVSTATE_CAPSCHANGE = &H100000
Private Const LINEDEVSTATE_CONFIGCHANGE = &H200000
Private Const LINEDEVSTATE_TRANSLATECHANGE = &H400000
Private Const LINEDEVSTATE_COMPLCANCEL = &H800000
Private Const LINEDEVSTATE_REMOVED = &H1000000

Private Declare Function lineSetStatusMessages Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwLineStates As Long, ByVal dwAddressStates As Long) As Long

Dim Index As Long
Dim hndOpenLine As Long

Public BearerModes As New Collection
Public AddressModes As New Collection
Public MediaModes As New Collection
Public GenerateToneModes As New Collection
Public GenerateDigitModes As New Collection
Public DeviceModes As New Collection
Public LineStates As New Collection
Public MonitorDigitModes As New Collection

Private mvarDeviceItemConfigObj As DeviceItemConfigObj
Private mvarAddresses As Addresses
Private mvarCallsObj As CallsObj

Private Sub LoadGenerateDigitModes()
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If LINEDIGITMODE_PULSE And lineDevCap.dwGenerateDigitModes Then GenerateDigitModes.Add "ADDRESSID", CStr(GenerateDigitModes.Count)
    If LINEDIGITMODE_DTMF And lineDevCap.dwGenerateDigitModes Then GenerateDigitModes.Add "DIALABLEADDR", CStr(GenerateDigitModes.Count)
End Sub

Private Sub LoadMonitorDigitModes()
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If LINEDIGITMODE_PULSE And lineDevCap.dwGenerateDigitModes Then MonitorDigitModes.Add "ADDRESSID", CStr(MonitorDigitModes.Count)
    If LINEDIGITMODE_DTMF And lineDevCap.dwGenerateDigitModes Then MonitorDigitModes.Add "DIALABLEADDR", CStr(MonitorDigitModes.Count)
    If LINEDIGITMODE_DTMFEND And lineDevCap.dwGenerateDigitModes Then MonitorDigitModes.Add "DIALABLEADDR", CStr(MonitorDigitModes.Count)
End Sub

Private Sub LoadLineStates()
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If LINEDEVSTATE_OTHER And lineDevCap.dwLineStates Then LineStates.Add "OTHER", CStr(LineStates.Count)
    If LINEDEVSTATE_RINGING And lineDevCap.dwDevCapFlags Then LineStates.Add "RINGING", CStr(LineStates.Count)
    If LINEDEVSTATE_CONNECTED And lineDevCap.dwDevCapFlags Then LineStates.Add "CONNECTED", CStr(LineStates.Count)
    If LINEDEVSTATE_DISCONNECTED And lineDevCap.dwDevCapFlags Then LineStates.Add "DISCONNECTED", CStr(LineStates.Count)
    If LINEDEVSTATE_MSGWAITON And lineDevCap.dwDevCapFlags Then LineStates.Add "MSGWAITON", CStr(LineStates.Count)
    If LINEDEVSTATE_MSGWAITOFF And lineDevCap.dwDevCapFlags Then LineStates.Add "MSGWAITOFF", CStr(LineStates.Count)
    If LINEDEVSTATE_NUMCOMPLETIONS And lineDevCap.dwDevCapFlags Then LineStates.Add "NUMCOMPLETIONS", CStr(LineStates.Count)
    If LINEDEVSTATE_INSERVICE And lineDevCap.dwDevCapFlags Then LineStates.Add "INSERVICE", CStr(LineStates.Count)
    If LINEDEVSTATE_OUTOFSERVICE And lineDevCap.dwDevCapFlags Then LineStates.Add "OUTOFSERVICE", CStr(LineStates.Count)
    If LINEDEVSTATE_MAINTENANCE And lineDevCap.dwDevCapFlags Then LineStates.Add "MAINTENANCE", CStr(LineStates.Count)
    If LINEDEVSTATE_OPEN And lineDevCap.dwDevCapFlags Then LineStates.Add "OPEN", CStr(LineStates.Count)
    If LINEDEVSTATE_CLOSE And lineDevCap.dwDevCapFlags Then LineStates.Add "CLOSE", CStr(LineStates.Count)
    If LINEDEVSTATE_NUMCALLS And lineDevCap.dwDevCapFlags Then LineStates.Add "NUMCALLS", CStr(LineStates.Count)
    If LINEDEVSTATE_TERMINALS And lineDevCap.dwDevCapFlags Then LineStates.Add "TERMINALS", CStr(LineStates.Count)
    If LINEDEVSTATE_ROAMMODE And lineDevCap.dwDevCapFlags Then LineStates.Add "ROAMMODE", CStr(LineStates.Count)
    If LINEDEVSTATE_BATTERY And lineDevCap.dwDevCapFlags Then LineStates.Add "BATTERY", CStr(LineStates.Count)
    If LINEDEVSTATE_SIGNAL And lineDevCap.dwDevCapFlags Then LineStates.Add "SIGNAL", CStr(LineStates.Count)
    If LINEDEVSTATE_DEVSPECIFIC And lineDevCap.dwDevCapFlags Then LineStates.Add "DEVSPECIFIC", CStr(LineStates.Count)
    If LINEDEVSTATE_REINIT And lineDevCap.dwDevCapFlags Then LineStates.Add "REINIT", CStr(LineStates.Count)
    If LINEDEVSTATE_LOCK And lineDevCap.dwDevCapFlags Then LineStates.Add "LOCK", CStr(LineStates.Count)
    If LINEDEVSTATE_CAPSCHANGE And lineDevCap.dwDevCapFlags Then LineStates.Add "CAPSCHANGE", CStr(LineStates.Count)
    If LINEDEVSTATE_CONFIGCHANGE And lineDevCap.dwDevCapFlags Then LineStates.Add "CONFIGCHANGE", CStr(LineStates.Count)
    If LINEDEVSTATE_TRANSLATECHANGE And lineDevCap.dwDevCapFlags Then LineStates.Add "TRANSLATECHANGE", CStr(LineStates.Count)
    If LINEDEVSTATE_COMPLCANCEL And lineDevCap.dwDevCapFlags Then LineStates.Add "COMPLCANCEL", CStr(LineStates.Count)
    If LINEDEVSTATE_REMOVED And lineDevCap.dwDevCapFlags Then LineStates.Add "REMOVED", CStr(LineStates.Count)
End Sub

Private Sub LoadMediaModes()
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If LINEMEDIAMODE_UNKNOWN And lineDevCap.dwMediaModes Then MediaModes.Add "UNKNOWN", CStr(MediaModes.Count)
    If LINEMEDIAMODE_INTERACTIVEVOICE And lineDevCap.dwMediaModes Then MediaModes.Add "INTERACTIVEVOICE", CStr(MediaModes.Count)
    If LINEMEDIAMODE_AUTOMATEDVOICE And lineDevCap.dwMediaModes Then MediaModes.Add "AUTOMATEDVOICE", CStr(MediaModes.Count)
    If LINEMEDIAMODE_DATAMODEM And lineDevCap.dwMediaModes Then MediaModes.Add "DATAMODEM", CStr(MediaModes.Count)
    If LINEMEDIAMODE_G3FAX And lineDevCap.dwMediaModes Then MediaModes.Add "G3FAX", CStr(MediaModes.Count)
    If LINEMEDIAMODE_TDD And lineDevCap.dwMediaModes Then MediaModes.Add "TDD", CStr(MediaModes.Count)
    If LINEMEDIAMODE_G4FAX And lineDevCap.dwMediaModes Then MediaModes.Add "G4FAX", CStr(MediaModes.Count)
    If LINEMEDIAMODE_DIGITALDATA And lineDevCap.dwMediaModes Then MediaModes.Add "DIGITALDATA", CStr(MediaModes.Count)
    If LINEMEDIAMODE_TELETEX And lineDevCap.dwMediaModes Then MediaModes.Add "TELETEX", CStr(MediaModes.Count)
    If LINEMEDIAMODE_VIDEOTEX And lineDevCap.dwMediaModes Then MediaModes.Add "VIDEOTEX", CStr(MediaModes.Count)
    If LINEMEDIAMODE_TELEX And lineDevCap.dwMediaModes Then MediaModes.Add "TELEX", CStr(MediaModes.Count)
    If LINEMEDIAMODE_MIXED And lineDevCap.dwMediaModes Then MediaModes.Add "MIXED", CStr(MediaModes.Count)
    If LINEMEDIAMODE_ADSI And lineDevCap.dwMediaModes Then MediaModes.Add "ADSI", CStr(MediaModes.Count)
End Sub

Private Sub LoadDeviceModes()
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If LINEDEVCAPFLAGS_CROSSADDRCONF And lineDevCap.dwDevCapFlags Then DeviceModes.Add "CROSSADDRCONF", CStr(DeviceModes.Count)
    If LINEDEVCAPFLAGS_MULTIPLEADDR And lineDevCap.dwDevCapFlags Then DeviceModes.Add "MULTIPLEADDR", CStr(DeviceModes.Count)
    If LINEDEVCAPFLAGS_CLOSEDROP And lineDevCap.dwDevCapFlags Then DeviceModes.Add "CLOSEDROP", CStr(DeviceModes.Count)
    If LINEDEVCAPFLAGS_DIALBILLING And lineDevCap.dwDevCapFlags Then DeviceModes.Add "DIALBILLING", CStr(DeviceModes.Count)
    If LINEDEVCAPFLAGS_DIALQUIET And lineDevCap.dwDevCapFlags Then DeviceModes.Add "DIALQUIET", CStr(DeviceModes.Count)
    If LINEDEVCAPFLAGS_DIALDIALTONE And lineDevCap.dwDevCapFlags Then DeviceModes.Add "DIALDIALTONE", CStr(DeviceModes.Count)
    If LINEDEVCAPFLAGS_HIGHLEVCOMP And lineDevCap.dwDevCapFlags Then DeviceModes.Add "HIGHLEVCOMP", CStr(DeviceModes.Count)
    If LINEDEVCAPFLAGS_LOWLEVCOMP And lineDevCap.dwDevCapFlags Then DeviceModes.Add "LOWLEVCOMP", CStr(DeviceModes.Count)
    If LINEDEVCAPFLAGS_MEDIACONTROL And lineDevCap.dwDevCapFlags Then DeviceModes.Add "MEDIACONTROL", CStr(DeviceModes.Count)
End Sub

Private Sub LoadAddressModes()
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If LINEADDRESSMODE_ADDRESSID And lineDevCap.dwAddressModes Then AddressModes.Add "ADDRESSID", CStr(AddressModes.Count)
    If LINEADDRESSMODE_DIALABLEADDR And lineDevCap.dwAddressModes Then AddressModes.Add "DIALABLEADDR", CStr(AddressModes.Count)
End Sub

Private Sub LoadGenerateToneModes()
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If LINETONEMODE_CUSTOM And lineDevCap.dwGenerateToneModes Then GenerateToneModes.Add "CUSTOM", CStr(GenerateToneModes.Count)
    If LINETONEMODE_RINGBACK And lineDevCap.dwGenerateToneModes Then GenerateToneModes.Add "RINGBACK", CStr(GenerateToneModes.Count)
    If LINETONEMODE_BUSY And lineDevCap.dwGenerateToneModes Then GenerateToneModes.Add "BUSY", CStr(GenerateToneModes.Count)
    If LINETONEMODE_BEEP And lineDevCap.dwGenerateToneModes Then GenerateToneModes.Add "BEEP", CStr(GenerateToneModes.Count)
    If LINETONEMODE_BILLING And lineDevCap.dwGenerateToneModes Then GenerateToneModes.Add "BILLING", CStr(GenerateToneModes.Count)
End Sub

Private Sub LoadBearerModes()
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If LINEBEARERMODE_VOICE And lineDevCap.dwBearerModes Then BearerModes.Add "VOICE", CStr(BearerModes.Count)
    If LINEBEARERMODE_SPEECH And lineDevCap.dwBearerModes Then BearerModes.Add "SPEECH", CStr(BearerModes.Count)
    If LINEBEARERMODE_DATA And lineDevCap.dwBearerModes Then BearerModes.Add "Data", CStr(BearerModes.Count)
    If LINEBEARERMODE_ALTSPEECHDATA And lineDevCap.dwBearerModes Then BearerModes.Add "ALTSPEECHDATA", CStr(BearerModes.Count)
    If LINEBEARERMODE_MULTIUSE And lineDevCap.dwBearerModes Then BearerModes.Add "MULTIUSE", CStr(BearerModes.Count)
    If LINEBEARERMODE_NONCALLSIGNALING And lineDevCap.dwBearerModes Then BearerModes.Add "NONCALLSIGNALING", CStr(BearerModes.Count)
End Sub

Public Property Get Addresses() As Addresses
Attribute Addresses.VB_Description = "This returns a collection of AddressItemObj."
    If mvarAddresses Is Nothing Then Set mvarAddresses = New Addresses
    Set Addresses = mvarAddresses
End Property

Public Property Get Calls() As CallsObj
Attribute Calls.VB_Description = "This returns a collection of CallObj."
    If mvarCallsObj Is Nothing Then Set mvarCallsObj = New CallsObj
    Set Calls = mvarCallsObj
End Property

Public Property Get Config() As DeviceItemConfigObj
    If mvarDeviceItemConfigObj Is Nothing Then Set mvarDeviceItemConfigObj = New DeviceItemConfigObj
    Set Config = mvarDeviceItemConfigObj
End Property

Private Sub Class_Terminate()
    Set mvarAddresses = Nothing
    Set mvarCallsObj = Nothing
    Set mvarDeviceItemConfigObj = Nothing
End Sub

Private Sub Class_Initialize()
    Index = DeviceItem
    
    Set mvarAddresses = New Addresses
    Set mvarCallsObj = New CallsObj
    Set mvarDeviceItemConfigObj = New DeviceItemConfigObj
    
    For AddressItem = 0 To NumberOfAddresses - 1
        mvarAddresses.Add New AddressItemObj, AddressItem
    Next AddressItem
    
    Call LoadMediaModes
    Call LoadLineStates
    Call LoadBearerModes
    Call LoadDeviceModes
    Call LoadAddressModes
    Call LoadGenerateToneModes
    Call LoadGenerateDigitModes
    Call LoadMonitorDigitModes
End Sub

Public Function SetStatusOTHER() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_OTHER, 0)
End Function

Public Function SetStatusRINGING() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_RINGING, 0)
End Function

Public Function SetStatusCONNECTED() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_CONNECTED, 0)
End Function

Public Function SetStatusMSGWAITON() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_MSGWAITON, 0)
End Function

Public Function SetStatusMSGWAITOFF() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_MSGWAITOFF, 0)
End Function

Public Function SetStatusINSERVICE() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_INSERVICE, 0)
End Function

Public Function SetStatusOUTOFSERVICE() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_OUTOFSERVICE, 0)
End Function

Public Function SetStatusMAINTENANCE() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_MAINTENANCE, 0)
End Function

Public Function SetStatusOPEN() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_OPEN, 0)
End Function

Public Function SetStatusCLOSE() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_CLOSE, 0)
End Function

Public Function SetStatusNUMCALLS() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_NUMCALLS, 0)
End Function

Public Function SetStatusNUMCOMPLETIONS() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_NUMCOMPLETIONS, 0)
End Function

Public Function SetStatusTERMINALS() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_TERMINALS, 0)
End Function

Public Function SetStatusROAMMODE() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_ROAMMODE, 0)
End Function

Public Function SetStatusBATTERY() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_BATTERY, 0)
End Function

Public Function SetStatusSIGNAL() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_SIGNAL, 0)
End Function

Public Function SetStatusSPECIFIC() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_DEVSPECIFIC, 0)
End Function

Public Function SetStatusREINIT() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_REINIT, 0)
End Function

Public Function SetStatusLOCK() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_LOCK, 0)
End Function

Public Function SetStatusCAPSCHANGE() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_CAPSCHANGE, 0)
End Function

Public Function SetStatusCONFIGCHANGE() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_CONFIGCHANGE, 0)
End Function

Public Function SetStatusTRANSLATECHANGE() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_TRANSLATECHANGE, 0)
End Function

Public Function SetStatusCOMPLCANCEL() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_COMPLCANCEL, 0)
End Function

Public Function SetStatusREMOVED() As Long
    Call Func_lineSetStatusMessages(LINEDEVSTATE_REMOVED, 0)
End Function

Private Function Func_lineSetStatusMessages(ByRef LineStates As Long, ByRef AddressStates As Long) As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineSetStatusMessages"
        Close #FreeFileNo
    End If
    
    Func_lineSetStatusMessages = lineSetStatusMessages(hndOpenLine, LineStates, AddressStates)
End Function

Public Property Get BatteryLevel() As Long
    Dim lineDevState As LINEDEVSTATUS
    
    Let lineDevState.dwTotalSize = Len(lineDevState)
    
    Call lineGetLineDevStatus(hndOpenLine, lineDevState)
    
    BatteryLevel = lineDevState.dwBatteryLevel
End Property

Public Property Get DevStatusSpecific() As String
    Dim lineDevState As LINEDEVSTATUS
    
    Let lineDevState.dwTotalSize = Len(lineDevState)
    
    Call lineGetLineDevStatus(hndOpenLine, lineDevState)
    
    If lineDevState.dwDevSpecificOffset <> 0 Then
        DevStatusSpecific = GetDevStatusString(lineDevState.mem, lineDevState.dwDevSpecificOffset, lineDevState.dwDevSpecificSize)
    End If
End Property

Public Property Get NumActiveCalls() As Long
    Dim lineDevState As LINEDEVSTATUS
    
    Let lineDevState.dwTotalSize = Len(lineDevState)
    
    Call lineGetLineDevStatus(hndOpenLine, lineDevState)
    
    NumActiveCalls = lineDevState.dwNumActiveCalls
End Property

Public Property Get NumOpens() As Long
    Dim lineDevState As LINEDEVSTATUS
    
    Let lineDevState.dwTotalSize = Len(lineDevState)
    
    Call lineGetLineDevStatus(hndOpenLine, lineDevState)
    
    NumOpens = lineDevState.dwNumOpens
End Property

Public Property Get NumOnHoldCalls() As Long
Attribute NumOnHoldCalls.VB_Description = "The number of calls on hold."
    Dim lineDevState As LINEDEVSTATUS
    
    Let lineDevState.dwTotalSize = Len(lineDevState)
    Call lineGetLineDevStatus(hndOpenLine, lineDevState)
    
    NumOnHoldCalls = lineDevState.dwNumOnHoldCalls
End Property

Public Property Get NumOnHoldPendCalls() As Long
    Dim lineDevState As LINEDEVSTATUS
    
    Let lineDevState.dwTotalSize = Len(lineDevState)
    
    Call lineGetLineDevStatus(hndOpenLine, lineDevState)
    
    NumOnHoldPendCalls = lineDevState.dwNumOnHoldPendCalls
End Property

Public Property Get NumCallCompletions() As Long
Attribute NumCallCompletions.VB_Description = "The number of calls completed since the line was opened."
    Dim lineDevState As LINEDEVSTATUS
    
    Let lineDevState.dwTotalSize = Len(lineDevState)
    
    Call lineGetLineDevStatus(hndOpenLine, lineDevState)
    
    NumCallCompletions = lineDevState.dwNumCallCompletions
End Property

Public Property Get SignalLevel() As Long
Attribute SignalLevel.VB_Description = "The signal level on the line device."
    Dim lineDevState As LINEDEVSTATUS
    
    Let lineDevState.dwTotalSize = Len(lineDevState)
    
    Call lineGetLineDevStatus(hndOpenLine, lineDevState)
    
    SignalLevel = lineDevState.dwSignalLevel
End Property

Private Function GetDevStatusString(mem As String, offset As Long, size As Long) As String
    GetDevStatusString = Trim(Replace(Replace(Mid(mem, offset + 1 - LINEDEVSTATUS_FIXEDSIZE, size - 1), Chr(0), " "), "|", " "))
End Function

Public Function Func_lineGetStatusMessages(ByRef LineStates As Long, ByRef AddressStates As Long)
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineGetStatusMessages"
        Close #FreeFileNo
    End If
    
    Func_lineGetStatusMessages = lineGetStatusMessages(hndOpenLine, LineStates, AddressStates)
End Function

Public Property Get TerminalText() As String
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If lineDevCap.dwTerminalTextSize <> 0 Then
        TerminalText = GetDevCapString(lineDevCap.mem, lineDevCap.dwTerminalTextOffset, lineDevCap.dwTerminalTextSize)
    End If
End Property

Public Property Get ProviderInfo() As String
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If lineDevCap.dwProviderInfoSize <> 0 Then
        ProviderInfo = GetDevCapString(lineDevCap.mem, lineDevCap.dwProviderInfoOffset, lineDevCap.dwProviderInfoSize)
    End If
End Property

Public Property Get SwitchInfo() As String
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If lineDevCap.dwSwitchInfoSize <> 0 Then
        SwitchInfo = GetDevCapString(lineDevCap.mem, lineDevCap.dwSwitchInfoOffset, lineDevCap.dwSwitchInfoSize)
    End If
End Property

Public Property Get DevCapsSpecific() As String
    Dim lineDevCap As LINEDEVCAPS
    
    lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If lineDevCap.dwDevSpecificSize <> 0 Then
        DevCapsSpecific = GetDevCapString(lineDevCap.mem, lineDevCap.dwDevSpecificOffset, lineDevCap.dwDevSpecificSize)
    End If
End Property

Public Property Get TerminalCaps() As String
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If lineDevCap.dwTerminalCapsSize <> 0 Then
        TerminalCaps = GetDevCapString(lineDevCap.mem, lineDevCap.dwTerminalCapsOffset, lineDevCap.dwTerminalCapsSize)
    End If
End Property

Public Property Get DeviceName() As String
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    If lineDevCap.dwLineNameSize <> 0 Then
        DeviceName = GetDevCapString(lineDevCap.mem, lineDevCap.dwLineNameOffset, lineDevCap.dwLineNameSize)
    End If
End Property

Public Property Get NumberOfAddresses() As Long
Attribute NumberOfAddresses.VB_Description = "The number of addresses on the line device"
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    NumberOfAddresses = lineDevCap.dwNumAddresses
End Property

Public Property Get MaxRate() As Long
Attribute MaxRate.VB_Description = "Max data rate of the line device"
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    MaxRate = lineDevCap.dwMaxRate
End Property

Public Property Get PermanentLineID() As Long
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    PermanentLineID = lineDevCap.dwPermanentLineID
End Property

Public Property Get NumTerminals() As Long
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    NumTerminals = lineDevCap.dwNumTerminals
End Property

Public Property Get RingModes() As Long
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    RingModes = lineDevCap.dwRingModes
End Property

Public Property Get GenerateToneMaxFreq() As Long
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    GenerateToneMaxFreq = lineDevCap.dwGenerateToneMaxNumFreq
End Property

Public Property Get MonitorToneMaxFreq() As Long
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    MonitorToneMaxFreq = lineDevCap.dwMonitorToneMaxNumFreq
End Property

Public Property Get MonitorToneMaxEntries() As Long
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    MonitorToneMaxEntries = lineDevCap.dwMonitorToneMaxNumEntries
End Property

Public Property Get StringFormat() As String
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    Select Case lineDevCap.dwStringFormat
        Case STRINGFORMAT_ASCII: StringFormat = "ASCII"
        Case STRINGFORMAT_DBCS: StringFormat = "DBCS"
        Case STRINGFORMAT_UNICODE: StringFormat = "UNICODE"
        Case STRINGFORMAT_BINARY: StringFormat = "BINARY"
    End Select
End Property

Public Property Get AnswerMode() As String
    Dim lineDevCap As LINEDEVCAPS
    
    Let lineDevCap.dwTotalSize = Len(lineDevCap)
    
    Call lineGetDevCaps(hndLine, Index, Func_lineNegotiateAPIVersion, &H0, lineDevCap)
    
    Select Case lineDevCap.dwAnswerMode
        Case LINEANSWERMODE_NONE: AnswerMode = "NONE"
        Case LINEANSWERMODE_DROP: AnswerMode = "DROP"
        Case LINEANSWERMODE_HOLD: AnswerMode = "HOLD"
    End Select
End Property

Private Function GetDevCapString(mem As String, offset As Long, size As Long) As String
    GetDevCapString = Trim(Replace(Replace(Mid(mem, offset + 1 - LINEDEVCAPS_FIXEDSIZE, size - 1), Chr(0), " "), "|", " "))
End Function

Private Function Func_lineNegotiateAPIVersion() As Long
    Dim lAPIVer As Long, linexID As LINEEXTENSIONID
    
    Call lineNegotiateAPIVersion(hndLine, Index, TAPI_LOWAPI, TAPI_HIGHAPI, lAPIVer, linexID)
    
    Func_lineNegotiateAPIVersion = lAPIVer
End Function

Public Function CurrentLocationID() As Long
    Dim lineTranslateCap As LINETRANSLATECAPS
    
    Let lineTranslateCap.dwTotalSize = Len(lineTranslateCap)
    
    Call lineGetTranslateCaps(hndLine, Func_lineNegotiateAPIVersion, lineTranslateCap)
    
    CurrentLocationID = lineTranslateCap.dwCurrentLocationID
End Function

Public Property Get APIVersion() As String
Attribute APIVersion.VB_Description = "The version of TAPI in use on the line device."
    APIVersion = (Trim(Str(Val(Left(Hex(Func_lineNegotiateAPIVersion), 1)))) & "." & Trim(Str(Val(Right(Hex(Func_lineNegotiateAPIVersion), 4)))))
End Property

Public Function Func_lineClose() As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineClose"
        Close #FreeFileNo
    End If
    
    Func_lineClose = lineClose(hndOpenLine)
    
    Let Module1.LastFunction(Index) = "lineClose"
End Function

Public Function Func_lineOpen(ByVal Privilege As Long, ByVal Mode As Long) As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
            
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineOpen"
        Close #FreeFileNo
    End If
    
    Dim lineCallParam As LINECALLPARAMS
    
    Let lineCallParam.dwTotalSize = Len(lineCallParam)
    
    Func_lineOpen = LineOpen( _
        hndLine, _
        Index, _
        hndOpenLine, _
        Func_lineNegotiateAPIVersion, _
        &H0, _
        ObjPtr(Me), _
        Privilege, _
        Mode, _
        lineCallParam _
    )
    
    Let Module1.LastFunction(Index) = "lineOpen"
End Function

Public Function Func_lineMakeCall(ByVal DestAddress As String, ByVal CountryCode As Long) As Long
    If Module1.DebugMode Then
        Dim AppPath As String, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
    
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineMakeCall"
        Close #FreeFileNo
    End If
    
    Dim lineCallParam As LINECALLPARAMS
    
    Let lineCallParam.dwTotalSize = Len(lineCallParam)
    
    Func_lineMakeCall = lineMakeCall( _
        hndOpenLine, _
        DeviceItemHndCall(Index, Me.Calls.Count), _
        Trim(DestAddress), _
        CountryCode, _
        lineCallParam _
    )
    
    Let Module1.LastFunction(Index) = "lineMakeCall"
End Function

Public Property Get Last_Function() As String
    Last_Function = Module1.LastFunction(Index)
End Property

Public Property Get Last_Event() As String
    Last_Event = Module1.LastEvent(Index)
End Property

'Private Function Func_lineMonitorTones(ByRef monitorTones As LINEMONITORTONE, ByVal numTones As Integer) As Long
'    Func_lineMonitorTones = lineMonitorTones(itemParams(activeItem).hndCall, monitorTones, numTones)
'End Function


'Public Function Func_lineGetProviderList(ByVal DeviceID As Long, ByRef providerlist As LPPROVIDERLIST)
'    Dim providerlist As lineproviderlist
'    providerlist.dwTotalSize = Len(providerlist)
'    Func_lineGetProviderList = lineGetProviderList(Func_lineNegotiateAPIVersion(ID), providerlist)
'End Function



