VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddressItemObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is an address on the line device."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private Const TAPI_LOWAPI = &H10003
Private Const TAPI_HIGHAPI = &H20001

Private Const LINEPARKMODE_DIRECTED = &H1
Private Const LINEPARKMODE_NONDIRECTED = &H2

Private Const LINETRANSFERMODE_TRANSFER = &H1
Private Const LINETRANSFERMODE_CONFERENCE = &H2

Private Const LINECALLCOMPLCOND_BUSY = &H1
Private Const LINECALLCOMPLCOND_NOANSWER = &H2

Private Const LINECALLCOMPLMODE_CAMPON = &H1
Private Const LINECALLCOMPLMODE_CALLBACK = &H2
Private Const LINECALLCOMPLMODE_INTRUDE = &H4
Private Const LINECALLCOMPLMODE_MESSAGE = &H8

Private Const LINECALLSELECT_LINE = &H1&
Private Const LINECALLSELECT_ADDRESS = &H2&
Private Const LINECALLSELECT_CALL = &H4&

Private Const LINEREMOVEFROMCONF_NONE = 1&
Private Const LINEREMOVEFROMCONF_LAST = 2&
Private Const LINEREMOVEFROMCONF_ANY = 3&

Private Const LINEBUSYMODE_STATION = &H1
Private Const LINEBUSYMODE_TRUNK = &H2
Private Const LINEBUSYMODE_UNKNOWN = &H4
Private Const LINEBUSYMODE_UNAVAIL = &H8

Private Const LINESPECIALINFO_NOCIRCUIT = &H1
Private Const LINESPECIALINFO_CUSTIRREG = &H2
Private Const LINESPECIALINFO_REORDER = &H4
Private Const LINESPECIALINFO_UNKNOWN = &H8
Private Const LINESPECIALINFO_UNAVAIL = &H10

Private Const LINEDIALTONEMODE_NORMAL = &H1
Private Const LINEDIALTONEMODE_SPECIAL = &H2
Private Const LINEDIALTONEMODE_INTERNAL = &H4
Private Const LINEDIALTONEMODE_EXTERNAL = &H8
Private Const LINEDIALTONEMODE_UNKNOWN = &H10
Private Const LINEDIALTONEMODE_UNAVAIL = &H20

Private Const LINEFORWARDMODE_UNCOND = &H1
Private Const LINEFORWARDMODE_UNCONDINTERNAL = &H2
Private Const LINEFORWARDMODE_UNCONDEXTERNAL = &H4
Private Const LINEFORWARDMODE_UNCONDSPECIFIC = &H8
Private Const LINEFORWARDMODE_BUSY = &H10
Private Const LINEFORWARDMODE_BUSYINTERNAL = &H20
Private Const LINEFORWARDMODE_BUSYEXTERNAL = &H40
Private Const LINEFORWARDMODE_BUSYSPECIFIC = &H80
Private Const LINEFORWARDMODE_NOANSW = &H100
Private Const LINEFORWARDMODE_NOANSWINTERNAL = &H200
Private Const LINEFORWARDMODE_NOANSWEXTERNAL = &H400
Private Const LINEFORWARDMODE_NOANSWSPECIFIC = &H800
Private Const LINEFORWARDMODE_BUSYNA = &H1000
Private Const LINEFORWARDMODE_BUSYNAINTERNAL = &H2000
Private Const LINEFORWARDMODE_BUSYNAEXTERNAL = &H4000
Private Const LINEFORWARDMODE_BUSYNASPECIFIC = &H8000
Private Const LINEFORWARDMODE_UNKNOWN = &H10000         '   // TAPI v1.4
Private Const LINEFORWARDMODE_UNAVAIL = &H20000         ' // TAPI v1.4

Private Const LINEDISCONNECTMODE_NORMAL = &H1
Private Const LINEDISCONNECTMODE_UNKNOWN = &H2
Private Const LINEDISCONNECTMODE_REJECT = &H4
Private Const LINEDISCONNECTMODE_PICKUP = &H8
Private Const LINEDISCONNECTMODE_FORWARDED = &H10
Private Const LINEDISCONNECTMODE_BUSY = &H20
Private Const LINEDISCONNECTMODE_NOANSWER = &H40
Private Const LINEDISCONNECTMODE_BADADDRESS = &H80
Private Const LINEDISCONNECTMODE_UNREACHABLE = &H100
Private Const LINEDISCONNECTMODE_CONGESTION = &H200
Private Const LINEDISCONNECTMODE_INCOMPATIBLE = &H400
Private Const LINEDISCONNECTMODE_UNAVAIL = &H800
Private Const LINEDISCONNECTMODE_NODIALTONE = &H1000    ' TAPI v1.4
Private Const LINEDISCONNECTMODE_NUMBERCHANGED = &H2000 ' TAPI v2.0
Private Const LINEDISCONNECTMODE_OUTOFORDER = &H4000    ' TAPI v2.0
Private Const LINEDISCONNECTMODE_TEMPFAILURE = &H8000   ' TAPI v2.0
Private Const LINEDISCONNECTMODE_QOSUNAVAIL = &H10000   ' TAPI v2.0
Private Const LINEDISCONNECTMODE_BLOCKED = &H20000      ' TAPI v2.0
Private Const LINEDISCONNECTMODE_DONOTDISTURB = &H40000 ' TAPI v2.0
Private Const LINEDISCONNECTMODE_CANCELLED = &H80000    ' TAPI v2.0

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

Private Const LINEADDRCAPFLAGS_FWDNUMRINGS = &H1
Private Const LINEADDRCAPFLAGS_PICKUPGROUPID = &H2
Private Const LINEADDRCAPFLAGS_SECURE = &H4
Private Const LINEADDRCAPFLAGS_BLOCKIDDEFAULT = &H8
Private Const LINEADDRCAPFLAGS_BLOCKIDOVERRIDE = &H10
Private Const LINEADDRCAPFLAGS_DIALED = &H20
Private Const LINEADDRCAPFLAGS_ORIGOFFHOOK = &H40
Private Const LINEADDRCAPFLAGS_DESTOFFHOOK = &H80
Private Const LINEADDRCAPFLAGS_FWDCONSULT = &H100
Private Const LINEADDRCAPFLAGS_SETUPCONFNULL = &H200
Private Const LINEADDRCAPFLAGS_AUTORECONNECT = &H400
Private Const LINEADDRCAPFLAGS_COMPLETIONID = &H800
Private Const LINEADDRCAPFLAGS_TRANSFERHELD = &H1000
Private Const LINEADDRCAPFLAGS_TRANSFERMAKE = &H2000
Private Const LINEADDRCAPFLAGS_CONFERENCEHELD = &H4000
Private Const LINEADDRCAPFLAGS_CONFERENCEMAKE = &H8000
Private Const LINEADDRCAPFLAGS_PARTIALDIAL = &H10000
Private Const LINEADDRCAPFLAGS_FWDSTATUSVALID = &H20000
Private Const LINEADDRCAPFLAGS_FWDINTEXTADDR = &H40000
Private Const LINEADDRCAPFLAGS_FWDBUSYNAADDR = &H80000
Private Const LINEADDRCAPFLAGS_ACCEPTTOALERT = &H100000
Private Const LINEADDRCAPFLAGS_CONFDROP = &H200000
Private Const LINEADDRCAPFLAGS_PICKUPCALLWAIT = &H400000
Private Const LINEADDRCAPFLAGS_PREDICTIVEDIALER = &H800000  ' TAPI v2.0
Private Const LINEADDRCAPFLAGS_QUEUE = &H1000000            ' TAPI v2.0
Private Const LINEADDRCAPFLAGS_ROUTEPOINT = &H2000000       ' TAPI v2.0
Private Const LINEADDRCAPFLAGS_HOLDMAKESNEW = &H4000000     ' TAPI v2.0
Private Const LINEADDRCAPFLAGS_NOINTERNALCALLS = &H8000000  ' TAPI v2.0
Private Const LINEADDRCAPFLAGS_NOEXTERNALCALLS = &H10000000 ' TAPI v2.0
Private Const LINEADDRCAPFLAGS_SETCALLINGID = &H20000000    ' TAPI v2.0

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

Private Const LINECALLPARTYID_BLOCKED = &H1&
Private Const LINECALLPARTYID_OUTOFAREA = &H2&
Private Const LINECALLPARTYID_NAME = &H4&
Private Const LINECALLPARTYID_ADDRESS = &H8&
Private Const LINECALLPARTYID_PARTIAL = &H10&
Private Const LINECALLPARTYID_UNKNOWN = &H20&
Private Const LINECALLPARTYID_UNAVAIL = &H40&

Private Const LINEADDRESSSHARING_PRIVATE = &H1&
Private Const LINEADDRESSSHARING_BRIDGEDEXCL = &H2&
Private Const LINEADDRESSSHARING_BRIDGEDNEW = &H4&
Private Const LINEADDRESSSHARING_BRIDGEDSHARED = &H8&
Private Const LINEADDRESSSHARING_MONITORED = &H10&

Private Const LINEADDRESSMODE_ADDRESSID = &H1&
Private Const LINEADDRESSMODE_DIALABLEADDR = &H2&

Private Const LINEADDRESSSTATE_OTHER = &H1&
Private Const LINEADDRESSSTATE_DEVSPECIFIC = &H2&
Private Const LINEADDRESSSTATE_INUSEZERO = &H4&
Private Const LINEADDRESSSTATE_INUSEONE = &H8&
Private Const LINEADDRESSSTATE_INUSEMANY = &H10&
Private Const LINEADDRESSSTATE_NUMCALLS = &H20&
Private Const LINEADDRESSSTATE_FORWARD = &H40&
Private Const LINEADDRESSSTATE_TERMINALS = &H80&
Private Const LINEADDRESSSTATE_CAPSCHANGE = &H100&

Private Type LINEEXTENSIONID
    dwExtensionID0 As Long
    dwExtensionID1 As Long
    dwExtensionID2 As Long
    dwExtensionID3 As Long
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

Private Type LINEADDRESSSTATUS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwNumInUse As Long
    dwNumActiveCalls As Long
    dwNumOnHoldCalls As Long
    dwNumOnHoldPendCalls As Long
    dwAddressFeatures As Long

    dwNumRingsNoAnswer As Long
    dwForwardNumEntries As Long
    dwForwardSize As Long
    dwForwardOffset As Long

    dwTerminalModesSize As Long
    dwTerminalModesOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long

    mem As String * 2048
End Type

Private Type LINECALLLIST
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwCallsNumEntries As Long
    dwCallsSize As Long
    dwCallsOffset As Long
    
    mem As String * 2048
End Type

Private Const LINECALLLIST_FIXEDSIZE = 24
Private Const LINEADDRESSCAPS_FIXEDSIZE = 176
Private Const LINEADDRESSSTATUS_FIXEDSIZE = 64

Private Declare Function lineNegotiateAPIVersion Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPILowVersion As Long, ByVal dwAPIHighVersion As Long, lpdwAPIVersion As Long, lpExtensionID As LINEEXTENSIONID) As Long
Private Declare Function lineGetAddressCaps Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAddressID As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion As Long, ByRef lpAddressCaps As LINEADDRESSCAPS) As Long
Private Declare Function lineTranslateAddress Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPIVersion As Long, ByVal lpszAddressIn As String, ByVal dwCard As Long, ByVal dwTranslateOptions As Long, lpTranslateOutput As Any) As Long
Private Declare Function lineSetNumRings Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal dwNumRings As Long) As Long
Private Declare Function lineGetNumRings Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, lpdwNumRings As Long) As Long
Private Declare Function lineGetNewCalls Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal dwSelect As Long, lpCallList As Any) As Long
'Private Declare Function lineGetAddressID Lib "TAPI32.DLL" (ByVal hLine As Long, lpdwAddressID As Long, ByVal dwAddressMode As Long, lpsAddress As Any, ByVal dwSize As Long) As Long
Private Declare Function lineGetAddressStatus Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, lpAddressStatus As LINEADDRESSSTATUS) As Long

Public AddressStates As New Collection
Public CallInfoStates As New Collection
Public RedirectionIDFlags As New Collection
Public RedirectingIDFlags As New Collection
Public CallerIDFlags As New Collection
Public CalledIDFlags As New Collection
Public CallStates As New Collection
Public ConnectedIDFlags As New Collection
Public DialToneModes As New Collection
Public BusyModes As New Collection
Public SpecialInfo As New Collection
Public DisconnectModes As New Collection
Public AddrCapFlags As New Collection
Public RemoveFromConfCaps As New Collection
Public TransferModes As New Collection
Public ParkModes As New Collection
Public ForwardModes As New Collection
Public CallCompletionConds As New Collection
Public CallCompletionModes As New Collection

Dim Index As Long
Dim AddressIndex As Long

Public Property Get CallList() As String
    Dim typCallList As LINECALLLIST
    
    Let typCallList.dwTotalSize = Len(typCallList)
    
    Call lineGetNewCalls(Index, AddressIndex, LINECALLSELECT_ADDRESS, typCallList)

    If typCallList.dwCallsSize <> 0 Then
        CallList = GetCallListString(typCallList.mem, typCallList.dwCallsOffset, typCallList.dwCallsSize)
    End If
End Property

Private Function GetCallListString(mem As String, offset As Long, size As Long) As String
    GetCallListString = Trim(Replace(Replace(Mid(mem, offset + 1 - LINECALLLIST_FIXEDSIZE, size - 1), Chr(0), " "), "|", " "))
End Function
    
Private Function GetAddrStatusString(mem As String, offset As Long, size As Long) As String
    GetAddrStatusString = Trim(Replace(Replace(Mid(mem, offset + 1 - LINEADDRESSSTATUS_FIXEDSIZE, size - 1), Chr(0), " "), "|", " "))
End Function

Public Property Get Forward() As String
    Dim AddressStatus As LINEADDRESSSTATUS
    
    Let AddressStatus.dwTotalSize = Len(AddressStatus)
    Call lineGetAddressStatus(hndLine, Index, AddressStatus)
    
    If AddressStatus.dwForwardSize <> 0 Then
        Forward = GetAddrStatusString(AddressStatus.mem, AddressStatus.dwForwardOffset, AddressStatus.dwForwardSize)
    End If
End Property

Public Property Get TerminalModes() As String
    Dim AddressStatus As LINEADDRESSSTATUS
    
    Let AddressStatus.dwTotalSize = Len(AddressStatus)
    Call lineGetAddressStatus(hndLine, Index, AddressStatus)
    
    If AddressStatus.dwForwardSize <> 0 Then
        TerminalModes = GetAddrStatusString(AddressStatus.mem, AddressStatus.dwTerminalModesOffset, AddressStatus.dwTerminalModesSize)
    End If
End Property

Public Property Get StatusDeviceSpecific() As String
    Dim AddressStatus As LINEADDRESSSTATUS
    
    Let AddressStatus.dwTotalSize = Len(AddressStatus)
    Call lineGetAddressStatus(hndLine, Index, AddressStatus)
    
    If AddressStatus.dwForwardSize <> 0 Then
        StatusDeviceSpecific = GetAddrStatusString(AddressStatus.mem, AddressStatus.dwDevSpecificOffset, AddressStatus.dwDevSpecificSize)
    End If
End Property

Public Property Get NumForwardEntries() As Long
    Dim AddressStatus As LINEADDRESSSTATUS
    
    Let AddressStatus.dwTotalSize = Len(AddressStatus)
    Call lineGetAddressStatus(hndLine, Index, AddressStatus)
    
    NumForwardEntries = AddressStatus.dwForwardNumEntries
End Property

Public Property Get NumRingsNoAnswer() As Long
    Dim AddressStatus As LINEADDRESSSTATUS
    
    Let AddressStatus.dwTotalSize = Len(AddressStatus)
    Call lineGetAddressStatus(hndLine, Index, AddressStatus)
    
    NumRingsNoAnswer = AddressStatus.dwNumRingsNoAnswer
End Property

Public Property Get AddressFeatures() As Long
    Dim AddressStatus As LINEADDRESSSTATUS
    
    Let AddressStatus.dwTotalSize = Len(AddressStatus)
    Call lineGetAddressStatus(hndLine, Index, AddressStatus)
    
    AddressFeatures = AddressStatus.dwAddressFeatures
End Property

Public Property Get NumOnHoldPendCalls() As Long
    Dim AddressStatus As LINEADDRESSSTATUS

    Let AddressStatus.dwTotalSize = Len(AddressStatus)
    Call lineGetAddressStatus(hndLine, Index, AddressStatus)

    NumOnHoldPendCalls = AddressStatus.dwNumOnHoldPendCalls
End Property

Public Property Get NumOnHoldCalls() As Long
    Dim AddressStatus As LINEADDRESSSTATUS

    Let AddressStatus.dwTotalSize = Len(AddressStatus)
    Call lineGetAddressStatus(hndLine, Index, AddressStatus)

    NumOnHoldCalls = AddressStatus.dwNumOnHoldCalls
End Property

Public Property Get NumActiveCalls() As Long
    Dim AddressStatus As LINEADDRESSSTATUS

    Let AddressStatus.dwTotalSize = Len(AddressStatus)
    Call lineGetAddressStatus(hndLine, Index, AddressStatus)

    NumActiveCalls = AddressStatus.dwNumActiveCalls
End Property

Public Property Get NumInUse() As Long
    Dim AddressStatus As LINEADDRESSSTATUS

    Let AddressStatus.dwTotalSize = Len(AddressStatus)
    Call lineGetAddressStatus(hndLine, Index, AddressStatus)

    NumInUse = AddressStatus.dwNumInUse
End Property

Public Property Get Rings() As Long
    Dim numRings As Long
    Call lineGetNumRings(hndLine, Index, numRings)
    Rings = numRings
End Property

Public Property Let Rings(numRings As Long)
    Call lineSetNumRings(hndLine, Index, numRings)
End Property

Private Sub Class_Initialize()
    Index = DeviceItem
    AddressIndex = AddressItem

    Call LoadAddressStates
    Call LoadCallInfoStates
    Call LoadRedirectionIDFlags
    Call LoadRedirectingIDFlags
    Call LoadConnectedIDFlags
    Call LoadCallerIDFlags
    Call LoadCalledIDFlags
    Call LoadCallStates
    Call LoadDialToneModes
    Call LoadBusyModes
    Call LoadSpecialInfo
    Call LoadDisconnectModes
    Call LoadAddrCapFlags
    Call LoadRemoveFromConfCaps
    Call LoadTransferModes
    Call LoadParkModes
    Call LoadForwardModes
    Call LoadCallCompletionConds
    Call LoadCallCompletionModes
End Sub

Private Function GetAddrCapString(mem As String, offset As Long, size As Long) As String
    GetAddrCapString = Trim(Replace(Replace(Mid(mem, offset + 1 - LINEADDRESSCAPS_FIXEDSIZE, size - 1), Chr(0), " "), "|", " "))
End Function

Private Function Func_lineNegotiateAPIVersion() As Long
    Dim lAPIVer As Long, linexID As LINEEXTENSIONID

    Call lineNegotiateAPIVersion(hndLine, Index, TAPI_LOWAPI, TAPI_HIGHAPI, lAPIVer, linexID)

    Func_lineNegotiateAPIVersion = lAPIVer
End Function

Public Property Get Address() As String
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If addressCap.dwAddressSize <> 0 Then
        Address = GetAddrCapString(addressCap.mem, addressCap.dwAddressOffset, addressCap.dwAddressSize)
    End If
End Property

Public Property Get CapsDeviceSpecific() As String
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If addressCap.dwAddressSize <> 0 Then
        CapsDeviceSpecific = GetAddrCapString(addressCap.mem, addressCap.dwDevSpecificOffset, addressCap.dwDevSpecificSize)
    End If
End Property

Public Property Get CompletionMsgText() As String
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If addressCap.dwAddressSize <> 0 Then
        CompletionMsgText = GetAddrCapString(addressCap.mem, addressCap.dwCompletionMsgTextOffset, addressCap.dwCompletionMsgTextSize)
    End If
End Property

Public Property Get CallTreatmentList() As String
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If addressCap.dwAddressSize <> 0 Then
        CallTreatmentList = GetAddrCapString(addressCap.mem, addressCap.dwCallTreatmentListOffset, addressCap.dwCallTreatmentListSize)
    End If
End Property

Public Property Get DeviceClasses() As String
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If addressCap.dwAddressSize <> 0 Then
        DeviceClasses = GetAddrCapString(addressCap.mem, addressCap.dwDeviceClassesOffset, addressCap.dwDeviceClassesSize)
    End If
End Property

Private Sub LoadAddrCapFlags()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINEADDRCAPFLAGS_FWDNUMRINGS And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_PICKUPGROUPID And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_SECURE And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_BLOCKIDDEFAULT And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_BLOCKIDOVERRIDE And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_DIALED And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_ORIGOFFHOOK And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_DESTOFFHOOK And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_FWDCONSULT And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_SETUPCONFNULL And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_AUTORECONNECT And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_COMPLETIONID And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_TRANSFERHELD And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_CONFERENCEHELD And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_PARTIALDIAL And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_FWDSTATUSVALID And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_FWDINTEXTADDR And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_ACCEPTTOALERT And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_CONFDROP And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
    If LINEADDRCAPFLAGS_PICKUPCALLWAIT And addressCap.dwAddrCapFlags Then AddrCapFlags.Add "OTHER", CStr(AddrCapFlags.Count)
End Sub

Public Property Get AddressSharing() As String
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    Select Case addressCap.dwAddressSharing
        Case LINEADDRESSSHARING_PRIVATE: AddressSharing = "PRIVATE"
        Case LINEADDRESSSHARING_BRIDGEDEXCL: AddressSharing = "BRIDGEDEXCL"
        Case LINEADDRESSSHARING_BRIDGEDNEW: AddressSharing = "BRIDGEDNEW"
        Case LINEADDRESSSHARING_BRIDGEDSHARED: AddressSharing = "BRIDGEDSHARED"
        Case LINEADDRESSSHARING_MONITORED: AddressSharing = "MONITORED"
    End Select
End Property

Private Sub LoadAddressStates()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINEADDRESSSTATE_OTHER And addressCap.dwAddressStates Then AddressStates.Add "OTHER", CStr(AddressStates.Count)
    If LINEADDRESSSTATE_DEVSPECIFIC And addressCap.dwAddressStates Then AddressStates.Add "DEVSPECIFIC", CStr(AddressStates.Count)
    If LINEADDRESSSTATE_INUSEZERO And addressCap.dwAddressStates Then AddressStates.Add "INUSEZERO", CStr(AddressStates.Count)
    If LINEADDRESSSTATE_INUSEONE And addressCap.dwAddressStates Then AddressStates.Add "INUSEONE", CStr(AddressStates.Count)
    If LINEADDRESSSTATE_INUSEMANY And addressCap.dwAddressStates Then AddressStates.Add "INUSEMANY", CStr(AddressStates.Count)
    If LINEADDRESSSTATE_NUMCALLS And addressCap.dwAddressStates Then AddressStates.Add "NUMCALLS", CStr(AddressStates.Count)
    If LINEADDRESSSTATE_FORWARD And addressCap.dwAddressStates Then AddressStates.Add "FORWARD", CStr(AddressStates.Count)
    If LINEADDRESSSTATE_TERMINALS And addressCap.dwAddressStates Then AddressStates.Add "TERMINALS", CStr(AddressStates.Count)
    If LINEADDRESSSTATE_CAPSCHANGE And addressCap.dwAddressStates Then AddressStates.Add "CAPSCHANGE", CStr(AddressStates.Count)
End Sub

Public Property Get AvailableMediaModes() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    AvailableMediaModes = addressCap.dwAvailableMediaModes
End Property

Private Sub LoadBusyModes()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINEBUSYMODE_STATION And addressCap.dwBusyModes Then BusyModes.Add "STATION", CStr(BusyModes.Count)
    If LINEBUSYMODE_TRUNK And addressCap.dwBusyModes Then BusyModes.Add "TRUNK", CStr(BusyModes.Count)
    If LINEBUSYMODE_UNKNOWN And addressCap.dwBusyModes Then BusyModes.Add "UNKNOWN", CStr(BusyModes.Count)
    If LINEBUSYMODE_UNAVAIL And addressCap.dwBusyModes Then BusyModes.Add "UNAVAIL", CStr(BusyModes.Count)
End Sub

Private Sub LoadCallCompletionConds()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINECALLCOMPLCOND_BUSY And addressCap.dwCallCompletionConds Then CallCompletionConds.Add "BUSY", CStr(CallCompletionConds.Count)
    If LINECALLCOMPLCOND_NOANSWER And addressCap.dwCallCompletionConds Then CallCompletionConds.Add "NOANSWER", CStr(CallCompletionConds.Count)
End Sub

Private Sub LoadCallCompletionModes()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINECALLCOMPLMODE_CAMPON And addressCap.dwCallCompletionModes Then CallCompletionModes.Add "CAMPON", CStr(CallCompletionModes.Count)
    If LINECALLCOMPLMODE_CALLBACK And addressCap.dwCallCompletionModes Then CallCompletionModes.Add "CALLBACK", CStr(CallCompletionModes.Count)
    If LINECALLCOMPLMODE_INTRUDE And addressCap.dwCallCompletionModes Then CallCompletionModes.Add "INTRUDE", CStr(CallCompletionModes.Count)
    If LINECALLCOMPLMODE_MESSAGE And addressCap.dwCallCompletionModes Then CallCompletionModes.Add "MESSAGE", CStr(CallCompletionModes.Count)
End Sub

Private Sub LoadCalledIDFlags()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINECALLPARTYID_BLOCKED And addressCap.dwCalledIDFlags Then CalledIDFlags.Add "BLOCKED", CStr(CalledIDFlags.Count)
    If LINECALLPARTYID_OUTOFAREA And addressCap.dwCalledIDFlags Then CalledIDFlags.Add "OUTOFAREA", CStr(CalledIDFlags.Count)
    If LINECALLPARTYID_NAME And addressCap.dwCalledIDFlags Then CalledIDFlags.Add "NAME", CStr(CalledIDFlags.Count)
    If LINECALLPARTYID_ADDRESS And addressCap.dwCalledIDFlags Then CalledIDFlags.Add "ADDRESS", CStr(CalledIDFlags.Count)
    If LINECALLPARTYID_PARTIAL And addressCap.dwCalledIDFlags Then CalledIDFlags.Add "PARTIAL", CStr(CalledIDFlags.Count)
    If LINECALLPARTYID_UNKNOWN And addressCap.dwCalledIDFlags Then CalledIDFlags.Add "UNKNOWN", CStr(CalledIDFlags.Count)
    If LINECALLPARTYID_UNAVAIL And addressCap.dwCalledIDFlags Then CalledIDFlags.Add "UNAVAIL", CStr(CalledIDFlags.Count)
    If LINECALLPARTYID_BLOCKED And addressCap.dwCalledIDFlags Then CalledIDFlags.Add "BLOCKED", CStr(CalledIDFlags.Count)
End Sub

Private Sub LoadCallerIDFlags()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINECALLPARTYID_BLOCKED And addressCap.dwCallerIDFlags Then CallerIDFlags.Add "BLOCKED", CStr(CallerIDFlags.Count)
    If LINECALLPARTYID_OUTOFAREA And addressCap.dwCallerIDFlags Then CallerIDFlags.Add "OUTOFAREA", CStr(CallerIDFlags.Count)
    If LINECALLPARTYID_NAME And addressCap.dwCallerIDFlags Then CallerIDFlags.Add "NAME", CStr(CallerIDFlags.Count)
    If LINECALLPARTYID_ADDRESS And addressCap.dwCallerIDFlags Then CallerIDFlags.Add "ADDRESS", CStr(CallerIDFlags.Count)
    If LINECALLPARTYID_PARTIAL And addressCap.dwCallerIDFlags Then CallerIDFlags.Add "PARTIAL", CStr(CallerIDFlags.Count)
    If LINECALLPARTYID_UNKNOWN And addressCap.dwCallerIDFlags Then CallerIDFlags.Add "UNKNOWN", CStr(CallerIDFlags.Count)
    If LINECALLPARTYID_UNAVAIL And addressCap.dwCallerIDFlags Then CallerIDFlags.Add "UNAVAIL", CStr(CallerIDFlags.Count)
    If LINECALLPARTYID_BLOCKED And addressCap.dwCallerIDFlags Then CallerIDFlags.Add "BLOCKED", CStr(CallerIDFlags.Count)
End Sub

Private Sub LoadConnectedIDFlags()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINECALLPARTYID_BLOCKED And addressCap.dwConnectedIDFlags Then ConnectedIDFlags.Add "BLOCKED", CStr(ConnectedIDFlags.Count)
    If LINECALLPARTYID_OUTOFAREA And addressCap.dwConnectedIDFlags Then ConnectedIDFlags.Add "OUTOFAREA", CStr(ConnectedIDFlags.Count)
    If LINECALLPARTYID_NAME And addressCap.dwConnectedIDFlags Then ConnectedIDFlags.Add "NAME", CStr(ConnectedIDFlags.Count)
    If LINECALLPARTYID_ADDRESS And addressCap.dwConnectedIDFlags Then ConnectedIDFlags.Add "ADDRESS", CStr(ConnectedIDFlags.Count)
    If LINECALLPARTYID_PARTIAL And addressCap.dwConnectedIDFlags Then ConnectedIDFlags.Add "PARTIAL", CStr(ConnectedIDFlags.Count)
    If LINECALLPARTYID_UNKNOWN And addressCap.dwConnectedIDFlags Then ConnectedIDFlags.Add "UNKNOWN", CStr(ConnectedIDFlags.Count)
    If LINECALLPARTYID_UNAVAIL And addressCap.dwConnectedIDFlags Then ConnectedIDFlags.Add "UNAVAIL", CStr(ConnectedIDFlags.Count)
    If LINECALLPARTYID_BLOCKED And addressCap.dwConnectedIDFlags Then ConnectedIDFlags.Add "BLOCKED", CStr(ConnectedIDFlags.Count)
End Sub

Public Property Get CallFeatures1() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    CallFeatures1 = addressCap.dwCallFeatures
End Property

Public Property Get CallFeatures2() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    CallFeatures2 = addressCap.dwCallFeatures2
End Property

Private Sub LoadCallInfoStates()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINEDEVSTATE_OTHER And addressCap.dwCallInfoStates Then CallInfoStates.Add "OTHER", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_RINGING And addressCap.dwCallInfoStates Then CallInfoStates.Add "RINGING", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_CONNECTED And addressCap.dwCallInfoStates Then CallInfoStates.Add "CONNECTED", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_DISCONNECTED And addressCap.dwCallInfoStates Then CallInfoStates.Add "DISCONNECTED", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_MSGWAITON And addressCap.dwCallInfoStates Then CallInfoStates.Add "MSGWAITON", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_MSGWAITOFF And addressCap.dwCallInfoStates Then CallInfoStates.Add "MSGWAITOFF", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_NUMCOMPLETIONS And addressCap.dwCallInfoStates Then CallInfoStates.Add "NUMCOMPLETIONS", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_INSERVICE And addressCap.dwCallInfoStates Then CallInfoStates.Add "INSERVICE", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_OUTOFSERVICE And addressCap.dwCallInfoStates Then CallInfoStates.Add "OUTOFSERVICE", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_MAINTENANCE And addressCap.dwCallInfoStates Then CallInfoStates.Add "MAINTENANCE", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_OPEN And addressCap.dwCallInfoStates Then CallInfoStates.Add "OPEN", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_CLOSE And addressCap.dwCallInfoStates Then CallInfoStates.Add "CLOSE", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_NUMCALLS And addressCap.dwCallInfoStates Then CallInfoStates.Add "NUMCALLS", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_TERMINALS And addressCap.dwCallInfoStates Then CallInfoStates.Add "TERMINALS", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_ROAMMODE And addressCap.dwCallInfoStates Then CallInfoStates.Add "ROAMMODE", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_BATTERY And addressCap.dwCallInfoStates Then CallInfoStates.Add "BATTERY", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_SIGNAL And addressCap.dwCallInfoStates Then CallInfoStates.Add "SIGNAL", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_DEVSPECIFIC And addressCap.dwCallInfoStates Then CallInfoStates.Add "DEVSPECIFIC", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_REINIT And addressCap.dwCallInfoStates Then CallInfoStates.Add "REINIT", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_LOCK And addressCap.dwCallInfoStates Then CallInfoStates.Add "LOCK", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_CAPSCHANGE And addressCap.dwCallInfoStates Then CallInfoStates.Add "CAPSCHANGE", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_CONFIGCHANGE And addressCap.dwCallInfoStates Then CallInfoStates.Add "CONFIGCHANGE", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_TRANSLATECHANGE And addressCap.dwCallInfoStates Then CallInfoStates.Add "TRANSLATECHANGE", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_COMPLCANCEL And addressCap.dwCallInfoStates Then CallInfoStates.Add "COMPLCANCEL", CStr(CallInfoStates.Count)
    If LINEDEVSTATE_REMOVED And addressCap.dwCallInfoStates Then CallInfoStates.Add "REMOVED", CStr(CallInfoStates.Count)
End Sub

Private Sub LoadCallStates()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINECALLSTATE_IDLE And addressCap.dwCallStates Then CallStates.Add "IDLE", CStr(CallStates.Count)
    If LINECALLSTATE_OFFERING And addressCap.dwCallStates Then CallStates.Add "OFFERING", CStr(CallStates.Count)
    If LINECALLSTATE_ACCEPTED And addressCap.dwCallStates Then CallStates.Add "ACCEPTED", CStr(CallStates.Count)
    If LINECALLSTATE_DIALTONE And addressCap.dwCallStates Then CallStates.Add "DIALTONE", CStr(CallStates.Count)
    If LINECALLSTATE_DIALING And addressCap.dwCallStates Then CallStates.Add "DIALING", CStr(CallStates.Count)
    If LINECALLSTATE_RINGBACK And addressCap.dwCallStates Then CallStates.Add "RINGBACK", CStr(CallStates.Count)
    If LINECALLSTATE_BUSY And addressCap.dwCallStates Then CallStates.Add "BUSY", CStr(CallStates.Count)
    If LINECALLSTATE_SPECIALINFO And addressCap.dwCallStates Then CallStates.Add "SPECIALINFO", CStr(CallStates.Count)
    If LINECALLSTATE_CONNECTED And addressCap.dwCallStates Then CallStates.Add "CONNECTED", CStr(CallStates.Count)
    If LINECALLSTATE_PROCEEDING And addressCap.dwCallStates Then CallStates.Add "PROCEEDING", CStr(CallStates.Count)
    If LINECALLSTATE_ONHOLD And addressCap.dwCallStates Then CallStates.Add "ONHOLD", CStr(CallStates.Count)
    If LINECALLSTATE_CONFERENCED And addressCap.dwCallStates Then CallStates.Add "CONFERENCED", CStr(CallStates.Count)
    If LINECALLSTATE_ONHOLDPENDCONF And addressCap.dwCallStates Then CallStates.Add "ONHOLDPENDCONF", CStr(CallStates.Count)
    If LINECALLSTATE_ONHOLDPENDTRANSFER And addressCap.dwCallStates Then CallStates.Add "ONHOLDPENDTRANSF", CStr(CallStates.Count)
    If LINECALLSTATE_DISCONNECTED And addressCap.dwCallStates Then CallStates.Add "DISCONNECTED", CStr(CallStates.Count)
    If LINECALLSTATE_UNKNOWN And addressCap.dwCallStates Then CallStates.Add "UNKNOWN", CStr(CallStates.Count)
End Sub

Public Property Get CompletionMsgTextEntrySize() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    CompletionMsgTextEntrySize = addressCap.dwCompletionMsgTextEntrySize
End Property

Public Property Get ConnectedModes() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    ConnectedModes = addressCap.dwConnectedModes
End Property

Private Sub LoadDialToneModes()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINEDIALTONEMODE_NORMAL And addressCap.dwDialToneModes Then DialToneModes.Add "NORMAL", CStr(DialToneModes.Count)
    If LINEDIALTONEMODE_SPECIAL And addressCap.dwDialToneModes Then DialToneModes.Add "SPECIAL", CStr(DialToneModes.Count)
    If LINEDIALTONEMODE_INTERNAL And addressCap.dwDialToneModes Then DialToneModes.Add "INTERNAL", CStr(DialToneModes.Count)
    If LINEDIALTONEMODE_EXTERNAL And addressCap.dwDialToneModes Then DialToneModes.Add "EXTERNAL", CStr(DialToneModes.Count)
    If LINEDIALTONEMODE_UNKNOWN And addressCap.dwDialToneModes Then DialToneModes.Add "UNKNOWN", CStr(DialToneModes.Count)
    If LINEDIALTONEMODE_UNAVAIL And addressCap.dwDialToneModes Then DialToneModes.Add "UNAVAIL", CStr(DialToneModes.Count)
End Sub

Private Sub LoadDisconnectModes()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINEDISCONNECTMODE_NORMAL And addressCap.dwDisconnectModes Then DisconnectModes.Add "NORMAL", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_UNKNOWN And addressCap.dwDisconnectModes Then DisconnectModes.Add "UNKNOWN", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_REJECT And addressCap.dwDisconnectModes Then DisconnectModes.Add "REJECT", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_PICKUP And addressCap.dwDisconnectModes Then DisconnectModes.Add "PICKUP", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_FORWARDED And addressCap.dwDisconnectModes Then DisconnectModes.Add "FORWARDED", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_BUSY And addressCap.dwDisconnectModes Then DisconnectModes.Add "BUSY", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_NOANSWER And addressCap.dwDisconnectModes Then DisconnectModes.Add "NOANSWER", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_NODIALTONE And addressCap.dwDisconnectModes Then DisconnectModes.Add "NODIALTONE", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_BADADDRESS And addressCap.dwDisconnectModes Then DisconnectModes.Add "BADADDRESS", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_UNREACHABLE And addressCap.dwDisconnectModes Then DisconnectModes.Add "UNREACHABLE", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_CONGESTION And addressCap.dwDisconnectModes Then DisconnectModes.Add "CONGESTION", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_INCOMPATIBLE And addressCap.dwDisconnectModes Then DisconnectModes.Add "INCOMPATIBLE", CStr(DisconnectModes.Count)
    If LINEDISCONNECTMODE_UNAVAIL And addressCap.dwDisconnectModes Then DisconnectModes.Add "UNAVAIL", CStr(DisconnectModes.Count)
End Sub

Private Sub LoadForwardModes()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINEFORWARDMODE_UNCOND And addressCap.dwForwardModes Then ForwardModes.Add "UNCOND", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_UNCONDINTERNAL And addressCap.dwForwardModes Then ForwardModes.Add "UNCONDINTERNAL", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_UNCONDEXTERNAL And addressCap.dwForwardModes Then ForwardModes.Add "UNCONDEXTERNAL", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_UNCONDSPECIFIC And addressCap.dwForwardModes Then ForwardModes.Add "UNCONDSPECIFIC", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_BUSY And addressCap.dwForwardModes Then ForwardModes.Add "BUSY", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_BUSYINTERNAL And addressCap.dwForwardModes Then ForwardModes.Add "BUSYINTERNAL", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_BUSYEXTERNAL And addressCap.dwForwardModes Then ForwardModes.Add "BUSYEXTERNAL", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_BUSYSPECIFIC And addressCap.dwForwardModes Then ForwardModes.Add "BUSYSPECIFIC", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_NOANSW And addressCap.dwForwardModes Then ForwardModes.Add "NOANSW", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_NOANSWINTERNAL And addressCap.dwForwardModes Then ForwardModes.Add "NOANSWINTERNAL", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_NOANSWEXTERNAL And addressCap.dwForwardModes Then ForwardModes.Add "NOANSWEXTERNAL", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_NOANSWSPECIFIC And addressCap.dwForwardModes Then ForwardModes.Add "NOANSWSPECIFIC", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_BUSYNA And addressCap.dwForwardModes Then ForwardModes.Add "BUSYNA", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_BUSYNAINTERNAL And addressCap.dwForwardModes Then ForwardModes.Add "BUSYNAINTERNAL", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_BUSYNAEXTERNAL And addressCap.dwForwardModes Then ForwardModes.Add "BUSYNAEXTERNAL", CStr(ForwardModes.Count)
    If LINEFORWARDMODE_BUSYNASPECIFIC And addressCap.dwForwardModes Then ForwardModes.Add "BUSYNASPECIFIC", CStr(ForwardModes.Count)
End Sub

Public Property Get LineDeviceID() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    LineDeviceID = addressCap.dwLineDeviceID
End Property

Public Property Get MaxCallCompletions() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxCallCompletions = addressCap.dwMaxCallCompletions
End Property

Public Property Get MaxCallDataSize() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxCallDataSize = addressCap.dwMaxCallDataSize
End Property

Public Property Get MaxForwardEntries() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxForwardEntries = addressCap.dwMaxForwardEntries
End Property

Public Property Get MaxFwdNumRings() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxFwdNumRings = addressCap.dwMaxFwdNumRings
End Property

Public Property Get MaxNoAnswerTimeout() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxNoAnswerTimeout = addressCap.dwMaxNoAnswerTimeout
End Property

Public Property Get MaxNumActiveCalls() As Long
Attribute MaxNumActiveCalls.VB_Description = "The maximum amount of calls for this address on the line device."
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxNumActiveCalls = addressCap.dwMaxNumActiveCalls
End Property

Public Property Get MaxNumConference() As Long
Attribute MaxNumConference.VB_Description = "The maximum amount of conference's for this address on the line device."
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxNumConference = addressCap.dwMaxNumConference
End Property

Public Property Get MaxNumOnHoldCalls() As Long
Attribute MaxNumOnHoldCalls.VB_Description = "The maximum amount of calls on hold for this address on the line device."
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxNumOnHoldCalls = addressCap.dwMaxNumOnHoldCalls
End Property

Public Property Get MaxNumOnHoldPendingCalls() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxNumOnHoldPendingCalls = addressCap.dwMaxNumOnHoldPendingCalls
End Property

Public Property Get MaxNumTransConf() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxNumTransConf = addressCap.dwMaxNumTransConf
End Property

Public Property Get MaxSpecificEntries() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MaxSpecificEntries = addressCap.dwMaxSpecificEntries
End Property

Public Property Get MinFwdNumRings() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    MinFwdNumRings = addressCap.dwMinFwdNumRings
End Property

Public Property Get NumCallTreatments() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    NumCallTreatments = addressCap.dwNumCallTreatments
End Property

Public Property Get NumCompletionMessages() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    NumCompletionMessages = addressCap.dwNumCompletionMessages
End Property

Public Property Get OfferingModes() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    OfferingModes = addressCap.dwOfferingModes
End Property

Private Sub LoadParkModes()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINEPARKMODE_DIRECTED And addressCap.dwParkModes Then ParkModes.Add "DIRECTED", CStr(ParkModes.Count)
    If LINEPARKMODE_NONDIRECTED And addressCap.dwParkModes Then ParkModes.Add "NONDIRECTED", CStr(ParkModes.Count)
End Sub

Public Property Get PredictiveAutoTransferStates() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    PredictiveAutoTransferStates = addressCap.dwPredictiveAutoTransferStates
End Property

Private Sub LoadRedirectingIDFlags()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINECALLPARTYID_BLOCKED And addressCap.dwRedirectingIDFlags Then RedirectingIDFlags.Add "BLOCKED", CStr(RedirectingIDFlags.Count)
    If LINECALLPARTYID_OUTOFAREA And addressCap.dwRedirectingIDFlags Then RedirectingIDFlags.Add "OUTOFAREA", CStr(RedirectingIDFlags.Count)
    If LINECALLPARTYID_NAME And addressCap.dwRedirectingIDFlags Then RedirectingIDFlags.Add "NAME", CStr(RedirectingIDFlags.Count)
    If LINECALLPARTYID_ADDRESS And addressCap.dwRedirectingIDFlags Then RedirectingIDFlags.Add "ADDRESS", CStr(RedirectingIDFlags.Count)
    If LINECALLPARTYID_PARTIAL And addressCap.dwRedirectingIDFlags Then RedirectingIDFlags.Add "PARTIAL", CStr(RedirectingIDFlags.Count)
    If LINECALLPARTYID_UNKNOWN And addressCap.dwRedirectingIDFlags Then RedirectingIDFlags.Add "UNKNOWN", CStr(RedirectingIDFlags.Count)
    If LINECALLPARTYID_UNAVAIL And addressCap.dwRedirectingIDFlags Then RedirectingIDFlags.Add "UNAVAIL", CStr(RedirectingIDFlags.Count)
    If LINECALLPARTYID_BLOCKED And addressCap.dwRedirectingIDFlags Then RedirectingIDFlags.Add "BLOCKED", CStr(RedirectingIDFlags.Count)
End Sub

Private Sub LoadRedirectionIDFlags()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)
 
    If LINECALLPARTYID_BLOCKED And addressCap.dwRedirectionIDFlags Then RedirectionIDFlags.Add "BLOCKED", CStr(RedirectionIDFlags.Count)
    If LINECALLPARTYID_OUTOFAREA And addressCap.dwRedirectionIDFlags Then RedirectionIDFlags.Add "OUTOFAREA", CStr(RedirectionIDFlags.Count)
    If LINECALLPARTYID_NAME And addressCap.dwRedirectionIDFlags Then RedirectionIDFlags.Add "NAME", CStr(RedirectionIDFlags.Count)
    If LINECALLPARTYID_ADDRESS And addressCap.dwRedirectionIDFlags Then RedirectionIDFlags.Add "ADDRESS", CStr(RedirectionIDFlags.Count)
    If LINECALLPARTYID_PARTIAL And addressCap.dwRedirectionIDFlags Then RedirectionIDFlags.Add "PARTIAL", CStr(RedirectionIDFlags.Count)
    If LINECALLPARTYID_UNKNOWN And addressCap.dwRedirectionIDFlags Then RedirectionIDFlags.Add "UNKNOWN", CStr(RedirectionIDFlags.Count)
    If LINECALLPARTYID_UNAVAIL And addressCap.dwRedirectionIDFlags Then RedirectionIDFlags.Add "UNAVAIL", CStr(RedirectionIDFlags.Count)
    If LINECALLPARTYID_BLOCKED And addressCap.dwRedirectionIDFlags Then RedirectionIDFlags.Add "BLOCKED", CStr(RedirectionIDFlags.Count)
End Sub

Private Sub LoadRemoveFromConfCaps()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINEREMOVEFROMCONF_NONE And addressCap.dwRemoveFromConfCaps Then RemoveFromConfCaps.Add "NONE", CStr(RemoveFromConfCaps.Count)
    If LINEREMOVEFROMCONF_LAST And addressCap.dwRemoveFromConfCaps Then RemoveFromConfCaps.Add "LAST", CStr(RemoveFromConfCaps.Count)
    If LINEREMOVEFROMCONF_ANY And addressCap.dwRemoveFromConfCaps Then RemoveFromConfCaps.Add "ANY", CStr(RemoveFromConfCaps.Count)
End Sub

Public Property Get RemoveFromConfState() As Long
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    RemoveFromConfState = addressCap.dwRemoveFromConfState
End Property

Private Sub LoadSpecialInfo()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINESPECIALINFO_NOCIRCUIT And addressCap.dwSpecialInfo Then SpecialInfo.Add "NOCIRCUIT", CStr(SpecialInfo.Count)
    If LINESPECIALINFO_CUSTIRREG And addressCap.dwSpecialInfo Then SpecialInfo.Add "CUSTIRREG", CStr(SpecialInfo.Count)
    If LINESPECIALINFO_REORDER And addressCap.dwSpecialInfo Then SpecialInfo.Add "REORDER", CStr(SpecialInfo.Count)
    If LINESPECIALINFO_UNKNOWN And addressCap.dwSpecialInfo Then SpecialInfo.Add "UNKNOWN", CStr(SpecialInfo.Count)
    If LINESPECIALINFO_UNAVAIL And addressCap.dwSpecialInfo Then SpecialInfo.Add "UNAVAIL", CStr(SpecialInfo.Count)
End Sub

Private Sub LoadTransferModes()
    Dim addressCap  As LINEADDRESSCAPS

    addressCap.dwTotalSize = Len(addressCap)
    Call lineGetAddressCaps(hndLine, Index, AddressIndex, Func_lineNegotiateAPIVersion, &H0, addressCap)

    If LINETRANSFERMODE_TRANSFER And addressCap.dwTransferModes Then TransferModes.Add "TRANSFER", CStr(TransferModes.Count)
    If LINETRANSFERMODE_CONFERENCE And addressCap.dwTransferModes Then TransferModes.Add "CONFERENCE", CStr(TransferModes.Count)
End Sub
