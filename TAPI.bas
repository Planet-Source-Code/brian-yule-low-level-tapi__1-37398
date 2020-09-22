Attribute VB_Name = "Module1"
Option Explicit

Private Const LINEDEVCAPS_FIXEDSIZE = 236
Private Const LINECALLINFO_FIXEDSIZE = 296

Private Const STRINGFORMAT_ASCII = 1&
Private Const STRINGFORMAT_DBCS = 2&
Private Const STRINGFORMAT_UNICODE = 3&
Private Const STRINGFORMAT_BINARY = 4&

Private Const LINECONNECTEDMODE_ACTIVE = &H1&
Private Const LINECONNECTEDMODE_INACTIVE = &H2&
Private Const LINECONNECTEDMODE_ACTIVEHELD = &H4&
Private Const LINECONNECTEDMODE_INACTIVEHELD = &H8&
Private Const LINECONNECTEDMODE_CONFIRMED = &H10&

Private Const LINEROAMMODE_UNKNOWN = &H1&
Private Const LINEROAMMODE_UNAVAIL = &H2&
Private Const LINEROAMMODE_HOME = &H4&
Private Const LINEROAMMODE_ROAMA = &H8&
Private Const LINEROAMMODE_ROAMB = &H10&

Private Const LINEFEATURE_DEVSPECIFIC = &H1&
Private Const LINEFEATURE_DEVSPECIFICFEAT = &H2&
Private Const LINEFEATURE_FORWARD = &H4&
Private Const LINEFEATURE_MAKECALL = &H8&
Private Const LINEFEATURE_SETMEDIACONTROL = &H10&
Private Const LINEFEATURE_SETTERMINAL = &H20&

Private Const LINEDEVSTATUSFLAGS_CONNECTED = &H1&
Private Const LINEDEVSTATUSFLAGS_MSGWAIT = &H2&
Private Const LINEDEVSTATUSFLAGS_INSERVICE = &H4&
Private Const LINEDEVSTATUSFLAGS_LOCKED = &H8&

Private Const LINEDEVCAPFLAGS_CROSSADDRCONF = &H1&
Private Const LINEDEVCAPFLAGS_HIGHLEVCOMP = &H2&
Private Const LINEDEVCAPFLAGS_LOWLEVCOMP = &H4&
Private Const LINEDEVCAPFLAGS_MEDIACONTROL = &H8&
Private Const LINEDEVCAPFLAGS_MULTIPLEADDR = &H10&
Private Const LINEDEVCAPFLAGS_CLOSEDROP = &H20&
Private Const LINEDEVCAPFLAGS_DIALBILLING = &H40&
Private Const LINEDEVCAPFLAGS_DIALQUIET = &H80&
Private Const LINEDEVCAPFLAGS_DIALDIALTONE = &H100&

Private Const LINEADDRESSSHARING_PRIVATE = &H1&
Private Const LINEADDRESSSHARING_BRIDGEDEXCL = &H2&
Private Const LINEADDRESSSHARING_BRIDGEDNEW = &H4&
Private Const LINEADDRESSSHARING_BRIDGEDSHARED = &H8&
Private Const LINEADDRESSSHARING_MONITORED = &H10&

Private Const LINEANSWERMODE_NONE = &H1&
Private Const LINEANSWERMODE_DROP = &H2&
Private Const LINEANSWERMODE_HOLD = &H4&

Private Const LINETONEMODE_CUSTOM = &H1&
Private Const LINETONEMODE_RINGBACK = &H2&
Private Const LINETONEMODE_BUSY = &H4&
Private Const LINETONEMODE_BEEP = &H8&
Private Const LINETONEMODE_BILLING = &H10&
    
Private Const TAPIMAXDESTADDRESSSIZE = 80&
Private Const TAPIMAXAPPNAMESIZE = 40&
Private Const TAPIMAXCALLEDPARTYSIZE = 40&
Private Const TAPIMAXCOMMENTSIZE = 80&
Private Const TAPIMAXDEVICECLASSSIZE = 40&
Private Const TAPIMAXDEVICEIDSIZE = 40&

Private Const LINEADDRESSMODE_ADDRESSID = &H1&
Private Const LINEADDRESSMODE_DIALABLEADDR = &H2&

Private Const LINEREQUESTMODE_MAKECALL = &H1&
Private Const LINEREQUESTMODE_MEDIACALL = &H2&
Private Const LINEREQUESTMODE_DROP = &H4&

Private Const LINECALLPRIVILEGE_NONE = &H1&
Private Const LINECALLPRIVILEGE_MONITOR = &H2&
Private Const LINECALLPRIVILEGE_OWNER = &H4&

Private Const LINEDIGITMODE_PULSE = &H1&
Private Const LINEDIGITMODE_DTMF = &H2&
Private Const LINEDIGITMODE_DTMFEND = &H4&

Private Const LINEADDRESSSTATE_OTHER = &H1&
Private Const LINEADDRESSSTATE_DEVSPECIFIC = &H2&
Private Const LINEADDRESSSTATE_INUSEZERO = &H4&
Private Const LINEADDRESSSTATE_INUSEONE = &H8&
Private Const LINEADDRESSSTATE_INUSEMANY = &H10&
Private Const LINEADDRESSSTATE_NUMCALLS = &H20&
Private Const LINEADDRESSSTATE_FORWARD = &H40&
Private Const LINEADDRESSSTATE_TERMINALS = &H80&
Private Const LINEADDRESSSTATE_CAPSCHANGE = &H100&

Private Const LINECALLORIGIN_OUTBOUND = &H1&
Private Const LINECALLORIGIN_INTERNAL = &H2&
Private Const LINECALLORIGIN_EXTERNAL = &H4&
Private Const LINECALLORIGIN_UNKNOWN = &H10&
Private Const LINECALLORIGIN_UNAVAIL = &H20&
Private Const LINECALLORIGIN_CONFERENCE = &H40&


Private Const LINEDISCONNECTMODE_NORMAL = &H1&
Private Const LINEDISCONNECTMODE_UNKNOWN = &H2&
Private Const LINEDISCONNECTMODE_REJECT = &H4&
Private Const LINEDISCONNECTMODE_PICKUP = &H8&
Private Const LINEDISCONNECTMODE_FORWARDED = &H10&
Private Const LINEDISCONNECTMODE_BUSY = &H20&
Private Const LINEDISCONNECTMODE_NOANSWER = &H40&
Private Const LINEDISCONNECTMODE_BADADDRESS = &H80&
Private Const LINEDISCONNECTMODE_UNREACHABLE = &H100&
Private Const LINEDISCONNECTMODE_CONGESTION = &H200&
Private Const LINEDISCONNECTMODE_INCOMPATIBLE = &H400&
Private Const LINEDISCONNECTMODE_UNAVAIL = &H800&

Private Const LINETERMMODE_BUTTONS = &H1&
Private Const LINETERMMODE_LAMPS = &H2&
Private Const LINETERMMODE_DISPLAY = &H4&
Private Const LINETERMMODE_RINGER = &H8&
Private Const LINETERMMODE_HOOKSWITCH = &H10&
Private Const LINETERMMODE_MEDIATOLINE = &H20&
Private Const LINETERMMODE_MEDIAFROMLINE = &H40&
Private Const LINETERMMODE_MEDIABIDIRECT = &H80&

Private Const LINECALLREASON_DIRECT = &H1&
Private Const LINECALLREASON_FWDBUSY = &H2&
Private Const LINECALLREASON_FWDNOANSWER = &H4&
Private Const LINECALLREASON_FWDUNCOND = &H8&
Private Const LINECALLREASON_PICKUP = &H10&
Private Const LINECALLREASON_UNPARK = &H20&
Private Const LINECALLREASON_REDIRECT = &H40&
Private Const LINECALLREASON_CALLCOMPLETION = &H80&
Private Const LINECALLREASON_TRANSFER = &H100&
Private Const LINECALLREASON_REMINDER = &H200&
Private Const LINECALLREASON_UNKNOWN = &H400&
Private Const LINECALLREASON_UNAVAIL = &H800&

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

Private Const LINECALLPARAMFLAGS_SECURE = &H1&
Private Const LINECALLPARAMFLAGS_IDLE = &H2&
Private Const LINECALLPARAMFLAGS_BLOCKID = &H4&
Private Const LINECALLPARAMFLAGS_ORIGOFFHOOK = &H8&
Private Const LINECALLPARAMFLAGS_DESTOFFHOOK = &H10&

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

Private Const LINECALLINFOSTATE_OTHER = &H1&
Private Const LINECALLINFOSTATE_DEVSPECIFIC = &H2&
Private Const LINECALLINFOSTATE_BEARERMODE = &H4&
Private Const LINECALLINFOSTATE_RATE = &H8&
Private Const LINECALLINFOSTATE_MEDIAMODE = &H10&
Private Const LINECALLINFOSTATE_APPSPECIFIC = &H20&
Private Const LINECALLINFOSTATE_CALLID = &H40&
Private Const LINECALLINFOSTATE_RELATEDCALLID = &H80&
Private Const LINECALLINFOSTATE_ORIGIN = &H100&
Private Const LINECALLINFOSTATE_REASON = &H200&
Private Const LINECALLINFOSTATE_COMPLETIONID = &H400&
Private Const LINECALLINFOSTATE_NUMOWNERINCR = &H800&
Private Const LINECALLINFOSTATE_NUMOWNERDECR = &H1000&
Private Const LINECALLINFOSTATE_NUMMONITORS = &H2000&
Private Const LINECALLINFOSTATE_TRUNK = &H4000&
Private Const LINECALLINFOSTATE_CALLERID = &H8000&
Private Const LINECALLINFOSTATE_CALLEDID = &H10000
Private Const LINECALLINFOSTATE_CONNECTEDID = &H20000
Private Const LINECALLINFOSTATE_REDIRECTIONID = &H40000
Private Const LINECALLINFOSTATE_REDIRECTINGID = &H80000
Private Const LINECALLINFOSTATE_DISPLAY = &H100000
Private Const LINECALLINFOSTATE_USERUSERINFO = &H200000
Private Const LINECALLINFOSTATE_HIGHLEVELCOMP = &H400000
Private Const LINECALLINFOSTATE_LOWLEVELCOMP = &H800000
Private Const LINECALLINFOSTATE_CHARGINGINFO = &H1000000
Private Const LINECALLINFOSTATE_TERMINAL = &H2000000
Private Const LINECALLINFOSTATE_DIALPARAMS = &H4000000
Private Const LINECALLINFOSTATE_MONITORMODES = &H8000000

Private Const LINEERR_ALLOCATED = &H80000001
Private Const LINEERR_BADDEVICEID = &H80000002
Private Const LINEERR_BEARERMODEUNAVAIL = &H80000003
Private Const LINEERR_CALLUNAVAIL = &H80000005
Private Const LINEERR_COMPLETIONOVERRUN = &H80000006
Private Const LINEERR_CONFERENCEFULL = &H80000007
Private Const LINEERR_DIALBILLING = &H80000008
Private Const LINEERR_DIALDIALTONE = &H80000009
Private Const LINEERR_DIALPROMPT = &H8000000A
Private Const LINEERR_DIALQUIET = &H8000000B
Private Const LINEERR_INCOMPATIBLEAPIVERSION = &H8000000C
Private Const LINEERR_INCOMPATIBLEEXTVERSION = &H8000000D
Private Const LINEERR_INIFILECORRUPT = &H8000000E
Private Const LINEERR_INUSE = &H8000000F
Private Const LINEERR_INVALADDRESS = &H80000010
Private Const LINEERR_INVALADDRESSID = &H80000011
Private Const LINEERR_INVALADDRESSMODE = &H80000012
Private Const LINEERR_INVALADDRESSSTATE = &H80000013
Private Const LINEERR_INVALAPPHANDLE = &H80000014
Private Const LINEERR_INVALAPPNAME = &H80000015
Private Const LINEERR_INVALBEARERMODE = &H80000016
Private Const LINEERR_INVALCALLCOMPLMODE = &H80000017
Private Const LINEERR_INVALCALLHANDLE = &H80000018
Private Const LINEERR_INVALCALLPARAMS = &H80000019
Private Const LINEERR_INVALCALLPRIVILEGE = &H8000001A
Private Const LINEERR_INVALCALLSELECT = &H8000001B
Private Const LINEERR_INVALCALLSTATE = &H8000001C
Private Const LINEERR_INVALCALLSTATELIST = &H8000001D
Private Const LINEERR_INVALCARD = &H8000001E
Private Const LINEERR_INVALCOMPLETIONID = &H8000001F
Private Const LINEERR_INVALCONFCALLHANDLE = &H80000020
Private Const LINEERR_INVALCONSULTCALLHANDLE = &H80000021
Private Const LINEERR_INVALCOUNTRYCODE = &H80000022
Private Const LINEERR_INVALDEVICECLASS = &H80000023
Private Const LINEERR_INVALDEVICEHANDLE = &H80000024
Private Const LINEERR_INVALDIALPARAMS = &H80000025
Private Const LINEERR_INVALDIGITLIST = &H80000026
Private Const LINEERR_INVALDIGITMODE = &H80000027
Private Const LINEERR_INVALDIGITS = &H80000028
Private Const LINEERR_INVALEXTVERSION = &H80000029
Private Const LINEERR_INVALGROUPID = &H8000002A
Private Const LINEERR_INVALLINEHANDLE = &H8000002B
Private Const LINEERR_INVALLINESTATE = &H8000002C
Private Const LINEERR_INVALLOCATION = &H8000002D
Private Const LINEERR_INVALMEDIALIST = &H8000002E
Private Const LINEERR_INVALMEDIAMODE = &H8000002F
Private Const LINEERR_INVALMESSAGEID = &H80000030
Private Const LINEERR_INVALPARAM = &H80000032
Private Const LINEERR_INVALPARKID = &H80000033
Private Const LINEERR_INVALPARKMODE = &H80000034
Private Const LINEERR_INVALPOINTER = &H80000035
Private Const LINEERR_INVALPRIVSELECT = &H80000036
Private Const LINEERR_INVALRATE = &H80000037
Private Const LINEERR_INVALREQUESTMODE = &H80000038
Private Const LINEERR_INVALTERMINALID = &H80000039
Private Const LINEERR_INVALTERMINALMODE = &H8000003A
Private Const LINEERR_INVALTIMEOUT = &H8000003B
Private Const LINEERR_INVALTONE = &H8000003C
Private Const LINEERR_INVALTONELIST = &H8000003D
Private Const LINEERR_INVALTONEMODE = &H8000003E
Private Const LINEERR_INVALTRANSFERMODE = &H8000003F
Private Const LINEERR_LINEMAPPERFAILED = &H80000040
Private Const LINEERR_NOCONFERENCE = &H80000041
Private Const LINEERR_NODEVICE = &H80000042
Private Const LINEERR_NODRIVER = &H80000043
Private Const LINEERR_NOMEM = &H80000044
Private Const LINEERR_NOREQUEST = &H80000045
Private Const LINEERR_NOTOWNER = &H80000046
Private Const LINEERR_NOTREGISTERED = &H80000047
Private Const LINEERR_OPERATIONFAILED = &H80000048
Private Const LINEERR_OPERATIONUNAVAIL = &H80000049
Private Const LINEERR_RATEUNAVAIL = &H8000004A
Private Const LINEERR_RESOURCEUNAVAIL = &H8000004B
Private Const LINEERR_REQUESTOVERRUN = &H8000004C
Private Const LINEERR_STRUCTURETOOSMALL = &H8000004D
Private Const LINEERR_TARGETNOTFOUND = &H8000004E
Private Const LINEERR_TARGETSELF = &H8000004F
Private Const LINEERR_UNINITIALIZED = &H80000050
Private Const LINEERR_USERUSERINFOTOOBIG = &H80000051
Private Const LINEERR_REINIT = &H80000052
Private Const LINEERR_ADDRESSBLOCKED = &H80000053
Private Const LINEERR_BILLINGREJECTED = &H80000054
Private Const LINEERR_INVALFEATURE = &H80000055
Private Const LINEERR_NOMULTIPLEINSTANCE = &H80000056
'
' tapi 2.0 only
Private Const LINEERR_INVALAGENTID = &H80000057                        ' TAPI v2.0
Private Const LINEERR_INVALAGENTGROUP = &H80000058                     ' TAPI v2.0
Private Const LINEERR_INVALPASSWORD = &H80000059                       ' TAPI v2.0
Private Const LINEERR_INVALAGENTSTATE = &H8000005A                     ' TAPI v2.0
Private Const LINEERR_INVALAGENTACTIVITY = &H8000005B                  ' TAPI v2.0
Private Const LINEERR_DIALVOICEDETECT = &H8000005C                     ' TAPI v2.0

Private Const LINEBEARERMODE_VOICE = &H1&
Private Const LINEBEARERMODE_SPEECH = &H2&
Private Const LINEBEARERMODE_MULTIUSE = &H4&
Private Const LINEBEARERMODE_DATA = &H8&
Private Const LINEBEARERMODE_ALTSPEECHDATA = &H10&
Private Const LINEBEARERMODE_NONCALLSIGNALING = &H20&


Private Type LINEDIALPARAMS
    dwDialPause As Long
    dwDialSpeed As Long
    dwDigitDuration As Long
    dwWaitForDialtone As Long
End Type

Private Type LINEREQMAKECALL
    szDestAddress As String * TAPIMAXDESTADDRESSSIZE
    szAppName As String * TAPIMAXAPPNAMESIZE
    szCalledParty As String * TAPIMAXCALLEDPARTYSIZE
    szComment As String * TAPIMAXCOMMENTSIZE
End Type

Private Type LINEMONITORTONE
    dwAppSpecific As Long
    dwDuration As Long
    dwFrequency1 As Long
    dwFrequency2 As Long
    dwFrequency3 As Long
End Type

Private Type LINECALLLIST
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwCallsNumEntries As Long
    dwCallsSize As Long
    dwCallsOffset As Long
End Type

Private Type LPPROVIDERLIST
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    
    dwNumProviders As Long
    dwProviderListSize As Long
    dwProviderListOffset As Long

    mem As String * 2048 ' added by mca
End Type

Private Type VARSTRING
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwStringFormat As Long
    dwStringSize As Long
    dwStringOffset As Long

    mem As String * 2048 ' added by mca
End Type

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
    
    mem As String * 2048 ' added by mca
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

    mem As String * 2048 ' added by mca
End Type

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

    mem As String * 2048 ' added by mca
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'   Global Variables in use
'
Public hndLine As Long ' This is used as the handle to the initialized tapi line
Public LastFunction() As String ' This holds the last line function called per device
Public LastEvent() As String ' This holds the last line event that occured per device
Public CallItem As Integer ' This is the current call item when setting up the call object
Public DeviceItem As Integer ' This is the current device item when setting up the device object
Public AddressItem As Integer ' This is the current address item when setting up the address object
Public AppObjPtr As Long ' This is the object pointer to the Application class
Public DeviceItemHndCall() As Long ' This is an array that holds call numbers for all calls per device.
Public RequestMode As Long ' This holds the last request mode that this application received
Public DebugMode As Boolean ' This holds whether debugging to file is on or off
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                                    (dest As Any, src As Any, ByVal Length As Long)
'Private Declare Function DeviceIoControl Lib "kernel32" _
'                                    (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
'                                    (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" _
'                                    (ByVal lpFileName As String) As Long
'
Public Sub LINECALLBACK( _
    ByVal hDevice As Long, _
    ByVal dwMsg As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long, _
    ByVal dwParam3 As Long _
)
    On Error Resume Next
    
    Dim objTemp  As APPLICATION
    
    Call CopyMemory(objTemp, AppObjPtr, 4)
    Call objTemp.RecieveCallBack(hDevice, dwMsg, dwCallbackInstance, dwParam1, dwParam2, dwParam3)
    Call CopyMemory(objTemp, 0&, 4)
End Sub

Public Function TAPI_GETERRORMSG(ERROR_MSG As Long) As String
    On Error Resume Next
    Dim strTemp As String
    Select Case ERROR_MSG
        Case 0: strTemp = "Success"
        Case LINEERR_ALLOCATED: strTemp = "LINEERR_ALLOCATED"
        Case LINEERR_BADDEVICEID: strTemp = "LINEERR_BADDEVICEID"
        Case LINEERR_BEARERMODEUNAVAIL: strTemp = "LINEERR_BEARERMODEUNAVAIL"
        Case LINEERR_CALLUNAVAIL: strTemp = "LINEERR_CALLUNAVAIL"
        Case LINEERR_COMPLETIONOVERRUN: strTemp = "LINEERR_COMPLETIONOVERRUN"
        Case LINEERR_CONFERENCEFULL: strTemp = "LINEERR_CONFERENCEFULL"
        Case LINEERR_DIALBILLING: strTemp = "LINEERR_DIALBILLING"
        Case LINEERR_DIALDIALTONE: strTemp = "LINEERR_DIALDIALTONE"
        Case LINEERR_DIALPROMPT: strTemp = "LINEERR_DIALPROMPT"
        Case LINEERR_DIALQUIET: strTemp = "LINEERR_DIALQUIET"
        Case LINEERR_INCOMPATIBLEAPIVERSION: strTemp = "LINEERR_INCOMPATIBLEAPIVERSION"
        Case LINEERR_INCOMPATIBLEEXTVERSION: strTemp = "LINEERR_INCOMPATIBLEEXTVERSION"
        Case LINEERR_INIFILECORRUPT: strTemp = "LINEERR_INIFILECORRUPT"
        Case LINEERR_INUSE: strTemp = "LINEERR_INUSE"
        Case LINEERR_INVALADDRESS: strTemp = "LINEERR_INVALADDRESS"
        Case LINEERR_INVALADDRESSID: strTemp = "LINEERR_INVALADDRESSID"
        Case LINEERR_INVALADDRESSMODE: strTemp = "LINEERR_INVALADDRESSMODE"
        Case LINEERR_INVALADDRESSSTATE: strTemp = "LINEERR_INVALADDRESSSTATE"
        Case LINEERR_INVALAPPHANDLE: strTemp = "LINEERR_INVALAPPHANDLE"
        Case LINEERR_INVALAPPNAME: strTemp = "LINEERR_INVALAPPNAME"
        Case LINEERR_INVALBEARERMODE: strTemp = "LINEERR_INVALBEARERMODE"
        Case LINEERR_INVALCALLCOMPLMODE:   strTemp = "LINEERR_INVALCALLCOMPLMODE"
        Case LINEERR_INVALCALLHANDLE: strTemp = "LINEERR_INVALCALLHANDLE"
        Case LINEERR_INVALCALLPARAMS: strTemp = "LINEERR_INVALCALLPARAMS"
        Case LINEERR_INVALCALLPRIVILEGE: strTemp = "LINEERR_INVALCALLPRIVILEGE"
        Case LINEERR_INVALCALLSELECT: strTemp = "LINEERR_INVALCALLSELECT"
        Case LINEERR_INVALCALLSTATE: strTemp = "LINEERR_INVALCALLSTATE"
        Case LINEERR_INVALCALLSTATELIST: strTemp = "LINEERR_INVALCALLSTATELIST"
        Case LINEERR_INVALCARD: strTemp = "LINEERR_INVALCARD"
        Case LINEERR_INVALCOMPLETIONID: strTemp = "LINEERR_INVALCOMPLETIONID"
        Case LINEERR_INVALCONFCALLHANDLE: strTemp = "LINEERR_INVALCONFCALLHANDLE"
        Case LINEERR_INVALCONSULTCALLHANDLE: strTemp = "LINEERR_INVALCONSULTCALLHANDLE"
        Case LINEERR_INVALCOUNTRYCODE: strTemp = "LINEERR_INVALCOUNTRYCODE"
        Case LINEERR_INVALDEVICECLASS: strTemp = "LINEERR_INVALDEVICECLASS"
        Case LINEERR_INVALDEVICEHANDLE: strTemp = "LINEERR_INVALDEVICEHANDLE"
        Case LINEERR_INVALDIALPARAMS: strTemp = "LINEERR_INVALDIALPARAMS"
        Case LINEERR_INVALDIGITLIST: strTemp = "LINEERR_INVALDIGITLIST"
        Case LINEERR_INVALDIGITMODE: strTemp = "LINEERR_INVALDIGITMODE"
        Case LINEERR_INVALDIGITS: strTemp = "LINEERR_INVALDIGITS"
        Case LINEERR_INVALEXTVERSION: strTemp = "LINEERR_INVALEXTVERSION"
        Case LINEERR_INVALGROUPID: strTemp = "LINEERR_INVALGROUPID"
        Case LINEERR_INVALLINEHANDLE: strTemp = "LINEERR_INVALLINEHANDLE"
        Case LINEERR_INVALLINESTATE: strTemp = "LINEERR_INVALLINESTATE"
        Case LINEERR_INVALLOCATION: strTemp = "LINEERR_INVALLOCATION"
        Case LINEERR_INVALMEDIALIST: strTemp = "LINEERR_INVALMEDIALIST"
        Case LINEERR_INVALMEDIAMODE: strTemp = "LINEERR_INVALMEDIAMODE"
        Case LINEERR_INVALMESSAGEID: strTemp = "LINEERR_INVALMESSAGEID"
        Case LINEERR_INVALPARAM: strTemp = "LINEERR_INVALPARAM"
        Case LINEERR_INVALPARKID: strTemp = "LINEERR_INVALPARKID"
        Case LINEERR_INVALPARKMODE: strTemp = "LINEERR_INVALPARKMODE"
        Case LINEERR_INVALPOINTER: strTemp = "LINEERR_INVALPOINTER"
        Case LINEERR_INVALPRIVSELECT: strTemp = "LINEERR_INVALPRIVSELECT"
        Case LINEERR_INVALRATE: strTemp = "LINEERR_INVALRATE"
        Case LINEERR_INVALREQUESTMODE: strTemp = "LINEERR_INVALREQUESTMODE"
        Case LINEERR_INVALTERMINALID: strTemp = "LINEERR_INVALTERMINALID"
        Case LINEERR_INVALTERMINALMODE: strTemp = "LINEERR_INVALTERMINALMODE"
        Case LINEERR_INVALTIMEOUT: strTemp = "LINEERR_INVALTIMEOUT"
        Case LINEERR_INVALTONE: strTemp = "LINEERR_INVALTONE"
        Case LINEERR_INVALTONELIST: strTemp = "LINEERR_INVALTONELIST"
        Case LINEERR_INVALTONEMODE: strTemp = "LINEERR_INVALTONEMODE"
        Case LINEERR_INVALTRANSFERMODE: strTemp = "LINEERR_INVALTRANSFERMODE"
        Case LINEERR_LINEMAPPERFAILED: strTemp = "LINEERR_LINEMAPPERFAILED"
        Case LINEERR_NOCONFERENCE: strTemp = "LINEERR_NOCONFERENCE"
        Case LINEERR_NODEVICE: strTemp = "LINEERR_NODEVICE"
        Case LINEERR_NODRIVER: strTemp = "LINEERR_NODRIVER"
        Case LINEERR_NOMEM: strTemp = "LINEERR_NOMEM"
        Case LINEERR_NOREQUEST: strTemp = "LINEERR_NOREQUEST"
        Case LINEERR_NOTOWNER:      strTemp = "LINEERR_NOTOWNER"
        Case LINEERR_NOTREGISTERED: strTemp = "LINEERR_NOTREGISTERED"
        Case LINEERR_OPERATIONFAILED: strTemp = "LINEERR_OPERATIONFAILED"
        Case LINEERR_OPERATIONUNAVAIL: strTemp = "LINEERR_OPERATIONUNAVAIL"
        Case LINEERR_RATEUNAVAIL: strTemp = "LINEERR_RATEUNAVAIL"
        Case LINEERR_RESOURCEUNAVAIL: strTemp = "LINEERR_RESOURCEUNAVAIL"
        Case LINEERR_REQUESTOVERRUN: strTemp = "LINEERR_REQUESTOVERRUN"
        Case LINEERR_STRUCTURETOOSMALL: strTemp = "LINEERR_STRUCTURETOOSMALL"
        Case LINEERR_TARGETNOTFOUND: strTemp = "LINEERR_TARGETNOTFOUND"
        Case LINEERR_TARGETSELF: strTemp = "LINEERR_TARGETSELF"
        Case LINEERR_UNINITIALIZED: strTemp = "LINEERR_UNINITIALIZED"
        Case LINEERR_USERUSERINFOTOOBIG: strTemp = "LINEERR_USERUSERINFOTOOBIG"
        Case LINEERR_REINIT:      strTemp = "LINEERR_REINIT"
        Case LINEERR_ADDRESSBLOCKED: strTemp = "LINEERR_ADDRESSBLOCKED"
        Case LINEERR_BILLINGREJECTED: strTemp = "LINEERR_BILLINGREJECTED"
        Case LINEERR_INVALFEATURE: strTemp = "LINEERR_INVALFEATURE"
        Case LINEERR_NOMULTIPLEINSTANCE: strTemp = "LINEERR_NOMULTIPLEINSTANCE"
        Case LINEERR_INVALAGENTID: strTemp = "LINEERR_INVALAGENTID"
        Case LINEERR_INVALAGENTGROUP: strTemp = "LINEERR_INVALAGENTGROUP"
        Case LINEERR_INVALPASSWORD: strTemp = "LINEERR_INVALPASSWORD"
        Case LINEERR_INVALAGENTSTATE: strTemp = "LINEERR_INVALAGENTSTATE"
        Case LINEERR_INVALAGENTACTIVITY: strTemp = "LINEERR_INVALAGENTACTIVITY"
        Case LINEERR_DIALVOICEDETECT: strTemp = "LINEERR_DIALVOICEDETECT"
        Case Else: strTemp = "Unknown TAPI Error!"
    End Select
    TAPI_GETERRORMSG = strTemp
End Function

Public Function TAPI_GETDEVSTATE(DEV_STATE) As String
    Select Case DEV_STATE
        Case LINEDEVSTATE_RINGING: TAPI_GETDEVSTATE = "RINGING"
        Case LINEDEVSTATE_CONNECTED: TAPI_GETDEVSTATE = "CONNECTED"
        Case LINEDEVSTATE_DISCONNECTED: TAPI_GETDEVSTATE = "DISCONNECTED"
        Case LINEDEVSTATE_MSGWAITON: TAPI_GETDEVSTATE = "MSGWAITON"
        Case LINEDEVSTATE_MSGWAITOFF: TAPI_GETDEVSTATE = "MSGWAITOFF"
        Case LINEDEVSTATE_INSERVICE: TAPI_GETDEVSTATE = "INSERVICE"
        Case LINEDEVSTATE_OUTOFSERVICE: TAPI_GETDEVSTATE = "OUTOFSERVICE"
        Case LINEDEVSTATE_MAINTENANCE: TAPI_GETDEVSTATE = "MAINTENANCE"
        Case LINEDEVSTATE_OPEN: TAPI_GETDEVSTATE = "OPEN"
        Case LINEDEVSTATE_CLOSE: TAPI_GETDEVSTATE = "CLOSE"
        Case LINEDEVSTATE_NUMCALLS: TAPI_GETDEVSTATE = "NUMCALLS"
        Case LINEDEVSTATE_NUMCOMPLETIONS: TAPI_GETDEVSTATE = "NUMCOMPLETIONS"
        Case LINEDEVSTATE_TERMINALS: TAPI_GETDEVSTATE = "TERMINALS"
        Case LINEDEVSTATE_ROAMMODE: TAPI_GETDEVSTATE = "ROAMMODE"
        Case LINEDEVSTATE_BATTERY: TAPI_GETDEVSTATE = "BATTERY"
        Case LINEDEVSTATE_SIGNAL: TAPI_GETDEVSTATE = "SIGNAL"
        Case LINEDEVSTATE_DEVSPECIFIC: TAPI_GETDEVSTATE = "DEVSPECIFIC"
        Case LINEDEVSTATE_REINIT: TAPI_GETDEVSTATE = "REINIT"
        Case LINEDEVSTATE_LOCK: TAPI_GETDEVSTATE = "LOCK"
        Case LINEDEVSTATE_CAPSCHANGE: TAPI_GETDEVSTATE = "CAPSCHANGE"
        Case LINEDEVSTATE_CONFIGCHANGE: TAPI_GETDEVSTATE = "CONFIGCHANGE"
        Case LINEDEVSTATE_TRANSLATECHANGE: TAPI_GETDEVSTATE = "TRANSLATECHANGE"
        Case LINEDEVSTATE_COMPLCANCEL: TAPI_GETDEVSTATE = "COMPLCANCEL"
        Case LINEDEVSTATE_REMOVED: TAPI_GETDEVSTATE = "REMOVED"
        Case LINEDEVSTATE_OTHER: TAPI_GETDEVSTATE = "OTHER"
        Case Else: TAPI_GETDEVSTATE = "OTHER"
    End Select
End Function

Public Function TAPI_GETCALLREASON(CALL_REASON As Long) As String
    Select Case CALL_REASON
        Case LINECALLREASON_DIRECT: TAPI_GETCALLREASON = "DIRECT"
        Case LINECALLREASON_FWDBUSY: TAPI_GETCALLREASON = "FWDBUSY"
        Case LINECALLREASON_FWDNOANSWER: TAPI_GETCALLREASON = "FWDNOANSWER"
        Case LINECALLREASON_FWDUNCOND: TAPI_GETCALLREASON = "FWDUNCOND"
        Case LINECALLREASON_PICKUP: TAPI_GETCALLREASON = "PICKUP"
        Case LINECALLREASON_UNPARK: TAPI_GETCALLREASON = "UNPARK"
        Case LINECALLREASON_REDIRECT: TAPI_GETCALLREASON = "REDIRECT"
        Case LINECALLREASON_CALLCOMPLETION: TAPI_GETCALLREASON = "CALLCOMPLETION"
        Case LINECALLREASON_TRANSFER: TAPI_GETCALLREASON = "TRANSFER"
        Case LINECALLREASON_REMINDER: TAPI_GETCALLREASON = "REMINDER"
        Case LINECALLREASON_UNAVAIL: TAPI_GETCALLREASON = "UNAVAIL"
        Case Else: TAPI_GETCALLREASON = "UNKNOWN"
    End Select
End Function

Public Function TAPI_GETCALLORIGIN(CALL_ORIGIN As Long) As String
    Select Case CALL_ORIGIN
        Case LINECALLORIGIN_OUTBOUND: TAPI_GETCALLORIGIN = "OUTBOUND"
        Case LINECALLORIGIN_INTERNAL:  TAPI_GETCALLORIGIN = "INTERNAL"
        Case LINECALLORIGIN_EXTERNAL:  TAPI_GETCALLORIGIN = "EXTERNAL"
        Case LINECALLORIGIN_UNAVAIL: TAPI_GETCALLORIGIN = "UNAVAIL"
        Case LINECALLORIGIN_CONFERENCE: TAPI_GETCALLORIGIN = "CONFERENCE"
        Case Else: TAPI_GETCALLORIGIN = "UNKNOWN"
    End Select
End Function

Public Function TAPI_GETMEDIAMODE(MEDIA_MODE As Long) As String
    Select Case MEDIA_MODE
        Case LINEMEDIAMODE_INTERACTIVEVOICE: TAPI_GETMEDIAMODE = "INTERACTIVE VOICE"
        Case LINEMEDIAMODE_AUTOMATEDVOICE: TAPI_GETMEDIAMODE = "AUTOMATED VOICE"
        Case LINEMEDIAMODE_DATAMODEM: TAPI_GETMEDIAMODE = "DATA MODEM"
        Case LINEMEDIAMODE_G3FAX: TAPI_GETMEDIAMODE = "G3 FAX"
        Case LINEMEDIAMODE_TDD: TAPI_GETMEDIAMODE = "TDD"
        Case LINEMEDIAMODE_G4FAX: TAPI_GETMEDIAMODE = "G4 FAX:"
        Case LINEMEDIAMODE_DIGITALDATA: TAPI_GETMEDIAMODE = "DIGITAL DATA"
        Case LINEMEDIAMODE_TELETEX: TAPI_GETMEDIAMODE = "TELETEX"
        Case LINEMEDIAMODE_VIDEOTEX: TAPI_GETMEDIAMODE = "VIDEO TEX"
        Case LINEMEDIAMODE_TELEX: TAPI_GETMEDIAMODE = "TELEX"
        Case LINEMEDIAMODE_MIXED: TAPI_GETMEDIAMODE = "MIXED"
        Case LINEMEDIAMODE_ADSI: TAPI_GETMEDIAMODE = "ADSI"
        Case Else: TAPI_GETMEDIAMODE = "UNKNOWN"
    End Select
End Function

Public Function TAPI_GETDIGITMODE(DIGIT_MODE As Long) As String
    Select Case DIGIT_MODE
        Case LINEDIGITMODE_PULSE: TAPI_GETDIGITMODE = "PULSE"
        Case LINEDIGITMODE_DTMF: TAPI_GETDIGITMODE = "DTMF"
        Case LINEDIGITMODE_DTMFEND: TAPI_GETDIGITMODE = "DTMF END"
        Case Else: TAPI_GETDIGITMODE = "UNKNOWN"
    End Select
End Function

Public Function TAPI_GETCALLSTATE(CALL_STATE As Long, REASON As Long) As String
    Select Case CALL_STATE
        Case LINECALLSTATE_IDLE: TAPI_GETCALLSTATE = "IDLE"
        Case LINECALLSTATE_OFFERING: TAPI_GETCALLSTATE = "OFFERING"
        Case LINECALLSTATE_ACCEPTED: TAPI_GETCALLSTATE = "ACCEPTED"
        Case LINECALLSTATE_DIALTONE: TAPI_GETCALLSTATE = "DIALTONE"
        Case LINECALLSTATE_DIALING: TAPI_GETCALLSTATE = "DIALING"
        Case LINECALLSTATE_RINGBACK: TAPI_GETCALLSTATE = "RINGBACK"
        Case LINECALLSTATE_BUSY: TAPI_GETCALLSTATE = "BUSY"
        Case LINECALLSTATE_SPECIALINFO: TAPI_GETCALLSTATE = "SPECIALINFO"
        Case LINECALLSTATE_CONNECTED: TAPI_GETCALLSTATE = "CONNECTED - " & TAPI_GETCONNECTEDMODE(REASON)
        Case LINECALLSTATE_PROCEEDING: TAPI_GETCALLSTATE = "PROCEDING"
        Case LINECALLSTATE_ONHOLD: TAPI_GETCALLSTATE = "ON HOLD"
        Case LINECALLSTATE_CONFERENCED: TAPI_GETCALLSTATE = "CONFERENCED"
        Case LINECALLSTATE_ONHOLDPENDCONF: TAPI_GETCALLSTATE = "ON HOLD PEND CONF"
        Case LINECALLSTATE_ONHOLDPENDTRANSFER: TAPI_GETCALLSTATE = "ON HOLD PENDING TRANSFER"
        Case LINECALLSTATE_DISCONNECTED: TAPI_GETCALLSTATE = "DISCONNECTED - " & TAPI_GETDISCONNECTMODE(REASON)
        Case Else: TAPI_GETCALLSTATE = "UNKNOWN"
    End Select
End Function

Public Function TAPI_GETCONNECTEDMODE(CONNECTED_MODE As Long) As String
    Select Case CONNECTED_MODE
        Case LINECONNECTEDMODE_ACTIVE: TAPI_GETCONNECTEDMODE = "ACTIVE"
        Case LINECONNECTEDMODE_INACTIVE: TAPI_GETCONNECTEDMODE = "INACTIVE"
        Case LINECONNECTEDMODE_ACTIVEHELD: TAPI_GETCONNECTEDMODE = "ACTIVEHELD"
        Case LINECONNECTEDMODE_INACTIVEHELD: TAPI_GETCONNECTEDMODE = "INACTIVEHELD"
        Case LINECONNECTEDMODE_CONFIRMED: TAPI_GETCONNECTEDMODE = "CONFIRMED"
    End Select
End Function

Public Function TAPI_GETDISCONNECTMODE(DISCONNECT_MODE As Long) As String
    Select Case DISCONNECT_MODE
        Case LINEDISCONNECTMODE_NORMAL: TAPI_GETDISCONNECTMODE = "NORMAL"
        Case LINEDISCONNECTMODE_REJECT: TAPI_GETDISCONNECTMODE = "REJECT"
        Case LINEDISCONNECTMODE_PICKUP: TAPI_GETDISCONNECTMODE = "PICKUP"
        Case LINEDISCONNECTMODE_FORWARDED: TAPI_GETDISCONNECTMODE = "FORWARDED"
        Case LINEDISCONNECTMODE_BUSY: TAPI_GETDISCONNECTMODE = "BUSY"
        Case LINEDISCONNECTMODE_NOANSWER: TAPI_GETDISCONNECTMODE = "NOANSWER"
        Case LINEDISCONNECTMODE_BADADDRESS: TAPI_GETDISCONNECTMODE = "BADADDRESS"
        Case LINEDISCONNECTMODE_UNREACHABLE: TAPI_GETDISCONNECTMODE = "UNREACHABLE"
        Case LINEDISCONNECTMODE_CONGESTION: TAPI_GETDISCONNECTMODE = "CONGESTION"
        Case LINEDISCONNECTMODE_INCOMPATIBLE: TAPI_GETDISCONNECTMODE = "INCOMPATIBLE"
        Case LINEDISCONNECTMODE_UNAVAIL: TAPI_GETDISCONNECTMODE = "UNAVAIL"
        Case Else: TAPI_GETDISCONNECTMODE = "UNKNOWN"
    End Select
End Function

Public Function TAPI_GETCALLPARAM(CALL_PARAM As Long) As String
    Select Case CALL_PARAM
        Case LINECALLPARAMFLAGS_SECURE: TAPI_GETCALLPARAM = "SECURE"
        Case LINECALLPARAMFLAGS_IDLE: TAPI_GETCALLPARAM = "IDLE"
        Case LINECALLPARAMFLAGS_BLOCKID: TAPI_GETCALLPARAM = "BLOCK ID"
        Case LINECALLPARAMFLAGS_ORIGOFFHOOK: TAPI_GETCALLPARAM = "ORIGIN OFF HOOK"
        Case LINECALLPARAMFLAGS_DESTOFFHOOK: TAPI_GETCALLPARAM = "DESTINATION OFF HOOK"
        Case Else: TAPI_GETCALLPARAM = "UNKNOWN"
    End Select
End Function

'Public Function TAPI_GETCALLFLAGS(CALL_FLAGS As Long) As String
'    Select Case CALL_FLAGS
'        Case LINECALLPARTYID_BLOCKED: TAPI_GETCALLFLAGS = "BLOCKED"
'        Case LINECALLPARTYID_OUTOFAREA:  TAPI_GETCALLFLAGS = "OUT OF AREA"
'        Case LINECALLPARTYID_NAME:  TAPI_GETCALLFLAGS = "NAME"
'        Case LINECALLPARTYID_ADDRESS:  TAPI_GETCALLFLAGS = "ADDRESS"
'        Case LINECALLPARTYID_PARTIAL:  TAPI_GETCALLFLAGS = "PARTIAL"
'        Case LINECALLPARTYID_UNAVAIL: TAPI_GETCALLFLAGS = "UNAVAIL"
'        Case Else:  TAPI_GETCALLFLAGS = "UNKNOWN"
'    End Select
'End Function

Public Function TAPI_GETBEARERMODE(BEARER_MODE As Long) As String
    Select Case BEARER_MODE
        Case LINEBEARERMODE_VOICE: TAPI_GETBEARERMODE = "VOICE"
        Case LINEBEARERMODE_SPEECH: TAPI_GETBEARERMODE = "SPEECH"
        Case LINEBEARERMODE_MULTIUSE:  TAPI_GETBEARERMODE = "MULTIUSE"
        Case LINEBEARERMODE_DATA:  TAPI_GETBEARERMODE = "DATA"
        Case LINEBEARERMODE_ALTSPEECHDATA:  TAPI_GETBEARERMODE = "ALTSPEECHDATA"
        Case LINEBEARERMODE_NONCALLSIGNALING:  TAPI_GETBEARERMODE = "NONCALLSIGNALING"
        Case Else: TAPI_GETBEARERMODE = "UNKNOWN"
    End Select
End Function

Public Function TAPI_GETCALLINFOSTATE(CALLINFO_STATE As Long) As String
    Select Case CALLINFO_STATE
        Case LINECALLINFOSTATE_DEVSPECIFIC: TAPI_GETCALLINFOSTATE = "DEV SPECIFIC"
        Case LINECALLINFOSTATE_BEARERMODE: TAPI_GETCALLINFOSTATE = "BEARER MODE"
        Case LINECALLINFOSTATE_RATE: TAPI_GETCALLINFOSTATE = "RATE"
        Case LINECALLINFOSTATE_MEDIAMODE: TAPI_GETCALLINFOSTATE = "MEDIA MODE"
        Case LINECALLINFOSTATE_APPSPECIFIC: TAPI_GETCALLINFOSTATE = "APP SPECIFIC"
        Case LINECALLINFOSTATE_CALLID: TAPI_GETCALLINFOSTATE = "CALL ID"
        Case LINECALLINFOSTATE_RELATEDCALLID: TAPI_GETCALLINFOSTATE = "RELATED CALL ID"
        Case LINECALLINFOSTATE_ORIGIN: TAPI_GETCALLINFOSTATE = "ORIGIN"
        Case LINECALLINFOSTATE_REASON: TAPI_GETCALLINFOSTATE = "REASON"
        Case LINECALLINFOSTATE_COMPLETIONID: TAPI_GETCALLINFOSTATE = "COMPLETION ID"
        Case LINECALLINFOSTATE_NUMOWNERINCR: TAPI_GETCALLINFOSTATE = "NUM OWNER INCR"
        Case LINECALLINFOSTATE_NUMOWNERDECR: TAPI_GETCALLINFOSTATE = "NUM OWNER DECR"
        Case LINECALLINFOSTATE_NUMMONITORS: TAPI_GETCALLINFOSTATE = "NUM MONITORS"
        Case LINECALLINFOSTATE_TRUNK: TAPI_GETCALLINFOSTATE = "TRUNK"
        Case LINECALLINFOSTATE_CALLERID: TAPI_GETCALLINFOSTATE = "CALLER ID"
        Case LINECALLINFOSTATE_CALLEDID: TAPI_GETCALLINFOSTATE = "CALLED ID"
        Case LINECALLINFOSTATE_CONNECTEDID: TAPI_GETCALLINFOSTATE = "CONNECTED ID"
        Case LINECALLINFOSTATE_REDIRECTIONID: TAPI_GETCALLINFOSTATE = "REDIRECTION ID"
        Case LINECALLINFOSTATE_REDIRECTINGID: TAPI_GETCALLINFOSTATE = "REDIRECTING ID"
        Case LINECALLINFOSTATE_DISPLAY: TAPI_GETCALLINFOSTATE = "DISPLAY"
        Case LINECALLINFOSTATE_USERUSERINFO: TAPI_GETCALLINFOSTATE = "USER-USER INFO"
        Case LINECALLINFOSTATE_HIGHLEVELCOMP: TAPI_GETCALLINFOSTATE = "HIGH LEVEL COMP"
        Case LINECALLINFOSTATE_LOWLEVELCOMP: TAPI_GETCALLINFOSTATE = "LOW LEVEL COMP"
        Case LINECALLINFOSTATE_CHARGINGINFO: TAPI_GETCALLINFOSTATE = "CHARGING INFO"
        Case LINECALLINFOSTATE_TERMINAL: TAPI_GETCALLINFOSTATE = "TERMINAL"
        Case LINECALLINFOSTATE_DIALPARAMS: TAPI_GETCALLINFOSTATE = "DIAL PARAMS"
        Case LINECALLINFOSTATE_MONITORMODES: TAPI_GETCALLINFOSTATE = "MONITOR MODES"
        Case Else: TAPI_GETCALLINFOSTATE = "UNKNOWN"
    End Select
End Function

Public Function TAPI_GETADDRESSSTATE(ADDRESS_STATE As Long) As String
    Select Case ADDRESS_STATE
        Case LINEADDRESSSTATE_DEVSPECIFIC: TAPI_GETADDRESSSTATE = "DEVSPECIFIC"
        Case LINEADDRESSSTATE_INUSEZERO: TAPI_GETADDRESSSTATE = "IN USE ZERO"
        Case LINEADDRESSSTATE_INUSEONE: TAPI_GETADDRESSSTATE = "IN USE ONE"
        Case LINEADDRESSSTATE_INUSEMANY: TAPI_GETADDRESSSTATE = "IN USE MANY"
        Case LINEADDRESSSTATE_NUMCALLS: TAPI_GETADDRESSSTATE = "NUMCALLS"
        Case LINEADDRESSSTATE_FORWARD: TAPI_GETADDRESSSTATE = "FORWARD"
        Case LINEADDRESSSTATE_TERMINALS: TAPI_GETADDRESSSTATE = "TERMINALS"
        Case LINEADDRESSSTATE_CAPSCHANGE: TAPI_GETADDRESSSTATE = "CAPSCHANGE"
        Case LINEADDRESSSTATE_OTHER: TAPI_GETADDRESSSTATE = "OTHER"
        Case Else: TAPI_GETADDRESSSTATE = "UNKNOWN"
    End Select
End Function

Public Function TAPI_GETREQUESTMODE(REQUEST_MODE As Long) As String
    Select Case REQUEST_MODE
        Case LINEREQUESTMODE_MAKECALL: TAPI_GETREQUESTMODE = "MAKE CALL"
        Case LINEREQUESTMODE_DROP: TAPI_GETREQUESTMODE = "DROP"
        Case LINEREQUESTMODE_MEDIACALL: TAPI_GETREQUESTMODE = "MEDIA CALL"
        Case Else: TAPI_GETREQUESTMODE = "UNKNOWN"
    End Select
End Function

Public Function TAPI_GETSTRINGFORMAT(STRING_FORAT As Long) As String
    Select Case STRING_FORAT
        Case STRINGFORMAT_ASCII: TAPI_GETSTRINGFORMAT = "ASCII"
        Case STRINGFORMAT_DBCS: TAPI_GETSTRINGFORMAT = "DBCS"
        Case STRINGFORMAT_UNICODE: TAPI_GETSTRINGFORMAT = "UNICODE"
        Case STRINGFORMAT_BINARY: TAPI_GETSTRINGFORMAT = "BINARY"
        Case Else
    End Select
End Function

'Public Sub DeviceStatesArray(ByVal DeviceID As Long, ByRef DeviceState() As Long)
'    Dim ModeSize As Integer
'    ModeSize = -1
'    If LINEDEVSTATE_OTHER And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_OTHER
'    End If
'    If LINEDEVSTATE_RINGING And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_RINGING
'    End If
'    If LINEDEVSTATE_CONNECTED And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_CONNECTED
'    End If
'    If LINEDEVSTATE_DISCONNECTED And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_DISCONNECTED
'    End If
'    If LINEDEVSTATE_MSGWAITON And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_MSGWAITON
'    End If
'    If LINEDEVSTATE_MSGWAITOFF And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_MSGWAITOFF
'    End If
'    If LINEDEVSTATE_NUMCOMPLETIONS And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_NUMCOMPLETIONS
'    End If
'    If LINEDEVSTATE_INSERVICE And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_INSERVICE
'    End If
'    If LINEDEVSTATE_OUTOFSERVICE And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_OUTOFSERVICE
'    End If
'    If LINEDEVSTATE_MAINTENANCE And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_MAINTENANCE
'    End If
'    If LINEDEVSTATE_OPEN And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_OPEN
'    End If
'    If LINEDEVSTATE_CLOSE And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_CLOSE
'    End If
'    If LINEDEVSTATE_NUMCALLS And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_NUMCALLS
'    End If
'    If LINEDEVSTATE_TERMINALS And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_TERMINALS
'    End If
'    If LINEDEVSTATE_ROAMMODE And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_ROAMMODE
'    End If
'    If LINEDEVSTATE_BATTERY And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_BATTERY
'    End If
'    If LINEDEVSTATE_SIGNAL And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_SIGNAL
'    End If
'    If LINEDEVSTATE_DEVSPECIFIC And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_DEVSPECIFIC
'    End If
'    If LINEDEVSTATE_REINIT And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_REINIT
'    End If
'    If LINEDEVSTATE_LOCK And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_LOCK
'    End If
'    If LINEDEVSTATE_CAPSCHANGE And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_CAPSCHANGE
'    End If
'    If LINEDEVSTATE_CONFIGCHANGE And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_CONFIGCHANGE
'    End If
'    If LINEDEVSTATE_TRANSLATECHANGE And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_TRANSLATECHANGE
'    End If
'    If LINEDEVSTATE_COMPLCANCEL And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_COMPLCANCEL
'    End If
'    If LINEDEVSTATE_REMOVED And lineDevCap(DeviceID).dwLineStates Then
'        GoSub Resize
'        DeviceState(ModeSize) = LINEDEVSTATE_REMOVED
'    End If
'Exit Sub
'Resize:
'    ModeSize = ModeSize + 1
'    ReDim Preserve DeviceState(0 To ModeSize) As Long
'    Return
'End Sub
'
'Public Property Get DeviceStates(DeviceID As Long) As Long
'    DeviceStates = lineDevCap(DeviceID).dwLineStates
'End Property
'
'Public Sub RingModesArray(ByVal DeviceID As Long, ByRef RingMode() As Long)
'    Dim ModeSize As Integer
'    ModeSize = -1
'    If LINETONEMODE_BEEP And lineDevCap(DeviceID).dwRingModes Then
'        GoSub Resize
'        RingMode(ModeSize) = LINETONEMODE_BEEP
'    End If
'    If LINETONEMODE_BILLING And lineDevCap(DeviceID).dwRingModes Then
'        GoSub Resize
'        RingMode(ModeSize) = LINETONEMODE_BILLING
'    End If
'    If LINETONEMODE_BUSY And lineDevCap(DeviceID).dwRingModes Then
'        GoSub Resize
'        RingMode(ModeSize) = LINETONEMODE_BUSY
'    End If
'    If LINETONEMODE_CUSTOM And lineDevCap(DeviceID).dwRingModes Then
'        GoSub Resize
'        RingMode(ModeSize) = LINETONEMODE_CUSTOM
'    End If
'    If LINETONEMODE_RINGBACK And lineDevCap(DeviceID).dwRingModes Then
'        GoSub Resize
'        RingMode(ModeSize) = LINETONEMODE_RINGBACK
'    End If
'Exit Sub
'Resize:
'    ModeSize = ModeSize + 1
'    ReDim Preserve RingMode(0 To ModeSize) As Long
'    Return
'End Sub
'
'Public Property Get RingModes(DeviceID As Long) As Long
'    RingModes = lineDevCap(DeviceID).dwRingModes
'End Property
'
'Public Sub MonitorDigitModesArray(ByVal DeviceID As Long, ByRef MonitorDigitMode() As Long)
'    Dim ModeSize As Integer
'    ModeSize = -1
'    If LINEDIGITMODE_PULSE And lineDevCap(DeviceID).dwMonitorDigitModes Then
'        GoSub Resize
'        MonitorDigitMode(ModeSize) = LINEDIGITMODE_PULSE
'    End If
'    If LINEDIGITMODE_DTMF And lineDevCap(DeviceID).dwMonitorDigitModes Then
'        GoSub Resize
'        MonitorDigitMode(ModeSize) = LINEDIGITMODE_DTMF
'    End If
'    If LINEDIGITMODE_DTMFEND And lineDevCap(DeviceID).dwMonitorDigitModes Then
'        GoSub Resize
'        MonitorDigitMode(ModeSize) = LINEDIGITMODE_DTMFEND
'    End If
'Exit Sub
'Resize:
'    ModeSize = ModeSize + 1
'    ReDim Preserve MonitorDigitMode(0 To ModeSize) As Long
'    Return
'End Sub
'
'Public Property Get MonitorDigitModes(DeviceID As Long) As Long
'    MonitorDigitModes = lineDevCap(DeviceID).dwMonitorDigitModes
'End Property
'
'Public Sub GenerateDigitModesArray(ByVal DeviceID As Long, ByRef GenerateDigitMode() As Long)
'    Dim ModeSize As Integer
'    ModeSize = -1
'    If LINEDIGITMODE_PULSE And lineDevCap(DeviceID).dwGenerateDigitModes Then
'        GoSub Resize
'        GenerateDigitMode(ModeSize) = LINEDIGITMODE_PULSE
'    End If
'    If LINEDIGITMODE_DTMF And lineDevCap(DeviceID).dwGenerateDigitModes Then
'        GoSub Resize
'        GenerateDigitMode(ModeSize) = LINEDIGITMODE_DTMF
'    End If
'Exit Sub
'Resize:
'    ModeSize = ModeSize + 1
'    ReDim Preserve GenerateDigitMode(0 To ModeSize) As Long
'    Return
'End Sub
'
'Public Property Get GenerateDigitModes(DeviceID As Long) As Long
'    GenerateDigitModes = lineDevCap(DeviceID).dwGenerateDigitModes
'End Property
'
'Public Sub AnswerModesArray(ByVal DeviceID As Long, ByRef AnswerMode() As Long)
'    Dim ModeSize As Integer
'    ModeSize = -1
'    If LINEANSWERMODE_NONE And lineDevCap(DeviceID).dwBearerModes Then
'        GoSub Resize
'        AnswerMode(ModeSize) = LINEANSWERMODE_NONE
'    End If
'    If LINEANSWERMODE_DROP And lineDevCap(DeviceID).dwAnswerMode Then
'        GoSub Resize
'        AnswerMode(ModeSize) = LINEANSWERMODE_DROP
'    End If
'    If LINEANSWERMODE_HOLD And lineDevCap(DeviceID).dwAnswerMode Then
'        GoSub Resize
'        AnswerMode(ModeSize) = LINEANSWERMODE_HOLD
'    End If
'Exit Sub
'Resize:
'    ModeSize = ModeSize + 1
'    ReDim Preserve AnswerMode(0 To ModeSize) As Long
'    Return
'End Sub
'
'Public Property Get AnswerModes(DeviceID As Long) As Long
'    AnswerModes = lineDevCap(DeviceID).dwAnswerMode
'End Property
'
'Public Sub BearerModesArray(ByVal DeviceID As Long, ByRef BearerMode() As Long)
'    Dim ModeSize As Integer
'    ModeSize = -1
'    If LINEBEARERMODE_VOICE And lineDevCap(DeviceID).dwBearerModes Then
'        GoSub Resize
'        BearerMode(ModeSize) = LINEBEARERMODE_VOICE
'    End If
'    If LINEBEARERMODE_SPEECH And lineDevCap(DeviceID).dwBearerModes Then
'        GoSub Resize
'        BearerMode(ModeSize) = LINEBEARERMODE_SPEECH
'    End If
'    If LINEBEARERMODE_DATA And lineDevCap(DeviceID).dwBearerModes Then
'        GoSub Resize
'        BearerMode(ModeSize) = LINEBEARERMODE_DATA
'    End If
'    If LINEBEARERMODE_ALTSPEECHDATA And lineDevCap(DeviceID).dwBearerModes Then
'        GoSub Resize
'        BearerMode(ModeSize) = LINEBEARERMODE_ALTSPEECHDATA
'    End If
'    If LINEBEARERMODE_MULTIUSE And lineDevCap(DeviceID).dwBearerModes Then
'        GoSub Resize
'        BearerMode(ModeSize) = LINEBEARERMODE_MULTIUSE
'    End If
'    If LINEBEARERMODE_NONCALLSIGNALING And lineDevCap(DeviceID).dwBearerModes Then
'        GoSub Resize
'        BearerMode(ModeSize) = LINEBEARERMODE_NONCALLSIGNALING
'    End If
'Exit Sub
'Resize:
'    ModeSize = ModeSize + 1
'    ReDim Preserve BearerMode(0 To ModeSize) As Long
'    Return
'End Sub
'
'Public Property Get BearerModes(DeviceID As Long) As Long
'    BearerModes = lineDevCap(DeviceID).dwBearerModes
'End Property
'
'Public Sub MediaModesArray(ByVal DeviceID As Long, ByRef MediaMode() As Long)
'    Dim ModeSize As Integer
'    ModeSize = -1
'    If LINEMEDIAMODE_ADSI And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_ADSI
'    End If
'    If LINEMEDIAMODE_AUTOMATEDVOICE And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_AUTOMATEDVOICE
'    End If
'    If LINEMEDIAMODE_DATAMODEM And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_DATAMODEM
'    End If
'    If LINEMEDIAMODE_G3FAX And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_G3FAX
'    End If
'    If LINEMEDIAMODE_G4FAX And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_G4FAX
'    End If
'    If LINEMEDIAMODE_INTERACTIVEVOICE And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_INTERACTIVEVOICE
'    End If
'    If LINEMEDIAMODE_MIXED And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_MIXED
'    End If
'    If LINEMEDIAMODE_TDD And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_TDD
'    End If
'    If LINEMEDIAMODE_TELETEX And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_TELETEX
'    End If
'    If LINEMEDIAMODE_TELEX And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_TELEX
'    End If
'    If LINEMEDIAMODE_UNKNOWN And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_UNKNOWN
'    End If
'    If LINEMEDIAMODE_VIDEOTEX And lineDevCap(DeviceID).dwMediaModes Then
'        GoSub Resize
'        MediaMode(ModeSize) = LINEMEDIAMODE_VIDEOTEX
'    End If
'Exit Sub
'Resize:
'    ModeSize = ModeSize + 1
'    ReDim Preserve MediaMode(0 To ModeSize) As Long
'    Return
'End Sub
'
'Public Property Get MediaModes(DeviceID As Long) As Long
'    MediaModes = lineDevCap(DeviceID).dwMediaModes
'End Property
'
'Public Sub DevCapFlagsArray(ByVal DeviceID As Long, ByRef DevCapFlag() As Long)
'    Dim ModeSize As Integer
'    ModeSize = -1
'    If LINEDEVCAPFLAGS_CROSSADDRCONF And lineDevCap(DeviceID).dwDevCapFlags Then
'        GoSub Resize
'        DevCapFlag(ModeSize) = LINEDEVCAPFLAGS_CROSSADDRCONF
'    End If
'    If LINEDEVCAPFLAGS_HIGHLEVCOMP And lineDevCap(DeviceID).dwDevCapFlags Then
'        GoSub Resize
'        DevCapFlag(ModeSize) = LINEDEVCAPFLAGS_HIGHLEVCOMP
'    End If
'    If LINEDEVCAPFLAGS_LOWLEVCOMP And lineDevCap(DeviceID).dwDevCapFlags Then
'        GoSub Resize
'        DevCapFlag(ModeSize) = LINEDEVCAPFLAGS_LOWLEVCOMP
'    End If
'    If LINEDEVCAPFLAGS_MEDIACONTROL And lineDevCap(DeviceID).dwDevCapFlags Then
'        GoSub Resize
'        DevCapFlag(ModeSize) = LINEDEVCAPFLAGS_MEDIACONTROL
'    End If
'    If LINEDEVCAPFLAGS_MULTIPLEADDR And lineDevCap(DeviceID).dwDevCapFlags Then
'        GoSub Resize
'        DevCapFlag(ModeSize) = LINEDEVCAPFLAGS_MULTIPLEADDR
'    End If
'    If LINEDEVCAPFLAGS_CLOSEDROP And lineDevCap(DeviceID).dwDevCapFlags Then
'        GoSub Resize
'        DevCapFlag(ModeSize) = LINEDEVCAPFLAGS_CLOSEDROP
'    End If
'    If LINEDEVCAPFLAGS_DIALBILLING And lineDevCap(DeviceID).dwDevCapFlags Then
'        GoSub Resize
'        DevCapFlag(ModeSize) = LINEDEVCAPFLAGS_DIALBILLING
'    End If
'    If LINEDEVCAPFLAGS_DIALQUIET And lineDevCap(DeviceID).dwDevCapFlags Then
'        GoSub Resize
'        DevCapFlag(ModeSize) = LINEDEVCAPFLAGS_DIALQUIET
'    End If
'    If LINEDEVCAPFLAGS_DIALDIALTONE And lineDevCap(DeviceID).dwDevCapFlags Then
'        GoSub Resize
'        DevCapFlag(ModeSize) = LINEDEVCAPFLAGS_DIALDIALTONE
'    End If
'Exit Sub
'Resize:
'    ModeSize = ModeSize + 1
'    ReDim Preserve DevCapFlag(0 To ModeSize) As Long
'    Return
'End Sub
'
'Public Property Get DevCapFlags(DeviceID As Long) As Long
'    DevCapFlags = lineDevCap(DeviceID).dwDevCapFlags
'End Property

'Private Function GetTAPIStructString(ByVal ptrTapistruct As Long, ByVal offset As Long, ByVal length As Long) As String
'    'ugly C-hacker way to deal with ugly C-hacker TAPI structs (UDTs)
'    Dim buffer() As Byte
'
'    If length < 1 Then Exit Function 'handle erroneous input
'
'    If offset Then '
'        ReDim buffer(0 To length - 1)
'        CopyMemory buffer(0), ByVal ptrTapistruct + offset, length
'        GetTAPIStructString = StrConv(buffer, vbUnicode)
'    End If
'
'End Function

