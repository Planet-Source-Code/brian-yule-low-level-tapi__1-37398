VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceItemConfigObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Newone"
Option Explicit

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


Private Const VARSTRING_FIXEDSIZE = 24

Private Type VARSTRING
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwStringFormat As Long
    dwStringSize As Long
    dwStringOffset As Long

    mem As String * 4096 ' added by mca
End Type

'"comm" Communications port.
'"comm/datamodem" Modem through a communications port.
'"comm/datamodem/portname" Name of the device to which a modem is connected.
'"wave/in"  Wave audio device (input only).
'"wave/out" Wave audio device (output only).
'"midi/in" Midi sequencer (input only).
'"midi/out" Midi sequencer (output only).
'"tapi/line" Line device.
'"tapi/phone" Phone device
'"ndis" Network device.
'"tapi/terminal" Terminal device.

Private Declare Function lineGetDevConfig Lib "TAPI32.DLL" (ByVal dwDeviceID As Long, lpDeviceConfig As VARSTRING, ByVal lpszDeviceClass As String) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

Dim Index As Long

'Private Function GetDevConfigString(mem As String, offset As Long, size As Long) As String
'    GetDevConfigString = Trim(Replace(Replace(Mid(mem, offset + 1 - VARSTRING_FIXEDSIZE, size - 1), Chr(0), " "), "|", " "))
'End Function

Public Property Get comm() As String
    Dim deviceConfig As VARSTRING
    Let deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "comm")
    If deviceConfig.dwStringSize <> 0 Then
        comm = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get comm_FORMAT() As String
    Dim deviceConfig As VARSTRING
    Let deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "comm")
    comm_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Public Property Get comm_datamodem() As String
    Dim deviceConfig As VARSTRING
    Let deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "comm/datamodem")
    If deviceConfig.dwStringSize <> 0 Then
        comm_datamodem = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get comm_datamodem_FORMAT() As String
    Dim deviceConfig As VARSTRING
    Let deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "comm/datamodem")
    comm_datamodem_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Public Property Get comm_datamodem_portname() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "comm/datamodem/portname")
    If deviceConfig.dwStringSize <> 0 Then
        comm_datamodem_portname = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get comm_datamodem_portname_FORMAT() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "comm/datamodem/portname")
    comm_datamodem_portname_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Public Property Get wave_in() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "wave/in")
    If deviceConfig.dwStringSize <> 0 Then
        wave_in = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get wave_in_FORMAT() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "wave/in")
    wave_in_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Public Property Get wave_out() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "wave/out")
    If deviceConfig.dwStringSize <> 0 Then
        wave_out = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get wave_out_FORMAT() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "wave/out")
    wave_out_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Public Property Get midi_in() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "midi/in")
    If deviceConfig.dwStringSize <> 0 Then
        midi_in = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get midi_in_FORMAT() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "midi/in")
    midi_in_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Public Property Get midi_out() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "midi/out")
    If deviceConfig.dwStringSize <> 0 Then
        midi_out = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get midi_out_FORMAT() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "midi/out")
    midi_out_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Public Property Get tapi_line() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "tapi/line")
    If deviceConfig.dwStringSize <> 0 Then
        tapi_line = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get tapi_line_FORMAT() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "tapi/line")
    tapi_line_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Public Property Get tapi_phone() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "tapi/phone")
    If deviceConfig.dwStringSize <> 0 Then
        tapi_phone = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get tapi_phone_FORMAT() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "tapi/phone")
    tapi_phone_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Public Property Get ndis() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "ndis")
    If deviceConfig.dwStringSize <> 0 Then
        ndis = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get ndis_FORMAT() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "ndis")
    ndis_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Public Property Get tapi_terminal() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "tapi/terminal")
    If deviceConfig.dwStringSize <> 0 Then
        tapi_terminal = GetDevConfigString(VarPtr(deviceConfig), deviceConfig.dwStringOffset, deviceConfig.dwStringSize)
    End If
End Property

Public Property Get tapi_terminal_FORMAT() As String
    Dim deviceConfig As VARSTRING
    deviceConfig.dwTotalSize = Len(deviceConfig)
    Call lineGetDevConfig(Index, deviceConfig, "tapi/terminal")
    tapi_terminal_FORMAT = TAPI_GETSTRINGFORMAT(deviceConfig.dwStringFormat)
End Property

Private Function GetDevConfigString(ByVal ptrTapiStruct As Long, ByVal offset As Long, ByVal Length As Long) As String
    Dim Buffer() As Byte
    ReDim Buffer(0 To Length) As Byte
    CopyMemory Buffer(0), ByVal ptrTapiStruct + offset, Length
    GetDevConfigString = Buffer
End Function

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

