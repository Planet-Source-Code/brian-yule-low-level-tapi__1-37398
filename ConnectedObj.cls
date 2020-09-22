VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectedObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private Const LINECALLPARTYID_BLOCKED = &H1&
Private Const LINECALLPARTYID_OUTOFAREA = &H2&
Private Const LINECALLPARTYID_NAME = &H4&
Private Const LINECALLPARTYID_ADDRESS = &H8&
Private Const LINECALLPARTYID_PARTIAL = &H10&
Private Const LINECALLPARTYID_UNKNOWN = &H20&
Private Const LINECALLPARTYID_UNAVAIL = &H40&

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

Private Declare Function lineGetCallInfo Lib "TAPI32.DLL" (ByVal hCall As Long, ByRef lpCallInfo As lineCallInfo) As Long

Private Const LINECALLINFO_FIXEDSIZE = 296

Dim Index As Long
Dim CallIndex As Long

Public Property Get ID() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwConnectedIDSize <> 0 Then
        ID = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwConnectedIDOffset, lineCallInformation.dwConnectedIDSize)
    End If
End Property

Public Property Get Name() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    If lineCallInformation.dwConnectedIDNameSize <> 0 Then
        Name = GetCallInfoString(lineCallInformation.mem, lineCallInformation.dwConnectedIDNameOffset, lineCallInformation.dwConnectedIDNameSize)
    End If
End Property

Public Property Get Status() As String
    Dim lineCallInformation As lineCallInfo, retVal As Long

    lineCallInformation.dwTotalSize = Len(lineCallInformation)
    retVal = lineGetCallInfo(DeviceItemHndCall(Index, CallIndex), lineCallInformation)

    Status = TAPI_GETCALLPARTYID(lineCallInformation.dwConnectedIDFlags)
End Property

Private Function GetCallInfoString(mem As String, offset As Long, size As Long) As String
    GetCallInfoString = Trim(Replace(Replace(Mid(mem, offset + 1 - LINECALLINFO_FIXEDSIZE, size - 1), Chr(0), " "), "|", " "))
End Function

Private Sub Class_Initialize()
    Index = DeviceItem
    CallIndex = CallItem
End Sub

Public Function TAPI_GETCALLPARTYID(CALLPARTYID) As String
    Select Case CALLPARTYID
        Case LINECALLPARTYID_BLOCKED: TAPI_GETCALLPARTYID = "BLOCKED"
        Case LINECALLPARTYID_OUTOFAREA: TAPI_GETCALLPARTYID = "OUTOFAREA"
        Case LINECALLPARTYID_NAME: TAPI_GETCALLPARTYID = "NAME"
        Case LINECALLPARTYID_ADDRESS: TAPI_GETCALLPARTYID = "ADDRESS"
        Case LINECALLPARTYID_PARTIAL: TAPI_GETCALLPARTYID = "PARTIAL"
        Case LINECALLPARTYID_UNAVAIL: TAPI_GETCALLPARTYID = "UNAVAIL"
        Case Else: TAPI_GETCALLPARTYID = "UNKNOWN"
    End Select
End Function

