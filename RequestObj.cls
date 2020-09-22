VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RequestObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

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

Private Declare Function lineGetRequest Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwRequestMode As Long, lpRequestBuffer As Any) As Long

Dim lineRequestMediaCall As LINEREQMEDIACALL
Dim lineRequestMakeCall As LINEREQMAKECALL

Public Property Get DestinationAddress() As String
    Select Case Module1.RequestMode
        Case LINEREQUESTMODE_MAKECALL
            If InStr(1, lineRequestMakeCall.szDestAddress, Chr(0)) > 0 Then
                DestinationAddress = Left( _
                    lineRequestMakeCall.szDestAddress, _
                    InStr( _
                        1, _
                        lineRequestMakeCall.szDestAddress, _
                        Chr(0) _
                    ) - 1 _
                )
            End If
        Case LINEREQUESTMODE_MEDIACALL
            If InStr(1, lineRequestMediaCall.szDestAddress, Chr(0)) > 0 Then
                DestinationAddress = Left( _
                    lineRequestMediaCall.szDestAddress, _
                    InStr( _
                        1, _
                        lineRequestMediaCall.szDestAddress, _
                        Chr(0) _
                    ) - 1 _
                )
            End If
    End Select
End Property

Public Property Get CalledParty() As String
    Select Case Module1.RequestMode
        Case LINEREQUESTMODE_MAKECALL
            If InStr(1, lineRequestMakeCall.szCalledParty, Chr(0)) > 0 Then
                CalledParty = Left( _
                    lineRequestMakeCall.szCalledParty, _
                    InStr( _
                        1, _
                        lineRequestMakeCall.szCalledParty, _
                        Chr(0) _
                    ) - 1 _
                )
            End If
        Case LINEREQUESTMODE_MEDIACALL
            If InStr(1, lineRequestMediaCall.szCalledParty, Chr(0)) > 0 Then
                CalledParty = Left( _
                    lineRequestMediaCall.szCalledParty, _
                    InStr( _
                        1, _
                        lineRequestMediaCall.szCalledParty, _
                        Chr(0) _
                    ) - 1 _
                )
            End If
    End Select
End Property

Public Property Get Comment() As String
    Select Case Module1.RequestMode
        Case LINEREQUESTMODE_MAKECALL
            If InStr(1, lineRequestMakeCall.szComment, Chr(0)) > 0 Then
                Comment = Left( _
                    lineRequestMakeCall.szComment, _
                    InStr( _
                        1, _
                        lineRequestMakeCall.szComment, _
                        Chr(0) _
                    ) - 1 _
                )
            End If
        Case LINEREQUESTMODE_MEDIACALL
            If InStr(1, lineRequestMediaCall.szComment, Chr(0)) > 0 Then
                Comment = Left( _
                    lineRequestMediaCall.szComment, _
                    InStr( _
                        1, _
                        lineRequestMediaCall.szComment, _
                        Chr(0) _
                    ) - 1 _
                )
            End If
    End Select
End Property

Public Property Get ApplicationName() As String
    Select Case Module1.RequestMode
        Case LINEREQUESTMODE_MAKECALL
            If InStr(1, lineRequestMakeCall.szAppName, Chr(0)) > 0 Then
                ApplicationName = Left( _
                    lineRequestMakeCall.szAppName, _
                    InStr( _
                        1, _
                        lineRequestMakeCall.szAppName, _
                        Chr(0) _
                    ) - 1 _
                )
            End If
        Case LINEREQUESTMODE_MEDIACALL
            If InStr(1, lineRequestMediaCall.szAppName, Chr(0)) > 0 Then
                ApplicationName = Left( _
                    lineRequestMediaCall.szAppName, _
                    InStr( _
                        1, _
                        lineRequestMediaCall.szAppName, _
                        Chr(0) _
                    ) - 1 _
                )
            End If
    End Select
End Property

Public Property Get Secure() As Long
    Select Case Module1.RequestMode
        Case LINEREQUESTMODE_MEDIACALL:  Secure = lineRequestMediaCall.dwSecure
    End Select
End Property

Public Property Get DeviceID() As String
    Select Case Module1.RequestMode
        Case LINEREQUESTMODE_MEDIACALL
            If InStr(1, lineRequestMediaCall.ucDeviceID, Chr(0)) > 0 Then
                DeviceID = Left( _
                    lineRequestMediaCall.ucDeviceID, _
                    InStr( _
                        1, _
                        lineRequestMediaCall.ucDeviceID, _
                        Chr(0) _
                    ) - 1 _
                )
            End If
    End Select
End Property

Public Property Get DeviceClass() As String
    Select Case Module1.RequestMode
        Case LINEREQUESTMODE_MEDIACALL
            If InStr(1, lineRequestMediaCall.szDeviceClass, Chr(0)) > 0 Then
                DeviceClass = Left( _
                    lineRequestMediaCall.szDeviceClass, _
                    InStr( _
                        1, _
                        lineRequestMediaCall.szDeviceClass, _
                        Chr(0) _
                    ) - 1 _
                )
            End If
    End Select
End Property

Public Property Get RequestID() As Long
    Select Case Module1.RequestMode
        Case LINEREQUESTMODE_MEDIACALL:  RequestID = lineRequestMediaCall.wRequestID
    End Select
End Property

Private Sub Class_Initialize()
    Call Func_lineGetRequest
End Sub

Private Function Func_lineGetRequest() As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Debug print the call back message to file '''''''''''''''''''''''''''''
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineGetRequest"
        Close #FreeFileNo
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Select Case Module1.RequestMode
        Case LINEREQUESTMODE_MAKECALL
            Func_lineGetRequest = lineGetRequest( _
                hndLine, _
                LINEREQUESTMODE_MAKECALL, _
                lineRequestMakeCall _
            )
        Case LINEREQUESTMODE_MEDIACALL
            Func_lineGetRequest = lineGetRequest( _
                hndLine, _
                LINEREQUESTMODE_MEDIACALL, _
                lineRequestMediaCall _
            )
    End Select
End Function

Public Property Get RequestType() As String
    Select Case Module1.RequestMode
        Case LINEREQUESTMODE_MAKECALL: RequestType = "MAKECALL"
        Case LINEREQUESTMODE_MEDIACALL: RequestType = "MEDIACALL"
        Case LINEREQUESTMODE_DROP: RequestType = "DROP"
    End Select
End Property
