VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppPriorityObj"
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

Private Type LINEEXTENSIONID
    dwExtensionID0 As Long
    dwExtensionID1 As Long
    dwExtensionID2 As Long
    dwExtensionID3 As Long
End Type

Private Declare Function lineGetAppPriority Lib "TAPI32.DLL" (ByVal lpszAppFilename As String, ByVal dwMediaMode As Long, lpExtensionID As LINEEXTENSIONID, ByVal dwRequestMode As Long, lpExtensionName As String, lpdwPriority As Long) As Long
Private Declare Function lineSetAppPriority Lib "TAPI32.DLL" (ByVal lpszAppFilename As String, ByVal dwMediaMode As Long, lpExtensionID As LINEEXTENSIONID, ByVal dwRequestMode As Long, ByVal lpszExtensionName As String, ByVal dwPriority As Long) As Long

Public Property Get ADSI() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_ADSI, Priority)
    
    ADSI = CBool(Priority)
End Property

Public Property Get Mixed() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_MIXED, Priority)
    
    Mixed = CBool(Priority)
End Property

Public Property Get Telex() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_TELEX, Priority)
    
    Telex = CBool(Priority)
End Property

Public Property Get VideoTex() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_VIDEOTEX, Priority)
    
    VideoTex = CBool(Priority)
End Property

Public Property Get TeleTex() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_TELETEX, Priority)
    
    TeleTex = CBool(Priority)
End Property

Public Property Get DigitalData() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_DIGITALDATA, Priority)
    
    DigitalData = CBool(Priority)
End Property

Public Property Get G4FAX() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_G4FAX, Priority)
    
    G4FAX = CBool(Priority)
End Property

Public Property Get TDD() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_TDD, Priority)
    
    TDD = CBool(Priority)
End Property

Public Property Get G3Fax() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_G3FAX, Priority)
    
    G3Fax = CBool(Priority)
End Property

Public Property Get DataModem() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_DATAMODEM, Priority)
    
    DataModem = CBool(Priority)
End Property

Public Property Get AutomatedVoice() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_AUTOMATEDVOICE, Priority)
    
    AutomatedVoice = CBool(Priority)
End Property

Public Property Get InteractiveVoice() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_INTERACTIVEVOICE, Priority)
    
    InteractiveVoice = CBool(Priority)
End Property

Public Property Get Unknown() As Boolean
    Dim Priority As Long
        
    Call Func_lineGetAppPriority(LINEMEDIAMODE_UNKNOWN, Priority)
    
    Unknown = CBool(Priority)
End Property

Private Function Func_lineGetAppPriority(ByVal Mode As Long, ByRef Priority As Long) As Long
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineGetAppPriority"
        Close #FreeFileNo
    End If
    
    Dim LineExtID  As LINEEXTENSIONID
    
    Func_lineGetAppPriority = lineGetAppPriority( _
        App.EXEName & vbNullString, _
        Mode, _
        LineExtID, _
        LINEREQUESTMODE_MAKECALL, _
        "", _
        Priority _
    )
End Function

Private Function Func_lineSetAppPriority(ByVal Mode As Long, ByVal Priority As Long) As Long
    If Module1.DebugMode Then
        Dim AppPath$, FreeFileNo%
        
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
        
        Let FreeFileNo = FreeFile
        
        Open AppPath & "\CallBackInstance.log" For Append As #FreeFileNo
            Print #FreeFileNo, Now, "Executing: lineSetAppPriority"
        Close #FreeFileNo
    End If
    
    Dim LineExtID  As LINEEXTENSIONID
    
    Func_lineSetAppPriority = lineSetAppPriority( _
        App.EXEName & vbNullString, _
        Mode, _
        LineExtID, _
        LINEREQUESTMODE_MAKECALL, _
        "", _
        Priority _
    )
End Function

Public Property Let ADSI(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_ADSI, Abs(CInt(Priority)))
End Property

Public Property Let Mixed(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_MIXED, Abs(CInt(Priority)))
End Property

Public Property Let Telex(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_TELEX, Abs(CInt(Priority)))
End Property

Public Property Let VideoTex(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_VIDEOTEX, Abs(CInt(Priority)))
End Property

Public Property Let TeleTex(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_TELETEX, Abs(CInt(Priority)))
End Property

Public Property Let DigitalData(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_DIGITALDATA, Abs(CInt(Priority)))
End Property

Public Property Let G4FAX(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_G4FAX, Abs(CInt(Priority)))
End Property

Public Property Let TDD(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_TDD, Abs(CInt(Priority)))
End Property

Public Property Let G3Fax(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_G3FAX, Abs(CInt(Priority)))
End Property

Public Property Let DataModem(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_DATAMODEM, Abs(CInt(Priority)))
End Property

Public Property Let AutomatedVoice(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_AUTOMATEDVOICE, Abs(CInt(Priority)))
End Property

Public Property Let InteractiveVoice(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_INTERACTIVEVOICE, Abs(CInt(Priority)))
End Property

Public Property Let Unknown(Priority As Boolean)
    Call Func_lineSetAppPriority(LINEMEDIAMODE_UNKNOWN, Abs(CInt(Priority)))
End Property
