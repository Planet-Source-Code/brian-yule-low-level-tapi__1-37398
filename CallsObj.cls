VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallsObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is a collection of calls for the line device."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CallObj"
Attribute VB_Ext_KEY = "Member0" ,"CallObj"
Option Explicit

Private mCol As Collection

Public Function Add(newCall As CallObj, Optional sKey As Variant) As CallObj
    If Len(sKey) = 0 Then mCol.Add newCall Else mCol.Add newCall, CStr(sKey)
    Set Add = newCall
End Function

Public Property Get Item(vntIndexKey As Variant) As CallObj
    If Len(vntIndexKey) > 0 Then Set Item = mCol.Item(CStr(vntIndexKey))
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    If Len(vntIndexKey) > 0 Then mCol.Remove CStr(vntIndexKey)
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
