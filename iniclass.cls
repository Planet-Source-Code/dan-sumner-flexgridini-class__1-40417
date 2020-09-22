VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iniclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

 

'API
Private Declare Function GetPrivateProfileInt Lib "kernel32.dll" _
Alias "GetPrivateProfileIntA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As String, _
ByVal nDefault As Long, _
ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32.dll" _
Alias "GetPrivateProfileStringA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32.dll" _
Alias "WritePrivateProfileStringA" ( _
ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpString As Any, _
ByVal lpFileName As String) As Long

Private sFile As String
'local variable(s) to hold property value(s)
Private mvarSection As String 'local copy
Private mvarKey As String 'local copy
Private mvarfile As String 'local copy

Public Property Let file(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.file = 5
    mvarfile = vData
End Property

Public Property Get file() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.file
      file = mvarfile
End Property

Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property

Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
        Key = mvarKey
End Property



Public Property Let Section(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Section = 5
    mvarSection = vData
End Property


Public Property Get Section() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Section
    Section = mvarSection
End Property

Public Function readData() As Integer
    readData = GetPrivateProfileInt(mvarSection, mvarKey, -1, mvarfile)
End Function

Public Sub writeData(ByVal iData As Integer)
    Dim sData As String
    sData = iData
    WritePrivateProfileString mvarSection, mvarKey, sData, mvarfile
End Sub

Public Function readText() As String
    Dim sText As String
    Dim lResult As Long
    sText = String$(255, 0)
    lResult = GetPrivateProfileString(mvarSection, mvarKey, "", sText, Len(sText), mvarfile)
    If lResult = 0 Then
        readText = ""
    Else
        readText = Left(sText, InStr(sText, Chr(0)) - 1)
    End If
End Function

Public Sub writetext(ByVal iText As String)
    Dim sText As String
    sText = iText
    WritePrivateProfileString mvarSection, mvarKey, sText, mvarfile
End Sub
'*****************************************************************
 

 
 

 

 
 


