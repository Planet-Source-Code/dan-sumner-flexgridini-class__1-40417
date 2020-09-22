VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "flexgridclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


'local variable(s) to hold property value(s)
Private mvarininame As String 'local copy
Private mvarGridObj As Object 'local copy
Private mvarKey As String 'local copy
Private mvarSection As String 'local copy
Private mvarstartup As Boolean
Private mvariniobj As iniclass

Public Sub GetColumnWidths()
Dim w As String
Dim x As Integer
Dim wid As Long

For x = 0 To mvarGridObj.Cols - 1
   w = "Col#" & LTrim(Str(x))
   mvariniobj.Key = w
   wid = mvariniobj.readData
   mvarGridObj.ColWidth(x) = wid
Next

End Sub

Public Sub SavecolumnWidths()
Dim w As String
Dim x As Integer
Dim wid As Long

For x = 0 To mvarGridObj.Cols - 1
   w = "Col#" & LTrim(Str(x))
   mvariniobj.Key = w
   wid = mvarGridObj.ColWidth(x)
   mvariniobj.writeData (wid)
Next

End Sub

Public Property Let Section(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Section = 5
    mvarSection = vData
    mvariniobj.Section = vData
End Property


Public Property Get Section() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Section
    Section = mvarSection
End Property



Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
    mvariniobj.Key = vData
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property



Public Property Set GridObj(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.GridObj = Form1
    Set mvarGridObj = vData
End Property


Public Property Get GridObj() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GridObj
    Set GridObj = mvarGridObj
End Property



Public Property Let ininame(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ininame = 5
    mvarininame = vData
    If mvarstartup = False Then
      mvarstartup = True
      Set mvariniobj = New iniclass
    End If
    mvariniobj.file = mvarininame
    
End Property


Public Property Get ininame() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ininame
    ininame = mvarininame
End Property



