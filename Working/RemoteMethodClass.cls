VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemoteMethodClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'=================================================================================
'   This object builds the SiteExplorer navigator
'
'   It is a webservice the responds with a list of child menus that are linked
'   back to the webservice for each of its nodes
'
'=================================================================================
'
' ----- Global Objects
'
Private Main As Object
'Private Main As MainClass
'
Const IconClosedDefault = "<img src=""/cclib/images/BookClosed1818.gif"" width=18 height=18 border=0 style=""vertical-align:middle;"">"
Const IconOpenDefault = "<img src=""/cclib/images/BookOpen1818.gif"" width=18 height=18 border=0 style=""vertical-align:middle;"">"
Const IconEmptyDefault = "<img src=""/cclib/images/IconPage1818.gif"" width=18 height=18 border=0 style=""vertical-align:middle;"">"
Const IconClosedBrokenDefault = "<img src=""/cclib/images/BookClosedBroken1818.gif"" width=18 height=18 border=0 style=""vertical-align:middle;"" title=""Error"">"
Const IconOpenBrokenDefault = "<img src=""/cclib/images/BookOpenBroken1818.gif"" width=18 height=18 border=0 style=""vertical-align:middle;"" title=""Error"">"
Const IconEmptyBrokenDefault = "<img src=""/cclib/images/IconPageBroken1818.gif"" width=18 height=18 border=0 style=""vertical-align:middle;"" title=""Error"">"
'
'=================================================================================
' Initilize this object
'   Must be called after creation, before any other calls
'=================================================================================
'
Public Sub Init(MainObject As Object)
    Set Main = MainObject
    End Sub
'
'=================================================================================
'   Aggregate Object Interface
'=================================================================================
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim InstanceID As Long
    '
    InstanceID = kmaEncodeInteger(Main.GetAggrOption("InstanceID", OptionString))
    OptionString = OptionString _
        & "&" & "IconClosed=" & IconClosedDefault _
        & "&" & "IconOpen=" & IconOpenDefault _
        & "&" & "IconEmpty=" & IconEmptyDefault _
        & "&" & "IconClosedBroken=" & IconClosedBrokenDefault _
        & "&" & "IconOpenBroken=" & IconOpenBrokenDefault _
        & "&" & "IconEmptyBroken=" & IconEmptyBrokenDefault _
        & "&" & "LinkToPageID=1" _
        & "&" & "ShowBrokenNodes=" & Main.GetVisitProperty("AllowOrphanRootPages", "0") _
        & "&" & "ShowAllChildPages=1" _
        & "&" & "ShowBlockedSections=1" _
        & "&" & "RemoteMethodName=Site Explorer Remote Method" _
        & ""
    GetContent = Main.GetAddonContent(0, "NavBase Remote Method", OptionString, ContextRemoteMethod, "", 0, "", InstanceID)
    'GetContent = NavBase.GetContent(OptionString)
    Set Main = Nothing
    Exit Function
ErrorTrap:
    Set Main = Nothing
    Call HandleError("RemoteMethodClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function

