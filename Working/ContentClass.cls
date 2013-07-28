VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContentClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
' ----- Global Objects
'
'Private Main As MainClass
Private Main As Object
Private Csv As Object
Private ColumnCnt As Long
'
' ----- Icons used
'
Const IconFolderOpen = "<img src=""/cclib/images/IconFolderOpen.gif"" border=""0"" width=""22"" height=""23"" ALT=""Close this folder"">"
Const IconFolderClosed = "<img src=""/cclib/images/IconFolderClosed.gif"" border=""0"" width=""22"" height=""23"" alt=""Open this folder"">"
Const IconFolderAdd = "<img src=""/cclib/images/IconFolderAdd2.gif"" border=""0"" width=""22"" height=""23"" alt=""Add a new folder"">"
Const IconFolderEdit = "<img src=""/cclib/images/IconFolderEdit.gif"" border=""0"" width=""22"" height=""23"" alt=""Edit this folder"">"
Const IconFile = "<img src=""/cclib/images/IconFile.gif"" border=""0"" width=""22"" height=""23"" alt=""file"">"
Const IconFileAdd = "<img src=""/cclib/images/IconContentAdd.gif"" border=""0"" width=""18"" height=""22"" alt=""Add a new  file"">"
Const IconFileEdit = "<img src=""/cclib/images/IconContentEdit.gif"" border=""0"" width=""18"" height=""22"" alt=""Edit this file"">"
Const IconPreview = "<img src=""/cclib/images/IconPreview.gif"" border=""0"" width=""22"" height=""23"" alt=""Preview this image"">"
Const IconDownload = "<img src=""/cclib/images/IconDownload3.gif"" border=""0"" width=""16"" height=""16"" alt=""Select this download"" valign=""absmiddle"">"
Const IconCreateImage = "<img src=""/cclib/images/IconimagePlace.gif"" border=""0"" width=""18"" height=""22"" alt=""Select this image"">"
Const IconCreateDownload = "<img src=""/cclib/images/IconDownload3.gif"" border=""0"" width=""16"" height=""16"" alt=""Select this download"" valign=""absmiddle"">"
Const IconSpacer = "<img src=""/cclib/images/spacer.gif"" width=""22"" height=""23"">"
Const IconView = "<img src=""/cclib/images/IconView.gif"" border=""0"" width=""22"" height=""23"" alt=""Preview this file"">"
Const IconImage = "<img src=""/cclib/images/IconImage2.gif"" border=""0"" width=""22"" height=""23"" alt=""Image"">"
Const IconPDF = "<img src=""/cclib/images/IconPDF.gif"" border=""0"" width=""16"" height=""16"" alt=""Adobe Pdf"">"
Const IconOther = "<img src=""/cclib/images/IconFile.gif"" border=""0"" width=""22"" height=""23"" alt=""Adobe Pdf"">"
Const IconNoFile = "<img src=/cclib/images/BulletRound2.gif width=5 height=5>"
'
' -----------------------------------------------------------------------------------
' ----- Privates
' -----------------------------------------------------------------------------------
'
Const iMinRows = 1
'Private iMinRows As Long
Private iFolderID  As Long                      ' Current Folder being Displayed, 0 for root
Private SourceMode As Long                      '
Private AllowSelectResource As Boolean
Private LinkObjectName As String
        '
        ' SourceMode
        '   0 = From Editor Object selector: allow image and download insert, provide close button
        '   1 = From Editor Image Properties: allow image insert, provide close button
        '   2 = From Admin site, no inserts, and provide cancel button
Const SourceModeDialog = 0
Const SourceModeSubDialog = 1
Const SourceModeOnPage = 2
        '
        '   0 caller is the editor directly, clicking on action icons calls InsertImaage, etc
        '   1 caller is the editor image page, clicking on action icons calls the image page methods
        '
Private HoldPosition As Long
'
' Page Selection
'
Private PageSelect As String
Private Type PageType
    PageID As Long
    ParentPageID As Long
    Name As String
    FullPath As String
    AccessSet As Boolean                ' true when HasAccess is correct
    HasAccess As Boolean                ' the permissions for this user
End Type
Private Pages() As PageType
Private PageCnt As Long
Private PageIndex As FastIndex4.FastIndexClass
Private PagePathIndex As New FastIndex4.FastIndexClass
'
' Exit Modes - how this add-on displays it's exit button
'
Private ExitMode As ExitModeEnum
Enum ExitModeEnum
    NoExit = 1
    CancelReturnNothing = 2
    CancelRedirectToPage = 3
    CloseWindow = 4
End Enum
'
'========================================================================
'   v3.3 Add-on Compatibility
'       To make an Add-on that works the same in v3.3 and v3.4, use this adapter instead of the execute above
'========================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set Csv = CsvObject
    Call Init(MainObject)
    Execute = GetContent(OptionString)
End Function
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
    Dim RootPageName As String
    Dim AllowGroupAdd As Boolean
    Dim BuildVersion As String
    '
    ' Determine exit mode
    '
    ExitMode = Main.GetStreamInteger("exitmode")
    If ExitMode = 0 Then
        If InStr(1, Main.ServerLink, Main.SiteProperty_AdminURL, vbTextCompare) <> 0 Then
            ExitMode = CancelReturnNothing
        End If
    End If
    RootPageName = Main.GetAddonOption("RootPageName", OptionString)
    LinkObjectName = Main.GetStreamText("LinkObjectName")
    BuildVersion = Main.SiteProperty_BuildVersion
    '
    GetContent = GetForm(RootPageName, AllowGroupAdd, BuildVersion)
    '
    If BuildVersion < "3.4.042" Then
        GetContent = GetContent & Main.GetAddonContent("", "NavBase", "", ContextRemoteMethod, "", 0, "", 0)
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetForm")
    End Function
'
'=================================================================================
' Returns the Resource Library HTML.
'   This HTML does not include the HTML, HEAD or BODY tags.
'=================================================================================
'
Private Function GetForm(RootPageName As String, AllowGroupAdd As Boolean, BuildVersion As String) As String
    On Error GoTo ErrorTrap
    '
    Dim AllowOrphanRootPages As Boolean
    Dim LineNumber As Long
    Dim LinePosition As Long
    Dim Copy As String
    Dim SectionID As Long
    Dim SectionName As String
    Dim SQL As String
    Dim QS As String
    Dim PageParentID As Long
    Dim CS As Long
    Dim RowClass As String
    Dim RQS As String
    Dim LoopPtr As Long
    Dim IsRootPage As Boolean
    Dim Lines() As String
    Dim Line As String
    Dim PageType As String
    Dim PageName As String
    Dim PageMenuHeadline As String
    Dim PageViewings As String
    Dim ChildPageCaption As String
    Dim PCCParentPtr As Long
    Dim ErrorMessage As String
    Dim ParentID As Long
    Dim ContentControlID As String
    Dim RowCnt As Long
    Dim Caption As String
    Dim Description As String
    Dim ButtonList As String
    'Dim AdminUI As AdminUIClass
    Dim AdminUI As Object
    Dim PCCPtr As Long
    Dim PtrCnt As Long
    Dim PCCPtrs() As Long
    Dim PCCPtrsSorted As Variant
    Dim OptionString As String
    Dim JumpSelect As String
    Dim OpenNodeList As String
    Dim RootPageID As Long
    Dim PageID As Long
    Dim FormFolders As String
    Dim FormDetails As New FastStringClass
    Dim RowCount As Long
    Dim IconLink As String
    Dim IconOnClick  As String
    Dim EditLink As String
    Dim ModifiedDate As Date
    Dim RecordName As String
    Dim recordId As Long
    Dim AllowFileAuthoring As Boolean
    Dim RecordCID As Long
    Dim AllowLocalFileAdd As Boolean
    Dim ButtonBar As String
    Dim Button As String
    Dim ButtonBarStyle As String
    Dim OptionPanelStyle As String
    Dim Link As String
    Dim Ptr As Long
    Dim ButtonExit
    Dim PCCLocal As Variant
    Dim PCCParentIDIndex As FastIndex4.FastIndexClass
    Dim SSCLocal As Variant
    Dim SSCCnt As Long
    Dim SSCPtr As Long
    '
    Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
    '
    Const Image5 = "<img src=/cclib/images/spacer.gif width=5 height=1>"
    Const Image10 = "<img src=/cclib/images/spacer.gif width=10 height=1>"
    Const Image15 = "<img src=/cclib/images/spacer.gif width=15 height=1>"
    Const Image20 = "<img src=/cclib/images/spacer.gif width=20 height=1>"
    Const Image30 = "<img src=/cclib/images/spacer.gif width=30 height=1>"
    Const Image50 = "<img src=/cclib/images/spacer.gif width=50 height=1>"
    '
    ButtonBarStyle = "" _
        & " color: black;" _
        & " font-weight: bold;" _
        & " padding: 5px;" _
        & " background-color: #a0a0a0;" _
        & " border-bottom: 1px solid #e0e0e0;" _
        & " border-right: 1px solid #e0e0e0;" _
        & " border-top: 1px solid #808080;" _
        & " border-left: 1px solid #808080;"
    '
    OptionPanelStyle = "" _
        & " color: black;" _
        & " font-weight: bold;" _
        & " padding: 5px;" _
        & " background-color: #d0d0d0;" _
        & " border-bottom: 1px solid #e0e0e0;" _
        & " border-right: 1px solid #e0e0e0;" _
        & " border-top: 1px solid #a0a0a0;" _
        & " border-left: 1px solid #a0a0a0;"
    '
    If Not (Main Is Nothing) Then
        '
        ' Determine Current Folder
        '
        PCCLocal = Main.GetPCC(False, False)
        AllowFileAuthoring = Main.IsContentManager("Page Content")
        Button = Main.GetStreamText("Button")
        PageID = Main.GetStreamInteger("PageID")
        If PageID <> 0 Then
            '
            ' Open node as well as display details
            '
            OpenNodeList = Main.GetVisitProperty("SiteExplorerOpenNodeList")
            OpenNodeList = OpenNodeList & "," & CStr(PageID)
            Call Main.SetVisitProperty("SiteExplorerOpenNodeList", OpenNodeList)
        End If
        '
        ' ----- Setup Local File Management
        '       Allow if Content Manager or user has group membership
        '       Always allow, everyone has access to the root folder, then if you can get to the folder, let em upload
        '
        'Call Main.AddRefreshQueryString("PageID", PageID)
        RQS = Main.RefreshQueryString
        AllowLocalFileAdd = True
        '
        ' ----- Process input
        '
        If Button = "" Then
            '
            ' no button, setup defaults
            '
            AllowOrphanRootPages = Main.GetVisitProperty("AllowOrphanRootPages", "0")
        Else
            Select Case Button
                Case ButtonApply
                    '
                    ' Apply Button ( for < 3.4.020 )
                    '
                    AllowOrphanRootPages = Main.GetStreamBoolean("AllowOrphanRootPages")
                    Call Main.SetVisitProperty("AllowOrphanRootPages", AllowOrphanRootPages)
                Case ButtonDelete
                    '
                    ' Delete Button
                    '
                    CS = Main.OpenCSContentRecord("page content", PageID)
                    If Main.IsCSOK(CS) Then
                        PageParentID = Main.GetCSInteger(CS, "ParentID")
                    End If
                    Call Main.CloseCS(CS)
                    Call Main.DeleteChildRecords("ccpagecontent", PageID, False)
                    Call Main.DeleteContentRecord("page content", PageID)
                    PageID = PageParentID
                    If PageID = 0 Then
                        PageID = Main.GetLandingPageID()
                    End If
                Case ButtonCancel
                    '
                    ' Cancel button
                    '
                    If ExitMode = CancelReturnNothing Then
                        GetForm = ""
                        Exit Function
                    ElseIf ExitMode = CancelRedirectToPage Then
                        Link = Main.ServerPage
                        Call Main.Redirect(Link)
                        Exit Function
                    End If
            End Select
        End If
        '
        ' ----- Guild Buttonbar
        '
        If (ExitMode = CancelRedirectToPage) Or (ExitMode = CancelReturnNothing) Then
            ButtonList = ButtonList & "," & ButtonCancel
            ButtonExit = Main.GetFormButton(ButtonCancel)
        ElseIf (ExitMode = CloseWindow) Then
            ButtonList = ButtonList & "," & ButtonClose
            ButtonExit = Main.GetFormButton(ButtonClose, , , "window.close();")
        End If
        ButtonList = ButtonList & "," & ButtonDelete
        ButtonList = ButtonList & "," & ButtonApply
        ButtonBar = "<div class=ccAdminButtonBar>" _
            & ButtonExit & AdminUI.GetButtonsFromList(Main, ButtonList, True, True, "Button") _
            & "</div>"
        JumpSelect = GetJumpSelect(PageID, RootPageName, PCCLocal)
        If JumpSelect = "" Then
            GetForm = GetForm _
                & "<TABLE class=ccAdminListCaption border=""0"" cellpadding=""5"" cellspacing=""5"" width=""100%"">" _
                & "<TR>" _
                & "<TD width=50% align=left>" & GetParentPagesLink(PCCLocal, RootPageName, RootPageID, PageID, PageID, Main.RefreshQueryString, "") & "</TD>" _
                & "<TD width=50% align=right>&nbsp;</TD>" _
                & "</TR>" _
                & "</TABLE>" _
                & "" _
                & ""
        Else
            GetForm = GetForm _
                & "<TABLE class=ccAdminListCaption border=""0"" cellpadding=""5"" cellspacing=""5"" width=""100%"">" _
                & "<TR><TD align=left>" & GetParentPagesLink(PCCLocal, RootPageName, RootPageID, PageID, PageID, Main.RefreshQueryString, "") & "</TD></TR>" _
                & "<TR><TD align=left style=""white-space:nowrap;"">" & "Jump&nbsp;to&nbsp;" & JumpSelect & "</TD></TR>" _
                & "</TABLE>" _
                & "" _
                & ""
        End If
        '
        ' From here down the form divides into FormFolder and FormDetails
        '
        FormDetails.Add "" _
            & "<style>" _
            & " .Header {text-align:left;padding:4px;border-bottom: 1px solid #C0C0C0;padding-top:5px}" _
            & " .ErrorCell {text-align:left;height:22px;padding:2px;border-bottom: 1px solid #C0C0C0;}" _
            & " .ErrorMsg {margin-left:20px;}" _
            & " .LeftCell {text-align:right;height:22px;padding:2px;width:140px;}" _
            & " .RightCell {text-align:left;height:22px;padding:2px;width:99%;}" _
            & " .EditCell {text-align:left;height:22px;padding:2px;border-bottom: 1px solid #C0C0C0}" _
            & "</style>"
        If PageID = 0 Then
            '
            ' Landing Page
            '
            Ptr = -1
            EditLink = ""
            PageName = "Landing Page unknown"
            PageMenuHeadline = "&nbsp;"
            PageViewings = "0"
            PageType = "Landing Page"
            ChildPageCaption = "Root Pages"
            If BuildVersion < "3.3.950" Then
                PageID = Main.GetSiteProperty("LandingPageID", "", True)
            Else
                PageID = Main.GetLandingPageID
            End If
            If PageID > 0 Then
                Ptr = Main.GetPCCPtr(PageID, False, False)
            End If
            If Ptr >= 0 Then
                PageName = kmaEncodeHTML(kmaEncodeText(PCCLocal(PCC_Name, Ptr)))
                PageMenuHeadline = kmaEncodeText(PCCLocal(PCC_MenuHeadline, Ptr))
                PageViewings = kmaEncodeText(PCCLocal(PCC_Viewings, Ptr))
                EditLink = GetRecordEditLink("Page Content", PageID, False, PageName, True, BuildVersion)
            End If
        Else
            SSCLocal = Main.GetSSC()
            SSCCnt = UBound(SSCLocal, 2) + 1
            Ptr = Main.GetPCCPtr(PageID, False, False)
            If Ptr < 0 Then
                '
                ' Page Not Found
                '
                ParentID = 0
                EditLink = ""
                PageName = ""
                PageMenuHeadline = ""
                PageViewings = ""
                ChildPageCaption = "Child Pages"
                PageType = "Page Not Found"
                FormDetails.Add "" _
                    & "<TABLE border=""0"" cellpadding=""4"" cellspacing=""0"" width=""100%"">" _
                    & vbCrLf & "<TR><TD width=200>Page Not Found<img src=""/cclib/images/spacer.gif"" width=132 height=1></TD><TD width=""99%""><img src=""/cclib/images/spacer.gif"" width=400 height=1></TD></TR>" _
                    & vbCrLf & "</Table>"
                FormDetails.Add "<div class=""ccAdminListCaption"" style=""clear:both;"">Child Pages</div>"
            Else
                '
                ' Page Details
                '
                ParentID = kmaEncodeInteger(PCCLocal(PCC_ParentID, Ptr))
                recordId = kmaEncodeInteger(PCCLocal(PCC_ID, Ptr))
                PageName = kmaEncodeHTML(kmaEncodeText(PCCLocal(PCC_Name, Ptr)))
                ContentControlID = kmaEncodeInteger(PCCLocal(PCC_ContentControlID, Ptr))
                If ParentID <> 0 Then
                    '
                    ' Child Page
                    '
                    PCCParentPtr = Main.GetPCCPtr(ParentID, False, False)
                    If PCCParentPtr < 0 Then
                        QS = RQS
                        QS = ModifyQueryString(QS, "pageid", kmaEncodeText(PageID), True)
                        PageType = "Child Page, but parent page is inactive or has been deleted"
                        ErrorMessage = ErrorMessage & vbCrLf & "The parent set for this page is missing. This page may not display on the site."
                        ErrorMessage = ErrorMessage & "<div class=""ErrorMsg""><a href=""" & Main.SiteProperty_AdminURL & "?" & QS & "&a904o2pa0cut=" & ContentControlID & "." & PageID & """>Fix</a> by cutting this page, then pasting it under a valid page.</div>"
                    Else
                        PageType = "Child Page"
                    End If
                    PageName = kmaEncodeHTML(kmaEncodeText(PCCLocal(PCC_Name, Ptr)))
                    PageMenuHeadline = kmaEncodeText(PCCLocal(PCC_MenuHeadline, Ptr))
                    PageViewings = kmaEncodeText(PCCLocal(PCC_Viewings, Ptr))
                    EditLink = GetRecordEditLink("Page Content", recordId, True, PageName, True, BuildVersion)
                Else
                    '
                    ' Root Page
                    '
                    PageName = kmaEncodeHTML(kmaEncodeText(PCCLocal(PCC_Name, Ptr)))
                    PageMenuHeadline = kmaEncodeText(PCCLocal(PCC_MenuHeadline, Ptr))
                    PageViewings = kmaEncodeText(PCCLocal(PCC_Viewings, Ptr))
                    If (SSCCnt > 0) Then
                        For SSCPtr = 0 To SSCCnt - 1
                            If recordId = kmaEncodeInteger(SSCLocal(SSC_RootPageID, SSCPtr)) Then
                                Exit For
                            End If
                        Next
                        If SSCPtr = SSCCnt Then
                            '
                            ' Missing section
                            '
                            PageType = "Root Page not assigned to a section"
                            QS = RQS
                            QS = ModifyQueryString(QS, "pageid", kmaEncodeText(PageID), True)
                            ErrorMessage = ErrorMessage & vbCrLf & "This page has no parent, but it is not associated directly to a section. It can not be displayed on a site using editable templates."
                            ErrorMessage = ErrorMessage & "<div style=""margin-left:10px""><a href=""" & Main.SiteProperty_AdminURL & "?af=4&id=0&cid=" & Main.GetContentID("Site Sections") & "&wc=RootPageID%3d" & PageID & ",name%3d" & PageName & ",caption%3d" & PageName & """>Fix</a> by creating a new section for this page.</div>"
                            ErrorMessage = ErrorMessage & "<div style=""margin-left:10px""><a href=""" & Main.SiteProperty_AdminURL & "?" & QS & "&a904o2pa0cut=" & kmaEncodeText(PCCLocal(PCC_ContentControlID, PCCPtr)) & "." & PageID & """>Fix</a> by cutting this page, then pasting it under a valid page.</div>"
                        Else
                            PageType = "Root Page"
                        End If
                    End If
                    EditLink = GetRecordEditLink("Page Content", recordId, False, PageName, True, BuildVersion)
                End If
                ChildPageCaption = "Child Pages"
            End If
        End If
        FormDetails.Add "" _
            & "<TABLE border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" _
            & vbCrLf & "<TR><TD><img src=""/cclib/images/spacer.gif"" width=132 height=1></TD><TD width=""99%""><img src=""/cclib/images/spacer.gif"" width=400 height=1></TD></TR>" _
            & vbCrLf & "<TR><TD class=""Header"" colspan=2><B>Page Details</B></TD></TR>" _
            & vbCrLf & "<TR><TD class=""LeftCell"">Name:&nbsp;</TD><TD class=""RightCell"">" & PageName & "</TD></TR>" _
            & vbCrLf & "<TR><TD class=""LeftCell"">Type:&nbsp;</TD><TD class=""RightCell"">" & PageType & "</TD></TR>" _
            & vbCrLf & "<TR><TD class=""LeftCell"">Menu Headline:&nbsp;</TD><TD  class=""RightCell"">" & PageMenuHeadline & "</TD></TR>" _
            & vbCrLf & "<TR><TD class=""LeftCell"">Viewings:&nbsp;</TD><TD  class=""RightCell"">" & PageViewings & "</TD></TR>"
            '& vbCrLf & "<TR><TD class=""LeftCell"">Edit:&nbsp;</TD><TD  class=""RightCell"">" & EditLink & "</TD></TR>"
        '
        ' Errors
        '
        If ErrorMessage <> "" Then
            FormDetails.Add vbCrLf & "<TR><TD class=""Header"" colspan=2 style=""color:red;""><B>Errors</B></TD></TR>"
            Lines = Split(ErrorMessage, vbCrLf)
            For Ptr = 0 To UBound(Lines)
                If Lines(Ptr) <> "" Then
                    FormDetails.Add vbCrLf & "<TR><TD class=""ErrorCell"" colspan=2>" & Lines(Ptr) & "</TD></TR>"
                End If
            Next
        End If
        '
        ' Spider Errors - bad links
        '
Dim UsedMsgList As String

        If Main.IsSQLTableField("default", "ccSpiderDocs", "PageID") Then
            SQL = "select e.* from ccSpiderErrors e left join ccSpiderDocs d on d.id=e.docid where d.pageid=" & PageID
            CS = Main.OpenCSSQL("default", SQL)
            If Main.IsCSOK(CS) Then
                FormDetails.Add vbCrLf & "<TR><TD class=""Header"" colspan=2 style=""color:red;""><B>Issues found by Content Spider</B></TD></TR>"
                Do While Main.IsCSOK(CS)
                    If Main.GetCSBoolean(CS, "Iswarning") Then
                        Copy = "Warning:"
                    Else
                        Copy = "Error:"
                    End If
                    LineNumber = Main.GetCSInteger(CS, "linenumber")
                    LinePosition = Main.GetCSInteger(CS, "LinePosition")
                    If LineNumber <> 0 Then
                        Copy = Copy & " line number " & LineNumber
                        If LinePosition <> 0 Then
                            Copy = Copy & " line position " & LinePosition
                        End If
                    End If
                    Copy = Copy & " " & Main.GetCSText(CS, "name")
                    If InStr(1, "," & UsedMsgList & ",", "," & Copy & ",", vbTextCompare) = 0 Then
                        UsedMsgList = UsedMsgList & "," & Copy
                        FormDetails.Add vbCrLf & "<TR><TD class=""ErrorCell"" colspan=2>" & Copy & "</TD></TR>"
                    End If
                    Call Main.NextCSRecord(CS)
                Loop
            End If
            Call Main.CloseCS(CS)
        End If
        '
        ' Edit Links
        '
        SectionID = 0
        FormDetails.Add vbCrLf & "<TR><TD class=""Header"" colspan=2><B>Edit Links</B></TD></TR>"
        FormDetails.Add vbCrLf & "<TR><TD class=""EditCell"" colspan=2>" & EditLink & "<div style=""margin-left:40px;"">edit this page [" & PageName & "]</div></TD></TR>"
        Do While ParentID <> 0 And LoopPtr < 20
            recordId = ParentID
            Ptr = Main.GetPCCPtr(recordId, False, False)
            If Ptr >= 0 Then
                ParentID = kmaEncodeInteger(PCCLocal(PCC_ParentID, Ptr))
            Else
                ParentID = 0
            End If
            RecordName = Main.GetRecordName("Page Content", recordId)
            IsRootPage = (ParentID = 0)
            If IsRootPage Then
                EditLink = GetRecordEditLink("Page Content", recordId, False, RecordName, True, BuildVersion)
                FormDetails.Add vbCrLf & "<TR><TD class=""EditCell"" colspan=2>" & EditLink & "<div style=""margin-left:40px;"">edit root page [" & RecordName & "]</div></TD></TR>"
            Else
                EditLink = GetRecordEditLink("Page Content", recordId, True, RecordName, True, BuildVersion)
                FormDetails.Add vbCrLf & "<TR><TD class=""EditCell"" colspan=2>" & EditLink & "<div style=""margin-left:40px;"">edit parent page [" & RecordName & "]</div></TD></TR>"
            End If
            
            LoopPtr = LoopPtr + 1
        Loop
        '
        ' Section Edit Link
        '
        CS = Main.OpenCSContent("Site Sections", "RootPageID=" & recordId)
        If Main.IsCSOK(CS) Then
            SectionID = Main.GetCSInteger(CS, "ID")
            SectionName = Main.GetCSText(CS, "name")
        End If
        Call Main.CloseCS(CS)
        If SectionID <> 0 Then
            EditLink = GetRecordEditLink("Site Sections", SectionID, False, SectionName, True, BuildVersion)
            FormDetails.Add vbCrLf & "<TR><TD class=""EditCell"" colspan=2>" & EditLink & "<div style=""margin-left:40px;"">edit section [" & SectionName & "]</div></TD></TR>"
        End If
        '
        ' Child Pages
        '
        FormDetails.Add vbCrLf & "<TR><TD class=""Header"" colspan=2><B>" & ChildPageCaption & "</B></TD></TR>"
        FormDetails.Add "</Table>"
        'FormDetails.Add "<div class=""ccAdminListCaption"" style=""clear:both;"">" & ChildPageCaption & "</div>"
        ColumnCnt = 7
        FormDetails.Add "" _
            & "<TABLE border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" _
            & vbCrLf & "<TR>" _
            & vbCrLf & vbTab & GetForm_HeaderCell("center", "15", "Edit<br>" & Image15) _
            & vbCrLf & vbTab & GetForm_HeaderCell("center", "15", "Place<br>" & Image15) _
            & vbCrLf & vbTab & GetForm_HeaderCell("center", "15", "Public<br>" & Image15) _
            & vbCrLf & vbTab & GetForm_HeaderCell("center", "20", "ID<br>" & Image20) _
            & vbCrLf & vbTab & GetForm_HeaderCell("left", "50%", "Name<br>" & Image20) _
            & vbCrLf & vbTab & GetForm_HeaderCell("left", "50%", "Menu&nbsp;Headline<br>" & Image20) _
            & vbCrLf & vbTab & GetForm_HeaderCell("center", "50", "Modified&nbsp;&nbsp;<br>" & Image50) _
            & vbCrLf & "</TR>"
            '& vbCrLf & vbTab & GetForm_HeaderCell("center", "10", "Select<BR>" & Image10) _

        '
        ' ----- Output the page
        '
        RowCount = 0
        If Not True Then
            '
            ' ----- No pages
            '
            FormDetails.Add "<TD>" & IconSpacer & "</TD><TD COLSPAN=" & ColumnCnt - 1 & ">You do not have access to this Page</TD>"
            RowCount = RowCount + 1
        Else
            '
            ' Lookup the files in the folder
            '
            If (PageType = "Landing Page") Then
                PCCPtr = Main.GetPCCFirstChildPtr(0, False, False)
            Else
                PCCPtr = Main.GetPCCFirstChildPtr(PageID, False, False)
            End If
            Set PCCParentIDIndex = Main.PCCParentIDIndex
            PtrCnt = 0
            Do While PCCPtr >= 0
                ReDim Preserve PCCPtrs(PtrCnt)
                PCCPtrs(PtrCnt) = PCCPtr
                PtrCnt = PtrCnt + 1
                PCCPtr = PCCParentIDIndex.GetNextPointerMatch(CStr(PageID))
            Loop
            If PtrCnt > 0 Then
                PCCPtrsSorted = Main.GetPCCPtrsSorted(PCCPtrs, "name")
            End If
            '
            ' List out the FILES
            '
            For Ptr = 0 To PtrCnt - 1
                PCCPtr = PCCPtrsSorted(Ptr)
                FormDetails.Add GetFormRow_Files(PCCLocal, PCCPtr, RowCount, AllowFileAuthoring)
                RowCount = RowCount + 1
            Next
            '
            ' ----- If nothing found, print no files found
            '
            If RowCount = 0 Then
                FormDetails.Add "<TD>" & IconSpacer & "</TD><TD COLSPAN=" & ColumnCnt - 1 & ">no child pages were found</TD>"
                RowCount = RowCount + 1
            End If
        End If
        '
        ' Add Child Page Row
        '
        RowClass = "ccPanelRowEven"
        If (RowCount Mod 2) = 0 Then
            RowClass = "ccPanelRowOdd"
        End If
        EditLink = GetRecordAddLink("Page Content", "ParentID=" & PageID, True, True, BuildVersion)
        FormDetails.Add vbCrLf & "<TR ID=Row" & RowCount & " Style=""Padding: 2px 2px 2px 2px"" align=""left"">"
        'FormDetails.Add vbCrLf & vbTab & "<TD Style=""border-top: 1px solid #C0C0C0; Padding: 2px 8px 2px 8px"" align=""left"" valign=""middle"">&nbsp;</TD>"
        FormDetails.Add vbCrLf & vbTab & "<TD Style=""border-top: 1px solid #C0C0C0; Padding: 2px 8px 2px 8px"" align=""left"" valign=""middle"" colspan=7>" & EditLink & "</TD>"
        FormDetails.Add vbCrLf & "</TR>"
        RowCount = RowCount + 1
        '
        ' Fill out the table to MinRows
        '
        For RowCount = RowCount To iMinRows
            FormDetails.Add GetFormRow_Blank(RowCount)
        Next
        '
        ' Bottom border
        '
        FormDetails.Add "</TABLE>"
        '
        ' Create the FormFolders
        '
        Dim Nav As New RemoteMethodClass
        Call Nav.Init(Main)
        OptionString = "ShowBrokenNodes=" & AllowOrphanRootPages
        FormFolders = Nav.GetContent(OptionString)
        FormFolders = "<div class=""NavCon"">" & FormFolders & "</div>"
        '
        ' Assemble the form
        '
        GetForm = GetForm _
            & "<TABLE border=0 cellpadding=0 cellspacing=0 width=""100%""><TR>" _
            & "<TD style=""background-color:white;"" class=ccPanel3DInput width=140 valign=top align=left>" & FormFolders & "<BR><img src=/cclib/images/spacer.gif width=140 height=1></TD>" _
            & "<TD style=""background-color:white;"" class=ccPanel3D width=5><img src=/cclib/images/spacer.gif width=5 height=1></TD>" _
            & "<TD style=""background-color:white;"" width=100% valign=top align=left>" & FormDetails.Text & "</TD>" _
            & "</TR>" _
            & "<tr><td colspan=3 style=""background-color:white;"" width=100% valign=top align=left></td></tr>" _
            & "</Table>" _
            & "<div class=""ccControlsCon"">" _
            & "<TABLE border=0 cellpadding=0 cellspacing=0 width=""100%""><TR>" _
            & "<TD class=""left"">" & Main.GetFormInputCheckBox("AllowOrphanRootPages", AllowOrphanRootPages) & "</TD>" _
            & "<TD class=""right"">Show Orphan Root Pages (pages with no parent page and no associated section)</TD>" _
            & "</TR></Table>" _
            & "</div>" _
            & "" _
            & Main.GetFormInputHidden("PageID", PageID) _
            & Main.GetFormInputHidden(RequestNameRefreshBlock, Main.GetFormSN) _
            & Main.GetFormInputHidden("RowCount", RowCount) _
            & ""
        Call Main.GetAddonContent(0, "NavBase", "RootPageID=-1", ContextPage, "", 0, "", 0)
        '
        Caption = SpanClassAdminNormal & "<strong>Site Explorer</strong></SPAN>"
        Description = ""
        GetForm = AdminUI.GetBody(Main, Caption, ButtonList, "", False, True, Description, "", 0, GetForm)
        Call Main.AddPageTitle("Site Explorer")
    End If
    '
    '
    ' Help Link
    '
    GetForm = Main.GetHelpLink(42, "Using the Resource Library", "The Resource Library is a convenient place to store reusable content, such as images and downloads. Objects in the Library can be placed on any page. The Library itself can be added to any page on your site.") & GetForm
    GetForm = "<div class=""SiteExplorerCon"">" & GetForm & "</div>"
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetForm")
    End Function
'
'=================================================================================
' Returns the Resource Library Row HTML.
'=================================================================================
'
Private Function GetFormRow_Files(ByRef PCCLocal As Variant, PCCPtr As Long, RowCount As Long, AllowFileAuthoring As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim EditLink As String
    Dim Copy As String
    Dim ImageLink As String
    Dim JSCopy As String
    Dim RowClass As String
    Dim AnchorTag As String
    Dim ImageTag As String
    Dim CellStart As String
    Dim CellStartRight As String
    Dim CellStart2 As String
    Dim CellStart5 As String
    Dim CellEnd As String
    Dim IconIMG As String
    Dim IconLinkStart As String
    Dim IconLinkEnd As String
    Dim DateString As String
    Dim InnerCell As String
    Dim PreviewImageURL As String
    Dim CellStartCenter As String
    Dim FileTypePtr As Long
    Dim IconFilename As String
    Dim IsImage As Boolean
    Dim IsVideo As Boolean
    Dim IsFlash As Boolean
    Dim IsMedia As Boolean
    Dim Mediafilename As String
    Dim IsDownload As Boolean
    Dim Downloadfilename As String
    Dim FileTypeName As String
    Dim TestFileTYpe As String
    Dim FileTypeFound As Boolean
    Dim MediaIMG As String
    Dim JSClose As String
    Dim PageName As String
    '
    Dim PageID As Long
    Dim ModifiedDate As Date
    Dim s As New FastStringClass
    '
    If (RowCount Mod 2) = 0 Then
        RowClass = "ccPanelRowOdd"
    Else
        RowClass = "ccPanelRowEven"
    End If
    '
    CellStart = vbCrLf & vbTab & "<TD Style=""border-top: 1px solid #C0C0C0; Padding: 2px 8px 2px 8px"" align=""left"" valign=""middle""><SPAN class=""ccAdminSmall"">"
    CellStartCenter = vbCrLf & vbTab & "<TD Style=""border-top: 1px solid #C0C0C0; Padding: 2px 8px 2px 8px"" align=""center"" valign=""middle""><SPAN class=""ccAdminSmall"">"
    CellStartRight = vbCrLf & vbTab & "<TD Style=""border-top: 1px solid #C0C0C0; Padding: 2px 8px 2px 8px"" align=""right"" valign=""middle""><SPAN class=""ccAdminSmall"">"
    CellEnd = "</SPAN></TD>"
    '
    PageID = kmaEncodeInteger(PCCLocal(PCC_ID, PCCPtr))
    PageName = kmaEncodeText(PCCLocal(PCC_Name, PCCPtr))
    If AllowFileAuthoring Then
        EditLink = Main.SiteProperty_AdminURL & "?cid=" & kmaEncodeInteger(PCCLocal(PCC_ContentControlID, PCCPtr)) & "&id=" & PageID & "&af=4" & "&aa=2&depth=1"
    Else
        EditLink = ""
    End If
    '
    '------------------------------------
    '   Output the rows
    '------------------------------------
    '
    s.Add vbCrLf & "<TR ID=Row" & RowCount & " Style=""Padding: 2px 2px 2px 2px"" align=""left"">"
'    '
'    ' Select
'    '
'    s.Add CellStartCenter & "<input type=checkbox name=DeleteRow" & RowCount & " value=1 >" & Main.GetFormInputHidden("DeleteRowID" & RowCount, PageID) & CellEnd
    '
    ' Edit
    '
    If EditLink <> "" Then
        s.Add CellStartCenter & "<A href=""" & EditLink & """>" & IconFileEdit & "</A>" & CellEnd
    Else
        s.Add CellStart & "&nbsp;" & CellEnd
    End If
    '
    ' Place
    '
    If LinkObjectName = "" Then
        '
        ' Can not select resources
        '
        s.Add CellStartCenter & IconNoFile & CellEnd
    Else
        '
        ' Allow selection of resources to be placed on the opening pages
        '
        JSCopy = kmaEncodeJavascript(Main.GetPageLink(PageID))
        ImageLink = "<img src=""/cclib/images/PageLink1616.gif"" border=""0"" width=""16"" height=""16"" alt=""Place a link to this page"" title=""Place a link to this page"" valign=""absmiddle"">"
        s.Add CellStartCenter & "<a href=""#"" onClick=""var e=window.opener.document.getElementById('" & LinkObjectName & "');e.value='" & JSCopy & "'; window.close();"">" & ImageLink & "</A>" & CellEnd
    End If
    '
    ' View
    '
Copy = Main.GetPageLink(PageID)
'Copy = ""
    If Copy = "" Then
        s.Add CellStartCenter & IconNoFile & CellEnd
    Else
        s.Add CellStartCenter & "<a href=""" & Copy & """>Go</a>" & CellEnd
    End If
    '
    ' ID
    '
    s.Add CellStartRight & PageID & CellEnd
    '
    ' Name
    '
    Copy = PageName
    If Copy = "" Then
        Copy = "&nbsp;"
    End If
    s.Add CellStart & Copy & CellEnd
    '
    ' Menu headline
    '
    Copy = kmaEncodeText(PCCLocal(PCC_MenuHeadline, PCCPtr))
    If Copy = "" Then
        Copy = "&nbsp;"
    End If
    s.Add CellStart & Copy & CellEnd
    '
    ' Modified
    '
    ModifiedDate = KmaEncodeDate(PCCLocal(PCC_ModifiedDate, PCCPtr))
    If ModifiedDate = CDate(0) Then
        DateString = "&nbsp;"
    Else
        DateString = FormatDateTime(ModifiedDate, vbShortDate)
    End If
    s.Add CellStartRight & DateString & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & CellEnd
    '
    s.Add "</TR>"
    '
    GetFormRow_Files = s.Text
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetFormRow_Files")
    End Function
''
''=================================================================================
'' Returns the Resource Library Row HTML.
''=================================================================================
''
'Private Function GetFormRow_AddChild(ignore0 As String, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Long, EditLink As String, Description As String, FileType As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Long, FolderID As Long) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim RowClass As String
'    Dim AnchorTag As String
'    Dim ImageTag As String
'    Dim CellStartCenter As String
'    Dim CellStart As String
'    Dim CellStart2 As String
'    Dim CellStart5 As String
'    Dim CellEnd As String
'    Dim Icon As String
'    Dim IconLinkStart As String
'    Dim IconLinkEnd As String
'    Dim DateString As String
'    Dim InnerCell As String
'    '
'    If Main.IsAdmin() Then
'        CellStart = "<TD Style=""border-top: 1px solid #C0C0C0; Padding: 2px 2px 2px 2px"" align=""left"" valign=""middle""><SPAN class=""ccAdminSmall"">"
'        CellStartCenter = "<TD Style=""border-top: 1px solid #C0C0C0; Padding: 2px 2px 2px 2px"" Class=""" & RowClass & """ align=""center"" valign=""middle""><SPAN class=""ccAdminSmall"">"
'        CellEnd = "</SPAN></TD>"
'        '
'        If ModifiedDate = CDate(0) Then
'            DateString = "&nbsp;"
'        Else
'            DateString = FormatDateTime(ModifiedDate, vbShortDate)
'            End If
'        If EditLink <> "" Then
'            GetFormRow_AddChild = GetFormRow_AddChild & "<TR Style=""Padding: 2px 2px 2px 2px"" align=""left"">"
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & "&nbsp;" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & "&nbsp;" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStart & "<b>Add Folder</b>" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & "<A href=""" & EditLink & """ target=""_blank"">" & IconFolderAdd & "</A>" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & IconNoFile & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & IconNoFile & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & "&nbsp;" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStartCenter & "&nbsp;" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddChild = GetFormRow_AddChild & "</TR>"
'        End If
'    End If
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleClassTrapError("GetFormRow_AddChild")
'    End Function
''
''=================================================================================
'' Returns the Resource Library Row HTML.
''=================================================================================
''
'Private Function GetFormRow_AddFile(ignore0 As String, IconLink As String, IconOnClick As String, Name As String, NameLink As String, ModifiedDate As Date, RowCount As Long, EditLink As String, Description As String, FileType As String, ResourceName As String, ResourceLink As String, ImageAlt As String, ImageWidth As String, ImageHeight As String, RecordID As Long, FolderID As Long) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim RowClass As String
'    Dim AnchorTag As String
'    Dim ImageTag As String
'    Dim CellStart As String
'    Dim CellStartCenter As String
'    Dim CellStart2 As String
'    Dim CellStart5 As String
'    Dim CellEnd As String
'    Dim Icon As String
'    Dim IconLinkStart As String
'    Dim IconLinkEnd As String
'    Dim DateString As String
'    Dim InnerCell As String
'    '
'    If Main.IsAdmin() Then
'        CellStart = "<TD Style=""border-top: 1px solid #C0C0C0; Padding: 2px 2px 2px 2px"" align=""left"" valign=""middle""><SPAN class=""ccAdminSmall"">"
'        CellStartCenter = "<TD Style=""border-top: 1px solid #C0C0C0; Padding: 2px 2px 2px 2px"" Class=""" & RowClass & """ align=""center"" valign=""middle""><SPAN class=""ccAdminSmall"">"
'        CellEnd = "</SPAN></TD>"
'        '
'        If ModifiedDate = CDate(0) Then
'            DateString = "&nbsp;"
'        Else
'            DateString = FormatDateTime(ModifiedDate, vbShortDate)
'        End If
'        If EditLink <> "" Then
'            GetFormRow_AddFile = GetFormRow_AddFile & "<TR Style=""Padding: 2px 2px 2px 2px"" align=""left"">"
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "<b>Add File</b>" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStartCenter & "<A href=""" & EditLink & """ target=""_blank"">" & IconFileAdd & "</A>" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStartCenter & IconNoFile & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStartCenter & IconNoFile & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & CellStart & "&nbsp;" & CellEnd
'            GetFormRow_AddFile = GetFormRow_AddFile & "</TR>"
'        End If
'    End If
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleClassTrapError("GetFormRow_AddFile")
'    End Function
'
'=================================================================================
' Returns the Resource Library Row HTML.
'=================================================================================
'
Private Function GetFormRow_Blank(RowCount As Long) As String
    On Error GoTo ErrorTrap
    '
    Const CellStart = "<TD Style=""border-top: 1px solid #C0C0C0; Padding: 2px 2px 2px 2px"" align=""left"" valign=""middle"">"
    Const CellEnd = "</SPAN></TD>"
    '
    Dim RowClass As String
    '
    If (RowCount Mod 2) = 0 Then
        RowClass = "ccPanelRowOdd"
    Else
        RowClass = "ccPanelRowEven"
    End If
    '
    GetFormRow_Blank = GetFormRow_Blank & vbCrLf & "<TR Style=""Padding: 2px 2px 2px 2px"" align=""left"">"
    GetFormRow_Blank = GetFormRow_Blank & vbCrLf & vbTab & CellStart & "&nbsp;" & CellEnd
    GetFormRow_Blank = GetFormRow_Blank & vbCrLf & vbTab & CellStart & IconSpacer & CellEnd
    GetFormRow_Blank = GetFormRow_Blank & vbCrLf & vbTab & CellStart & IconSpacer & CellEnd
    GetFormRow_Blank = GetFormRow_Blank & vbCrLf & vbTab & CellStart & IconSpacer & CellEnd
    GetFormRow_Blank = GetFormRow_Blank & vbCrLf & vbTab & CellStart & "&nbsp;" & CellEnd
    GetFormRow_Blank = GetFormRow_Blank & vbCrLf & vbTab & CellStart & "&nbsp;" & CellEnd
    GetFormRow_Blank = GetFormRow_Blank & vbCrLf & vbTab & CellStart & "&nbsp;" & CellEnd
    GetFormRow_Blank = GetFormRow_Blank & vbCrLf & vbTab & CellStart & "&nbsp;" & CellEnd
    GetFormRow_Blank = GetFormRow_Blank & vbCrLf & "</TR>"
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetFormRow_Blank")
    End Function
'
'
'
Private Sub HandleClassTrapError(MethodName As String, Optional ignore0 As String)
    Set Main = Nothing
    Call HandleError("SiteExplorer", MethodName, Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'
'
Private Function GetForm_HeaderCell(Align As String, Width As String, Copy As String) As String
    Dim Style
    '
    Style = "" _
        & "padding: 3px;" _
        & "font-size:10px;"
    GetForm_HeaderCell = "<TD WIDTH=""" & Width & """ ALIGN=""" & Align & """ class=ccAdminListCaption style=""" & Style & """>" _
        & Copy _
        & "</TD>"
        '& Main.GetPanel(kmaAddSpan(Copy, "ccAdminSmall") & "<BR /><IMG SRC=""/cclib/images/spacer.gif"" WIDTH=""1"" HEIGHT=""1"">")
    End Function
'
'
'
Private Function GetParentPagesLink(ByRef PCCLocal As Variant, RootPageName As String, RootPageID As Long, CurrentPageID As Long, PageID As Long, RefreshQS As String, ChildIDList As String) As String
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim CS As Long
    Dim ParentID As Long
    Dim PageName As String
    Dim RecordFound As Boolean
    Dim PageLink As String
    '
    If PageID = 0 Or (PageID = RootPageID) Then
        '
        ' Root Page
        '
        PageName = RootPageName
        If PageName = "" Then
            PageName = "Landing&nbsp;Page"
        End If
        If CurrentPageID = PageID Then
            GetParentPagesLink = "<B>" & PageName & "</B>"
            'GetParentPagesLink = "Page <B>" & PageName & "</B>"
        Else
            GetParentPagesLink = "<a href=?" & RefreshQS & "&PageID=0>" & PageName & "</a>"
            'GetParentPagesLink = "Page <a href=?" & RefreshQS & "&PageID=0>" & PageName & "</a>"
        End If
    Else
Dim PCCPtr As Long
        PCCPtr = Main.GetPCCPtr(PageID, False, False)
        If PCCPtr >= 0 Then
            RecordFound = True
            ParentID = kmaEncodeInteger(PCCLocal(PCC_ParentID, PCCPtr))
            PageName = kmaEncodeText(PCCLocal(PCC_Name, PCCPtr))
        End If
        '
        If CurrentPageID = PageID Then
            PageLink = "<B>" & PageName & "</B>"
        Else
            PageLink = "<a href=?" & RefreshQS & "&PageID=" & PageID & ">" & PageName & "</a>"
        End If
        '
        If (Not RecordFound) Or (PageID = RootPageID) Then
            '
            ' call this the top of the tree
            '
            If PageName = "" Then
                PageName = "Root"
            End If
            GetParentPagesLink = "Page " & PageLink
        ElseIf InStr(1, ChildIDList & ",", "," & PageID & ",") <> 0 Then
            '
            ' circular reference - end it here
            '
            GetParentPagesLink = "Page (Circular Reference) > " & PageLink
        ElseIf CurrentPageID = ParentID Then
            '
            ' circular reference - end it here
            '
            GetParentPagesLink = "Page " & PageLink
        Else
            GetParentPagesLink = GetParentPagesLink(PCCLocal, RootPageName, RootPageID, CurrentPageID, ParentID, RefreshQS, ChildIDList & "," & PageID) & " > " & PageLink
        End If
    End If
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetParentPageLinks")
End Function
'
'
'
Private Function GetJumpSelect(FolderID As Long, RootPageName As String, ByRef PCCLocal As Variant) As String
    On Error GoTo ErrorTrap
    '
    GetJumpSelect = GetPageSelect(FolderID, RootPageName, PCCLocal)
    If GetJumpSelect <> "" Then
        GetJumpSelect = Replace(GetJumpSelect, "FieldName", "JumpFolderID")
        GetJumpSelect = Replace(GetJumpSelect, "onChange", "onChange=""QJump(this);"" ")
        GetJumpSelect = Replace(GetJumpSelect, "value=", "value=?" & Main.RefreshQueryString & "&PageID=")
        GetJumpSelect = "<script language=JavaScript1.2>function QJump(e){var l=e.value;if(l!=''){window.name='RL';window.location.assign(l);}}</script>" & GetJumpSelect
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetParentFolderLinks")
End Function

'
'
'
Private Function GetMoveSelect(FolderID As Long, RootPageName As String, ByRef PCCLocal As Variant) As String
    On Error GoTo ErrorTrap
    '
    GetMoveSelect = GetPageSelect(FolderID, RootPageName, PCCLocal)
    GetMoveSelect = Replace(GetMoveSelect, "FieldName", "MoveFolderID")
    GetMoveSelect = Replace(GetMoveSelect, "onChange", "onChange=""var e=getElementById('Move');if(e){e.checked=true};"" ")
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetParentFolderLinks")
End Function
'
'=============================================================
'
'=============================================================
'
Private Function GetRLNav(CurrentFolderID As Long, RootPageName As String, RootFolderID As Long) As String
    '
    GetRLNav = "nav"
End Function
'
'=================================================================================
' Create
'=================================================================================
'
Private Sub Class_Initialize()
    'iMinRows = 10
    End Sub
'
'=================================================================================
' Kill
'=================================================================================
'
Private Sub Class_Terminate()
    Set Main = Nothing
    End Sub
'
'
'
Private Function GetPageSelect(TopPageID As Long, RootPageName As String, ByRef PCCLocal As Variant) As String
    On Error GoTo ErrorTrap
    '
    Dim s As New FastStringClass
    Dim SQL As String
    Dim CS As Long
    Dim Ptr As Long
    Dim PtrPageID As Long
    Dim Cnt As Long
    Dim BakeName As String
    Dim Pos As Long
    Dim FullPath As String
    Dim IndexCnt As Long
    Dim ParentPageID As Long
    Dim PageID As Long
    Dim PtrString As String
    '
    If RootPageName = "root" Then
        RootPageName = ""
    End If
    '
    GetPageSelect = PageSelect
    If GetPageSelect = "" Then
        BakeName = "SiteExplorerPageSelect-" & RootPageName & "-" & TopPageID
        If True Then
            '
            ' Collect all the Pages
            '
            PageCnt = Main.PCCCnt
            ReDim Pages(Main.PCCCnt)
            For Ptr = 0 To Main.PCCCnt - 1
                With Pages(Ptr)
                    .AccessSet = True
                    .FullPath = ""
                    .HasAccess = True
                    .Name = kmaEncodeHTML(kmaEncodeText(PCCLocal(PCC_Name, Ptr)))
                    .PageID = kmaEncodeInteger(PCCLocal(PCC_ID, Ptr))
                    .ParentPageID = kmaEncodeInteger(PCCLocal(PCC_ParentID, Ptr))
                End With
            Next
            '
            ' create full paths
            '
            If PageCnt > 0 Then
                For Ptr = 0 To PageCnt - 1
                    If Pages(Ptr).HasAccess Then
                        PtrPageID = Pages(Ptr).PageID
                        FullPath = Pages(Ptr).FullPath
                        If FullPath = "" Then
                            FullPath = GetPagePath(Ptr, "")
                            Pages(Ptr).FullPath = FullPath
                        End If
                        If RootPageName <> "" Then
                            Pos = InStr(1, FullPath, "> " & RootPageName, vbTextCompare)
                            If Pos = 0 Then
                                FullPath = ""
                            Else
                                FullPath = Mid(FullPath, Pos + 2)
                            End If
                        End If
                        If FullPath <> "" Then
                            IndexCnt = IndexCnt + 1
                            Pages(Ptr).FullPath = kmaEncodeHTML(FullPath)
                            Call PagePathIndex.SetPointer(FullPath, Ptr)
                        End If
                    End If
                Next
            End If
            '
            ' create select
            '
            If IndexCnt > 0 Then
                '
                If RootPageName = "" Then
                    If TopPageID = 0 Then
                        s.Add "<option value=0 selected>Landing Page</option>"
                    Else
                        s.Add "<option value=0>Landing Page</option>"
                    End If
                End If
                PtrString = PagePathIndex.GetFirstPointer
                Do While (PtrString <> "")
                    Ptr = kmaEncodeInteger(PtrString)
                    PtrPageID = Pages(Ptr).PageID
                    If Pages(Ptr).FullPath = "" Then
                        Pages(Ptr).FullPath = GetPagePath(Ptr, "")
                    End If
                    FullPath = Pages(Ptr).FullPath
                    If Len(FullPath) > 100 Then
                        
                        FullPath = Left(FullPath, 45) & "......" & Right(FullPath, 45)
                    End If
                    If PtrPageID = TopPageID Then
                        s.Add "<option value=" & PtrPageID & " selected>" & FullPath & "</option>"
                    Else
                        s.Add "<option value=" & PtrPageID & ">" & FullPath & "</option>"
                    End If
                    PtrString = PagePathIndex.GetNextPointer
                Loop
                GetPageSelect = s.Text
            End If
            '
            ' Create Select
            '
            If GetPageSelect <> "" Then
                GetPageSelect = "<select name=FieldName size=1 onChange>" & GetPageSelect & "</select>"
            End If
        End If
        PageSelect = GetPageSelect
    End If
    '
    Exit Function
ErrorTrap:
    HandleClassTrapError ("GetParentPageLinks")
End Function
'
'
'
Private Function GetPagePath(TargetPtr As Long, ChildIDList As String) As String
    On Error GoTo ErrorTrap
    '
    Dim ParentPtr As Long
    Dim ParentID As Long
    Dim PageID As Long
    
    '
    GetPagePath = Pages(TargetPtr).Name
    ParentID = Pages(TargetPtr).ParentPageID
    PageID = Pages(TargetPtr).PageID
    If ParentID = 0 Then
        '
        ' At the Root page
        '
        GetPagePath = GetPagePath
        'GetPagePath = "Root > " & GetPagePath
    ElseIf (PageID = ParentID) Or (InStr(1, "," & ChildIDList & ",", "," & ParentID & ",") <> 0) Then
        '
        ' circular reference - Make this a root page b
        '
    Else
        For ParentPtr = 0 To UBound(Pages)
            If Pages(ParentPtr).PageID = ParentID Then
                GetPagePath = GetPagePath(ParentPtr, ChildIDList & "," & ParentID) & " > " & GetPagePath
                Exit For
            End If
        Next
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetParentPageLinks")
End Function
'
'========================================================================
' DeleteChildRecords
'========================================================================
'
Private Function DeleteChildRecords(TableName As String, recordId As Long, ReturnList As Boolean, BuildVersion As String) As String
    On Error GoTo ErrorTrap
    '
    If BuildVersion > "3.3.950" Then
        Call Main.DeleteChildRecords("ccpagecontent", recordId, False)
    Else
        Call Main.AddUserError("To automatically delete childs page you need to upgrade Contensive and your database to version 3.3.950.")
    End If
    
    '
    Exit Function
    '
ErrorTrap:
    Call HandleClassTrapError("DeleteChildRecords")
    End Function
'
'
'
Private Function GetRecordEditLink(ContentName As String, recordId As Long, AllowCut As Boolean, PageName As String, YesNo As Boolean, BuildVersion As String)
    If BuildVersion > "3.3.950" Then
        GetRecordEditLink = Main.GetRecordEditLink2(ContentName, recordId, AllowCut, PageName, YesNo)
    Else
        GetRecordEditLink = "<a href=""" & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentName) & "&id=" & recordId & "&af=4""><img src=""/cclib/images/IconContentEdit.gif"" border=0></a>&nbsp;"
    End If

End Function
'
'
'
Private Function GetRecordAddLink(ContentName As String, NameValueString As String, AllowCut As Boolean, YesNo As Boolean, BuildVersion As String)
    If BuildVersion > "3.3.950" Then
        GetRecordAddLink = Main.GetRecordAddLink2(ContentName, NameValueString, AllowCut, YesNo)
    End If
End Function

