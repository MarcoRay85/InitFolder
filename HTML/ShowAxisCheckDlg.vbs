'********************************************************************************
'*                                                                              *
'*        S C R I P T   F I L E    F O R  S h o w H e a d C h e c k D l g         *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************
option explicit

    Dim objShellServices, objFSO, objWShell              
    Set objShellServices = CreateObject( "ScriptingToolsSO.ShellServices" )
    Set objFSO           = CreateObject( "Scripting.FileSystemObject"     )
    Set objWShell        = CreateObject( "WScript.Shell"                  )

'********************************************************************************


Const C_strIIDPrefix            =  "IID_"
Const C_strToolIDPrefix         =  "spanTool"
Const C_strToolBtnClass         =  "classToolButton"
Const C_strAttibToolPath        =  "Path"

Const C_AttrNameAppObj          =  "_AppObj_"
Const C_AttrShareCnt            =  "_SHARES_"
Const C_AttrCanceled            =  "_CANCEL_"

Const C_ClassError              =  "classERROR"
Const C_LangVBScript            =  "VBScript"

Const C_strStateProp            =  "State"

Const C_nDefaultSizeFactor      =  0.75

Const WM_CLOSE                  =  &H0010
Const WM_SYSCOMMAND             =  &H0112
Const SC_CLOSE                  =  61536    '0xF060


Dim g_nCurPage
Dim g_nProgressDlgRefCnt
Dim g_nOutputSizeFactor

Dim g_strCurPageName
Dim g_strStartupScript


Dim Application
Dim g_objFSO
Dim g_objSTSOFactory
Dim g_objVBXens
Dim g_objObjectInfo
Dim g_objWinHandle
Dim g_objLogHost
Dim g_objPLCList
Dim g_objEPCLogin

Dim g_objBeforeQuitHdl
Dim g_objHTMLProgressDlg


Set g_objFSO                =  CreateObject(  "Scripting.FileSystemObject"      )
Set g_objEPCLogin           =  CreateObject(  "EPCClient.EPCLogin"              )

Set g_objBeforeQuitHdl      =  Nothing
Set g_objHTMLProgressDlg    =  Nothing

Set g_objPLCList            =  Nothing

g_nProgressDlgRefCnt        =  0 
g_nOutputSizeFactor         =  C_nDefaultSizeFactor
    
Dim g_objDict

Dim g_objApplication
Dim g_objMachine
Dim g_objMachineScript
Dim g_objMachineState
Dim g_objVBSXtensions
Dim g_objHTMLHelper
Dim g_objWinTools

Dim g_nVideoTickerCookie
Dim g_nVideoTickerTime

Set g_objApplication            = Nothing
Set g_objMachine                = Nothing
Set g_objMachineScript          = Nothing
Set g_objMachineState           = Nothing
Set g_objHTMLHelper             = CreateObject( "IEHTMLToolsSO.HTMLHelper"      )
Set g_objVBSXtensions           = CreateObject( "ScriptingToolsSO.VBSXtensions" )
Set g_objWinTools               = CreateObject( "ScriptingToolsSO.Windowing"    )
    
'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************
'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim objStyle
    Dim objElem
    Dim vArgs
    Dim vTmp
    Dim nI
    Dim nLen

    With Document.Body

        With .WindowHandle

            .ModifyStyle  0, &H00080000
            .UpdateLook
            .InstallMessageFilter GetRef( "MsgHdl" ), WM_CLOSE, WM_CLOSE

        End With

    End With

    For Each objStyle In Document.StyleSheets       ' Match colors to different backgrounds in NT4 and W2K

        If objStyle.ID = "linkedstyleCustom"  Then

            nLen = objStyle.Rules.Length - 1

            For nI = 0 To nLen

                Set objElem = objStyle.Rules.Item( nI )

                If objElem.SelectorText = ".classTableHeader"  Then

                    vTmp = g_objHTMLHelper.AmpColor( "ThreeDFace", 0.85 )

                    objElem.Style.BackgroundColor = vTmp

                End If

                If objElem.SelectorText = ".classButtons"  Then

                    vTmp = buttonStopPos.OffsetHeight
                    buttonStopPos.Style.PixelWidth = vTmp

                    With buttonStopPos.Style
                    
                        .PixelWidth  = vTmp
                        .PixelHeight = vTmp

                    End With

                    With objElem.Style

                        .Width      = vTmp & "px"
                        .Height     = vTmp & "px"
                        .LineHeight = vTmp & "px"

                    End With

                End If

            Next

            Exit For

        End If

    Next


    If Not Document.Body.IsDialog Then

    Else
    
        With Document.ParentWindow

            vArgs  =  .DialogArguments

            Set g_objApplication    = vArgs( 0 )
            Set g_objMachineScript  = vArgs( 1 )
            .DialogTop              = vArgs( 2 )

        End With

        Set g_objMachine      = g_objApplication.Machine
        Set g_objMachineState = g_objMachine.State


        g_objWinTools.RepaintWindowTitle  g_objApplication.HWND, TRUE

   End If
   
   If Document.Body.IsDialog Then

        If IsObject( vArgs( 3 ))  Then

            Set vTmp = vArgs( 3 )

            If Not ( vTmp Is Nothing )  Then  vTmp  Me, GetRef( "Command" ), vArgs( 4 )

        End If

    End If
    
    buttonVideo.AttachEvent         "onclick",      GetRef(  "VideoButton_onclick"  )

    divFrameHostView.SizeToContent     "H"
    divFrameHostView.CenterContent     "X"

    divFrameHostViewRot.SizeToContent     "H"
    divFrameHostViewRot.CenterContent     "X"

    divFrameHostViewLift.SizeToContent     "H"
    divFrameHostViewLift.CenterContent     "X"

    divFrameHostViewTilt.SizeToContent     "H"
    divFrameHostViewTilt.CenterContent     "X"

    divFrameHostViewShift.SizeToContent     "H"
    divFrameHostViewShift.CenterContent     "X"


    divFrameHostLog.SizeToContent     "H"
    divFrameHostLog.CenterContent     "X"
  
    divToolPane.SizeToContent     "H"
    divToolPane.CenterContent     "X"
  
    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )
    STANDARDDLG_BUTTONS.WithoutAlt = False
    STANDARDDLG_BUTTONS.Default = 0

End Sub

'***********************************************************************
'****                                                               ****
'****       Handler for form initialisation                         ****
'****                                                               ****
'***********************************************************************

Sub window_onload()     'Initialisation

    Dim vTmp
    Dim nCurPage
    Dim nFirstPage
    Dim strPath
    Dim strName
    Dim strValue
    Dim objRoot
    Dim objElem

    strPath = Replace(  Left( document.url,                     _
                              InStrRev( document.url, "\" )),   _
                        "file://", "" )


    Set g_objVBXens     =  ScriptingToolsCreator(  "ScriptingToolsSO.VBSXtensions"   )
    Set g_objObjectInfo =  ScriptingToolsCreator(  "ScriptingToolsSO.ObjectInfo"     )
    Set g_objWinHandle  =  ScriptingToolsCreator(  "ScriptingToolsSO.WindowHandle"   )

    g_objWinHandle.ID   =  Document
    Set g_objWinHandle  =  g_objWinHandle.Parent

    With ScriptingToolsCreator(  "ScriptingToolsSO.ObjectFactory"  )

        Set Application   =  .Create( "Application" )
        Set g_objPLCList  =  .Create( "PLCList")

    End With
     
    With iframeLOG.Document

        For Each vTmp In Document.StyleSheets       ' Include style sheets...

            .CreateStyleSheet.CSSText = vTmp.CSSText

        Next

       .CreateStyleSheet.CSSText = " PRE { Padding:3px;Margin:0;Border:0;Font-Size:8pt;Line-Height:1em; }"

       .AttachEvent  "oncontextmenu", GetRef( "IFRAMEonContextMenu" )

        Set g_objLogHost = .Body

        With g_objLogHost.Style
    
            .Overflow   = "Scroll"
            .LineHeight = "1em"

        End With
  
    End With


                                        '***** Retrieving ScriptingToolsSO constants ****

    Set g_objObjectInfo.Instance = g_objObjectInfo
    Set objElem                  = g_objObjectInfo.GetConstantsOfTypeLib
    Set g_objObjectInfo.Instance = Nothing

    vTmp = Empty
    
    For Each strValue In objElem( -1 )

        vTmp =   vTmp & "Const " &  strValue &  " = " & objElem( strValue ) &  vbNewLine

    Next

    Set objElem = Nothing
    strName     = vTmp

if 0 then
    
    strValue = UCase( Window.ClientInformation.UserLanguage )

    nCurPage = Replace(  document.url, "file://", "" )
    vTmp     = InStr( 1, objThisApp.CommandLine, nCurPage, vbTextCompare )

    If  vTmp > 0  Then

        nFirstPage = Left( objThisApp.CommandLine, vTmp - 1 )               _
                   & Mid(  objThisApp.CommandLine, vTmp + Len( nCurPage ))

        If  Len( nFirstPage ) > 6  Then

            With  New RegExp

                .IgnoreCase = TRUE
                .Global     = TRUE
                .Pattern    = "\s(?:-|/)lang=(\w{2,})"

                For Each  vTmp  In  .Execute( nFirstPage )

                    With  vTmp.SubMatches

                        If  .Count > 0  Then

                            strValue = UCase( .Item( 0 ))
                            Exit For

                        End If

                    End With

                Next

            End With

        End If

        End If

end if

    Dim vArgs
    vArgs = window.dialogarguments
    Set g_objDict = vArgs(4)
 
                                    '***** Application object setup *****
    With Application.Member

        .AddProperty  "VBConstants",           , 2, strName
        .AddProperty  "onBeforeUnloadHdl",     ,  , Nothing
        .AddProperty  "LanguageID",            , 2, strValue
        .AddProperty  "MainWindowHandle",      , 2, g_objWinHandle

        .AddMethod    "CreateSTObject",     GetRef( "ScriptingToolsCreator"     ), 0
        .AddMethod    "GetNumberOfTabs",    GetRef( "GetNumberOfTabs"           ), 0
        .AddMethod    "ActivateTab",        GetRef( "ActivateTab"               ), 0
        .AddMethod    "GetActiveTab",       GetRef( "GetActiveTab"              ), 0
        .AddMethod    "GetActiveTabName",   GetRef( "GetActiveTabName"          ), 0
        .AddMethod    "GetDataOfActiveTab", GetRef( "GetDataOfActiveTab"        ), 0
        .AddMethod    "CreateProgressDlg",  GetRef( "CreateProgressDlg"         ), 0

        .AddChild     "Pages",  16
        .AddChild     "Paths",   4
        .AddChild     "PLC",    16
        .AddChild     "Output", 16



        With .Item( "Paths" ).Member

            .AddProperty    "Base",      , 2, strPath
            .AddProperty    "HTML",      , 2, g_objFSO.BuildPath( strPath, "HTML"      )
            .AddProperty    "Resources", , 2, g_objFSO.BuildPath( strPath, "Resources" )
            .AddProperty    "ScriptLib", , 2, g_objFSO.BuildPath( strPath, "ScriptLib" )

        End With

        With .Item( "Output" ).Member

            .AddMethod  "AddHTML",              GetRef( "AddHTMLToOutput"           ), 0
            .AddMethod  "AddError",             GetRef( "AddErrorToOutput"          ), 0
            .AddMethod  "AddSeparator",         GetRef( "AddSeparatorToOutput"      ), 0
            .AddMethod  "AddText",              GetRef( "AddTextToOutput"           ), 0
            .AddMethod  "AddTextWithStyle",     GetRef( "AddTextWithStyleToOutput"  ), 0
            .AddMethod  "AddTextWithClass",     GetRef( "AddTextWithClassToOutput"  ), 0
            .AddMethod  "AddTextToLastLine",    GetRef( "AddTextToOutputLastLine"   ), 0
            .AddMethod  "AddMessageWithPLC",    GetRef( "AddMessageWithPLCToOutput" ), 0
            .AddMethod  "RemoveLastLine",       GetRef( "RemoveOutputLastLine"      ), 0
            .AddMethod  "GetSizeFactor",        GetRef( "GetOutputSizeFactor"       ), 0
            .AddMethod  "SetSizeFactor",        GetRef( "SetOutputSizeFactor"       ), 0

        End With

    End With

'    ExecuteGlobal  strName

'    With  Document.Body
'
'        If  ( .ClientWidth < 900 )  Or  ( .ClientHeight < 700 )  Then  g_objWinHandle.Show eShowWindow_Maximized
'
'    End With

'    g_objWinHandle.Title = Base_GetText( "HeaderText" )

'    Window_onResize

'    g_objWinHandle.InstallMessageFilter GetRef( "MsgHdl" ), WM_SYSCOMMAND, WM_SYSCOMMAND

End Sub

'***********************************************************************
'****                                                               ****
'****       Handler for form unloading                              ****
'****                                                               ****
'***********************************************************************

Sub Window_OnUnload()

    Dim vTmp
 
End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Function executes given string within context of         *
'*             script of dialog                                         *
'*                                                                      *
'*  NEED:      String with VBScript statements                          *
'*             Arbitrary parameter                                      *
'*                                                                      *
'*  RETURN:    If return value is needed, use 'Command'                 *
'*                                                                      *
'************************************************************************

Function Command( strExec, vParameter )

    Execute  strExec

End Function

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************

Sub STANDARDDLG_BUTTONS_onClicked()
    
    If STANDARDDLG_BUTTONS.Value = 1 Then
    
    ElseIF STANDARDDLG_BUTTONS.Value = 256 Then
        if not fctPrintHTML() then 
        msgbox("Print failed!")
    End If
    Else
     '   msgbox("STANDARDDLG_BUTTONS.Value: " & STANDARDDLG_BUTTONS.Value)
        
    End if
    window.close

End Sub

'***********************************************************************
'****                                                               ****
'****       Converting from millimeter to inches                    ****
'****                                                               ****
'***********************************************************************

function fctMMToInch( nValue )

    Dim n
    Dim nAbs
    Dim nInt

    n    = Round( CDbl( nValue ) / 25.4, 1 )
    nAbs = Abs( n    )
    nInt = Int( nAbs )
    
    if ( nAbs - Int( nAbs )) >= 0.5 then nInt = nInt + 0.5

    fctMMToInch = nInt
    if n < 0 then fctMMToInch = -nInt

end function

'***********************************************************************
'****                                                               ****
'****                      Print String                             ****
'****                                                               ****
'***********************************************************************

Function fctPrintHTML()

    fctPrintHTML = False

   window.print

    fctPrintHTML = True

   Exit Function

    Do 
       Dim strMain, strTemplate
        strMain = objFSO.GetAbsolutePathName("")
        strMain = objFSO.BuildPath(strMain, "Html")
            strTemplate = cTemplate
            strMain  = strMain + strTemplate

        Dim objTestDocPage    
        Set objTestDocPage = ShowModelessDialog( strMain )
        
        objTestDocPage.idTimeDate.InnerText = Date & "  " & Time
        
        Dim objTestDoc     
        Set objTestDoc     = objTestDocPage.document  
        
        Dim nPageheight, bPortrait
        bPortrait = True
        If bPortrait = True then
            nPageheight = 24.8
        Else
            nPageheight = 16.5
        End If

        Dim strToPrint, strData, strHeader, strFooter  
    
        strHeader = objTestDocPage.idHEADER.InnerHTML
        strData = idDataTable.InnerHtml
        strFooter = objTestDocPage.idFooter.InnerHTML
        strToPrint = "<HTML>"&_
                "<Body>"  &_
                strHeader &_
                "<div style='height:" & nPageheight & "cm'>" &_
                strData   &_
                "</div>"&_
                strFooter &_
                "<Body>"  &_
                "</HTML>" 

        Dim objfile
        Set objfile = objFSO.createtextfile("e:\testFile.html")
        objfile.WriteLine( strToPrint )
            objfile.Close

        window.ShowModalDialog "PrintTemplate.HTML", Array( strToPrint, bPortrait )


    Dim objPrintDriver
    Set objPrintDriver          = CreateObject( "IEHTMLToolsSO.TemplatedPrinting" )
    Set objPrintDriver.Document = objTestDoc
    objPrintDriver.Execute        "PrintTemplate.HTML"   _
                                & "?"                                  _
                                & 0_                
                                & "|"                                  
                                '& Application.ApplicationProxyID
    
    g_nPrintJobID               = g_nPrintJobID + 1

                                                    ' Needed to trigger printing....
    g_objDocWinHandle.PostMessage  &H0100, 0, 0
    g_objDocWinHandle.PostMessage  &H0101, 0, &HC0000000


        fctPrintHTML = True
        
    Loop While False
    
End Function

'***********************************************************************
'****                                                               ****
'****            Context menu handler for output pane               ****
'****                                                               ****
'***********************************************************************

Sub IFRAMEonContextMenu( objEvt )

    Dim nCounter
    Dim strItem
    Dim vMenuItems
    Dim objMenu
    Dim objSource

    objEvt.CancelBubble = TRUE
    nCounter            = 0

    On Error Resume Next

        With objEvt.SrcElement

            If      .OwnerDocument Is g_objLogHost.OwnerDocument  Then

                Set objSource = g_objLogHost

            ElseIf  .OwnerDocument Is iFrameView.Document Then

                Set objSource = iFrameView.Document.Body
                nCounter      = 1

            End If

        End With

    On Error Goto 0


    If  IsEmpty( objSource )  Then  Exit Sub

    vMenuItems  = Split( Replace( BASE_GetText( "MenuOutput" ), "~", "&" ), "|"   )
    nCounter    = UBound( vMenuItems ) - LBound( vMenuItems ) + 1 - nCounter
    Set objMenu = ScriptingToolsCreator( "ScriptingToolsSO.PopupMenu" )

    For Each strItem In vMenuItems

        If  nCounter > 0  Then  objMenu.Add  strItem

        nCounter = nCounter -1

    Next


    Select Case  objMenu.Show( g_objWinHandle, objEvt.ScreenX, objEvt.ScreenY )

        Case 1  With  ScriptingToolsCreator( "ScriptingToolsSO.FileSelectDialog" )

                    .NoReadOnlyFiles = TRUE
                    .AutoExtension   = TRUE
                    .AddFilter  .GetFileTypeText( ".HTML" ) & " (*.HTML;*.HTM)", "*.HTML;*.HTM"
                    .AddFilter  .GetFileTypeText( ".TXT"  ) & " (*.TXT)",        "*.TXT"

                    If  Not .Save( Document )  Then  Exit Sub

                    strItem = g_objFSO.BuildPath( .FilePath, .FileName )

                End With

                On Error Resume Next

                    g_objFSO.DeleteFile  strItem, TRUE

                On Error Goto 0

                With  g_objFSO.CreateTextFile( strItem, TRUE )

                    If     UCase( Left( g_objFSO.GetExtensionName( strItem ), 3 )) = "HTM"  Then

                        .WriteLine  objSource.Document.DocumentElement.OuterHTML

                    ElseIf UCase(       g_objFSO.GetExtensionName( strItem ))      = "TXT"  Then

                        .WriteLine  objSource.InnerText

                    End If

                End With

        Case 2  With Window.ClipBoardData

                    .ClearData
                    .SetData   "Text", objSource.InnerText

                End With

        Case 3  With Window.ClipBoardData

                    .ClearData
                    .SetData   "Text", objSource.Document.DocumentElement.OuterHTML

                End With

        Case 4  objSource.InnerHTML = ""
        

    End Select

End Sub

'***********************************************************************
'****                                                               ****
'****       Handler for suppressing WM_CLOSE                        ****
'****                                                               ****
'***********************************************************************

Sub MsgHdl( objMessage )

    If  objMessage.wParam <> SC_CLOSE  Then  Exit Sub

    If  Not ( g_objBeforeQuitHdl Is Nothing )  Then

        If  Not g_objBeforeQuitHdl()  Then

            objMessage.Msg = 0
            Exit Sub

        End If

    End If
    
End Sub

'***********************************************************************
'****                                                               ****
'****       Handler for tool tab clicks                             ****
'****                                                               ****
'***********************************************************************

Sub onClickToolTab

    Dim vTmp
    Dim nI
    Dim objElem

    With Window.Event

        .CancelBubble = TRUE
        nI            = 0
        Set objElem   = .SrcElement

        For  Each vTmp In divFrameHostTool.Children

            If  vTmp Is objElem  Then 
            
                ActivateTab  nI
                Exit For

            End If

            nI = nI + 1

        Next


    End With

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   V I D E O    B U T T O N                         *
'*                                                                              *
'********************************************************************************

Sub VideoButton_onclick( objEvt )

    objEvt.CancelBubble = TRUE

    EnableLiveVideo  CBool( g_nVideoTickerCookie = 0 )

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Enables/disables displaying live video images            *
'*                                                                      *
'*  NEED:      If set to TRUE, live video images will be displayed      *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub EnableLiveVideo( bIsEnabled )

    If bIsEnabled  Then

        If  g_nVideoTickerCookie <> 0  Then  EnableLiveVideo  FALSE

        If Not ( g_objApplication Is Nothing )  Then

            g_nVideoTickerCookie = SetInterval(  GetRef( "VideoTicker" ), 10, "VBScript" )
            g_nVideoTickerTime   = DateAdd( "s", -60, Now )

            g_objApplication.ImageProcessing.ClearStop

        Else

            g_nVideoTickerCookie = SetInterval(  GetRef( "DummyVideoTicker" ), 10, "VBScript" )

        End If

        buttonVideo.Checked = TRUE

    Else

        If g_nVideoTickerCookie <> 0  Then

            ClearInterval g_nVideoTickerCookie

            g_nVideoTickerCookie = 0

        End If

        buttonVideo.Checked = FALSE

    End If

End Sub

'********************************************************************************
'*                                                                              *
'*  H A N D L E R   F O R   T I M E R   T I C K S   O F   L I V E   V I D E O   *
'*                                                                              *
'********************************************************************************

Sub VideoTicker()

    With g_objApplication.ImageProcessing

        .SnapNoAdjust
        .Display

    End With

    If DateDiff( "s", g_nVideoTickerTime, Now ) > 20  Then

        g_nVideoTickerTime = Now

        If g_objMachineState.Chamber.IsClosed  Then  g_objMachine.ActivateLaser

    End If

End Sub

Sub DummyVideoTicker()

End Sub
'************************************************************************
'*                                                                      *
'*  TASK:      Returns number of tab pages                              *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    Number of tabs pages                                     *
'*                                                                      *
'************************************************************************

Function  GetNumberOfTabs()

    GetNumberOfTabs = divFrameHostTool.Children.Length

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Loads given HTML page into given frame                   *
'*                                                                      *
'*  NEED:      Index of tab window to be displayed (zero based)         *
'*             URL of page to be displayed                              *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub ActivateTab( nTabIndex )

    Dim vTmp
    Dim objElem
    Dim colTabs

    If  nTabIndex = g_nCurPage Then  Exit Sub

    Set colTabs = divFrameHostTool.Children

    If  nTabIndex >= colTabs.Length Then  Exit Sub

    Set objElem = colTabs.Item( nTabIndex )

    If  objElem.IsDisabled  Then  Exit Sub


    If  Not ( g_objBeforeQuitHdl Is Nothing )  Then

        If  Not g_objBeforeQuitHdl()  Then  Exit Sub

    End If

    g_nCurPage              =               nTabIndex
    g_strCurPageName        = objElem.InnerText
    spanActiveTab.Title     = objElem.Title
    spanActiveTab.InnerText = g_strCurPageName

    With spanActiveTab.RuntimeStyle

        .PixelLeft  = divFrameHostTool.OffsetLeft           _
                    + objElem.OffsetLeft                    _
                    - Round( objElem.OffsetWidth * 0.1 )
        .Visibility = "Visible"

    End With

    g_objWinHandle.Title = Base_GetText( "HeaderText" )     _
                         & "  -  "                          _
                         & g_strCurPageName

    DisplayPage  objElem.GetAttribute(  C_strAttibToolPath  )

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Returns index of active tab                              *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    Index of tab window displayed currently (zero based)     *
'*                                                                      *
'************************************************************************

Function GetActiveTab

    GetActiveTab = g_nCurPage

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Returns name of active tab                               *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    Name of tab window displayed currently (zero based)      *
'*                                                                      *
'************************************************************************

Function GetActiveTabName

    GetActiveTabName = g_strCurPageName

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Returns 'DynamicData' root object                        *
'*             associated with active tab                               *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    DynamicData' root object                                 *
'*                                                                      *
'************************************************************************

Function GetDataOfActiveTab

    Dim strName
    strName = "Page" & g_nCurPage

    With Application.Pages.Member

        If  Not .ExistChild( strName )  Then  .AddChild  strName

        Set GetDataOfActiveTab = .Item( strName )

    End With

End Function


'************************************************************************
'*                                                                      *
'*  TASK:      Loads given HTML page into given frame                   *
'*                                                                      *
'*  NEED:      URL of page to be displayed                              *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub DisplayPage( strPageURL )

    Dim  vTmp
    Dim  objFrame

    Set  g_objBeforeQuitHdl             =  Nothing
    Set  Application.onBeforeUnloadHdl  =  Nothing
    Set  objFrame                       =  iframeView

    With objFrame.Document

        .Write ""
        .Close

    End With

    SetOutputSizeFactor  1.0 - C_nDefaultSizeFactor
    objFrame.Navigate    strPageURL

    With objFrame.Document

        Do

            g_objVBXens.Sleep  20, eVBSXSleepPostponeUserActions

        Loop While .ReadyState <> "complete"

        For Each vTmp In .Scripts

            Do

                g_objVBXens.Sleep  20, eVBSXSleepPostponeUserActions

            Loop While vTmp.ReadyState <> "complete"

        Next

        .DocumentElement.SetAttribute   C_AttrNameAppObj,  Application

        .ParentWindow.ExecScript    g_strStartupScript, C_LangVBScript

        Do

                g_objVBXens.Sleep  100, eVBSXSleepPostponeUserActions

        Loop While .ReadyState <> "complete"

       .AttachEvent  "oncontextmenu", GetRef( "IFRAMEonContextMenu" )

    End With

    Set  g_objBeforeQuitHdl  =  Application.onBeforeUnloadHdl

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Adds given HTML code text to output pane                 *
'*                                                                      *
'*  NEED:      HTML code to add                                         *
'*                                                                      *
'*  RETURN:    HTML object created by given HTML code                   *
'*                                                                      *
'************************************************************************

Function AddHTMLToOutput( strHTML )

    g_objLogHost.InsertAdjacentHTML  "BeforeEnd", strHTML
    Set AddHTMLToOutput = g_objLogHost.LastChild
    iframeLOG.SetTimeout "Window.ScrollTo 0, &H7FFFFFFF", 100, C_LangVBScript
    
End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Adds given text to output pane, formatted as error       *
'*                                                                      *
'*  NEED:      Text to add                                              *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Function AddErrorToOutput( strText )

    Set AddErrorToOutput = AddHTMLToOutput( "<DIV CLASS=""classERROR""><BR></DIV>" )
    AddErrorToOutput.InsertAdjacentText  "BeforeEnd", strText

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Adds given text to output pane                           *
'*                                                                      *
'*  NEED:      Text to add                                              *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Function AddSeparatorToOutput( strText )

    Set AddSeparatorToOutput = AddHTMLToOutput(  xmlSeparator.DocumentElement.XML  )

    If  Len( strText ) < 1  Then

        AddSeparatorToOutput.RemoveChild  AddSeparatorToOutput.LastChild

    Else

        AddSeparatorToOutput.LastChild.InnerText = strText

    End If

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Adds given text to output pane                           *
'*                                                                      *
'*  NEED:      Text to add                                              *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Function AddTextToOutput( strText )

    Set AddTextToOutput       = AddHTMLToOutput( "<PRE></PRE>" )
    AddTextToOutput.InnerText = strText

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Adds given text to output pane, using given HTML styles  *
'*                                                                      *
'*  NEED:      Text to add                                              *
'*             HTML style string                                        *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Function AddTextWithStyleToOutput( strText, strStyle )

    Set AddTextWithStyleToOutput       = AddHTMLToOutput( "<PRE STYLE=""" & strStyle & """></PRE>" )
    AddTextWithStyleToOutput.InnerText = strText

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Adds given text to output pane, using given HTML styles  *
'*                                                                      *
'*  NEED:      Text to add                                              *
'*             HTML style class name                                    *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Function AddTextWithClassToOutput( strText, strCSSClass )

    Set AddTextWithClassToOutput       = AddHTMLToOutput( "<PRE CLASS=""" & strCSSClass & """></PRE>" )
    AddTextWithClassToOutput.InnerText = strText

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Adds given text to last output line                      *
'*                                                                      *
'*  NEED:      Text to add                                              *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub AddTextToOutputLastLine( strText )

    g_objLogHost.LastChild.InsertAdjacentText "BeforeEnd", strText

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Formats given string and adds it to output               *
'*             after PLC identification was successful.                 *
'*                                                                      *
'*  NEED:      String containing dotted IP address of PLC               *
'*             Port number of PLC                                       *
'*             Text containing a '\0' at insertion position for         *
'*                address and port                                      *
'*             HTML style class name                                    *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Function AddMessageWithPLCToOutput( strDottedIPAdr, nPort, strText, strCSSClass )

    Dim vTmp

    vTmp = BASE_GetText( "PLCBaseMsg"               )
    vTmp = Replace( vTmp,    "\0",  strDottedIPAdr  )
    vTmp = Replace( vTmp,    "\1",  nPort           )
    vTmp = Replace( strText, "\0",  vTmp            )

    Set AddMessageWithPLCToOutput = AddTextWithClassToOutput( vTmp, strCSSClass )

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Removes last output line                                 *
'*                                                                      *
'*  NEED:      Text to add                                              *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub RemoveOutputLastLine

    Dim objLastLine
    Set objLastLine = g_objLogHost.LastChild

    If Not ( objLastLine Is Nothing )  Then  g_objLogHost.RemoveChild  objLastLine

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Creates an 'CProgressDlgHolder' object containing        *
'*             a progress dialog.                                       *
'*             If such a dialog already exists, it is used; otherwise   *
'*             it is created                                            *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:   'CProgressDlgHolder' object                                        *
'*                                                                      *
'************************************************************************

Function CreateProgressDlg

    Set CreateProgressDlg = New CProgressDlgHolder

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Returns size factor of output pane                       *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    Factor of output pane size (related to window height)    *
'*                                                                      *
'************************************************************************

Function GetOutputSizeFactor

    GetOutputSizeFactor = 1.0 - g_nOutputSizeFactor

End Function

'************************************************************************
'*                                                                      *
'*  TASK:     Sets size factor of output pane                           *
'*                                                                      *
'*  NEED:     Factor of output pane size (related to window height)     *
'*                                                                      *
'*  RETURN:   ---                                                       *
'*                                                                      *
'************************************************************************

Sub SetOutputSizeFactor( ByVal nFactor )

    If  nFactor < 0.1  Then  nFactor = 0.1
    If  nFactor > 0.9  Then  nFactor = 0.9

    nFactor = 1.0 - nFactor

    If  g_nOutputSizeFactor = nFactor  Then  Exit Sub

    g_nOutputSizeFactor = nFactor

    Window_onResize

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:     Function creates given object from registered             *
'*            ScriptingToolsSO DLL depending on version                 *
'*                                                                      *
'*  NEED:     Name of object as needed by 'CreateObject'                *
'*                                                                      *
'*  RETURN:   Object created or NOTHING if object can't be created      *
'*                                                                      *
'************************************************************************

Function ScriptingToolsCreator( strObjectName )

        Set ScriptingToolsCreator = CreateObject( strObjectName )

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Enables/disables STOP button                             *
'*                                                                      *
'*  NEED:      If set to TRUE, STOP button will be enabled              *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub EnableStopButton( bIsEnabled )

    buttonStopPos.Disabled = Not bIsEnabled

    With imgStopPos.filters

        .Item( "DXImageTransform.Microsoft.Emboss" ).Enabled = Not bIsEnabled
        .Item( "DXImageTransform.Microsoft.Alpha"  ).Enabled = Not bIsEnabled

    End With

End Sub

'***********************************************************************
'***********************************************************************
