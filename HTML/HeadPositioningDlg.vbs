'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   L i f t H e a d T o T i r e D l g         *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************

Option Explicit

Class CLiftModeData

    Public  bIsEnabled
    Public  nValue
    Public  nMin
    Public  nMax
    Public  strKey

    Private Sub Class_Initialize

        bIsEnabled  = FALSE
        nValue      = Null
        nMin        = 0
        nMax        = 0

    End Sub

End Class


const WM_CLOSE                  = &H0010


Const C_nLiftMeasModeCenter     =  0
Const C_nLiftMeasModeAbsolute   =  1
Const C_nLiftMeasModeRelative   =  2
Const C_nLiftMeasModeDive       =  3

Const C_LiftTypeKeepPosFlag = 16

Const C_RadioBtnGroupName       = "GroupRadioBtn"

Dim g_bPLCBC640
g_bPLCBC640 = False

Dim g_objApplication
Dim g_objMachine
Dim g_objMachineScript
Dim g_objMachineState
Dim g_objVBSXtensions
Dim g_objHTMLHelper
Dim g_objWinTools
Dim g_objHeadPosPopup
Dim g_vLiftModes
Dim g_bHasTiltAxis
Dim g_nDisablingCount
Dim g_nSelectedLiftMode
Dim g_nTimerTickerCookie
Dim g_nVideoTickerCookie
Dim g_nVideoTickerTime
Dim g_nMachineTickerCookie
Dim g_nAnimTickerCookie
Dim g_nAnimTickerCount
Dim g_nAnimObject
Dim g_bAnimIsFGColor
Dim g_nAnimCol1
Dim g_nAnimCol2
Dim g_nGITCookie1
Dim g_nGITCookie2


Set g_objApplication            = Nothing
Set g_objMachine                = Nothing
Set g_objMachineScript          = Nothing
Set g_objMachineState           = Nothing
Set g_objHeadPosPopup           = Nothing
Set g_objHTMLHelper             = CreateObject( "IEHTMLToolsSO.HTMLHelper"      )
Set g_objVBSXtensions           = CreateObject( "ScriptingToolsSO.VBSXtensions" )
Set g_objWinTools               = CreateObject( "ScriptingToolsSO.Windowing"    )


g_vLiftModes                    = Array(    New  CLiftModeData, _
                                            New  CLiftModeData, _
                                            New  CLiftModeData, _
                                            New  CLiftModeData  )

g_nDisablingCount               = 0
g_bHasTiltAxis                  = FALSE
g_nSelectedLiftMode             = C_nLiftMeasModeCenter
g_nTimerTickerCookie            = 0
g_nVideoTickerCookie            = 0
g_nVideoTickerTime              = Now
g_nAnimTickerCookie             = 0
g_nAnimTickerCount              = 0
g_nMachineTickerCookie          = 0

g_nGITCookie1                   = 0
g_nGITCookie2                   = 0

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

    vArgs = xmlLiftMenu.DocumentElement.GetAttribute( "Types" )
    vArgs = Split( vArgs, "|" )
    nI    = 0

    For Each vTmp In  g_vLiftModes

        vTmp.strKey = vArgs( nI )
        nI          = nI + 1

    Next

    vArgs = Null

    SetPosDisplay  spineditLift,  spanLiftPos  
    SetPosDisplay  spineditDiam,  spanDiamPos
    SetPosDisplay  spineditTilt,  spanTiltPos
    SetPosDisplay  spineditRot,   spanRotPos


    spineditLift.Precision   = 0
    spineditDiam.Precision   = 0
    spineditTilt.Precision   = 0
    spineditRot.Precision    = 0

    If Not Document.Body.IsDialog Then

        Randomize

        With  g_vLiftModes( C_nLiftMeasModeCenter )

            .bIsEnabled = TRUE
            .nMin       = -30
            .nMax       =  30
            .nValue     = Round( -30 + 60 * Rnd )

        End With

        With  g_vLiftModes( C_nLiftMeasModeAbsolute )

            .bIsEnabled = TRUE
            .nMin       = 100
            .nMax       = 600
            .nValue     = Round( 100 + 500 * Rnd )

        End With

        With  g_vLiftModes( C_nLiftMeasModeRelative )

            .bIsEnabled = TRUE
            .nMin       =   0
            .nMax       = 500
            .nValue     = Round( 500 * Rnd )

        End With

        With  g_vLiftModes( C_nLiftMeasModeDive )

            .bIsEnabled = TRUE
            .nMin       =   0
            .nMax       = 240
            .nValue     = Round( 240 * Rnd )

        End With


        SelectLiftMode    C_nLiftMeasModeCenter

        spineditDiam.Min      = 220
        spineditDiam.Max      = 440
        spineditDiam.Value    = Round( 220 + 220 * Rnd )

        spineditTilt.Min      = -45
        spineditTilt.Max      =  45
        spineditTilt.Value    = Round( -45 +  90 * Rnd )

        SetHasTiltAxis   Rnd > 0.5

        spineditRot.Min       =   0
        spineditRot.Max       = 350
        spineditRot.Value     = Round(       350 * Rnd )

        DisplayOuterDiameter    Round( 800 + 400 * Rnd )
        DisplayInnerDiameter    Round( 400 + 200 * Rnd )
        DisplayTireWidth        Round( 160 + 340 * Rnd )
        DisplayExcentricity     Round(        10 * Rnd )

        EnableStopButton  FALSE

    Else

        With Document.ParentWindow

            vArgs  =  .DialogArguments

            Set g_objApplication    = vArgs( 0 )
            Set g_objMachineScript  = vArgs( 1 )

        End With

        Set g_objMachine      = g_objApplication.Machine
        Set g_objMachineState = g_objMachine.State

        If g_objApplication.ScriptEngines.Member.ExistChild ("BC640Head") Then g_bPLCBC640 = True

        g_objWinTools.RepaintWindowTitle  g_objApplication.HWND, TRUE

        SetHasTiltAxis  g_objMachineState.Member.ExistChild( "HeadTilt0" )

        SetupLiftInputs
        SetupDiamInput
        SetupTiltInput
        SetupRotInput

        With g_objMachineState.Tire

            DisplayOuterDiameter    .OuterDiameter
            DisplayInnerDiameter    .InnerDiameter
            DisplayTireWidth        .Width
            DisplayExcentricity     .Excentricity

        End With

    End If

    With divHeadPosPane.Header

        .AttachEvent                "onmouseover",  GetRef(  "HeadPosInfo_onShow"   )
        .AttachEvent                "onmouseout",   GetRef(  "HeadPosInfo_onHide"   )

    End With

    spanLiftType.AttachEvent        "onclick",      GetRef(  "LiftType_onClick"     )
    buttonUpdatePos.AttachEvent     "onclick",      GetRef(  "PosButton_onclick"    )
    buttonHomePos.AttachEvent       "onclick",      GetRef(  "HomeButton_onclick"   )
    buttonStopPos.AttachEvent       "onclick",      GetRef(  "StopButton_onclick"   )

    buttonVideo.AttachEvent         "onclick",      GetRef(  "VideoButton_onclick"  )

    Document.AttachEvent            "onkeydown",    GetRef(  "DocBasic_onKeyDown"   )
    Document.AttachEvent            "onkeydown",    GetRef(  "DocPlus1_onKeyDown"   )

    EnableStopButton  FALSE

    divTireDataPane.HeaderStyle = "background-color:;"
    divHeadPosPane.HeaderStyle  = "background-color:;"
    divToolPane.HeaderStyle     = "background-color:;"

    divTireDataPane.SizeToContent "H"
    divTireDataPane.CenterContent "X"
    divHeadPosPane.SizeToContent  "H"
    divHeadPosPane.CenterContent  "X"
    divToolPane.SizeToContent     "H"
    divToolPane.CenterContent     "X"

    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )
    STANDARDDLG_BUTTONS.Default = 0

    g_nTimerTickerCookie = SetInterval(  GetRef( "TimerTicker" ), 10, "VBScript" )
    g_nMachineTickerCookie = MySetInterval(  GetRef( "MachineTicker" ), 100 )

    If Document.Body.IsDialog Then

        If IsObject( vArgs( 3 ))  Then

            Set vTmp = vArgs( 3 )

            With CreateObject( "ScriptingToolsSO.ProxyFactory")

                g_nGITCookie1 = .Register( Document )
                g_nGITCookie2 = .Register( GetRef( "Command" ))

            End With

            If Not ( vTmp Is Nothing )  Then  vTmp  g_nGITCookie1, g_nGITCookie2

        End If

        Document.ParentWindow.DialogTop  = vArgs( 2 )

    End If

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

'************************************************************************
'*                                                                      *
'*  TASK:      Function adds tool pane to dialog                        *
'*                                                                      *
'*  NEED:      String with ID of new tool pane                          *
'*                                                                      *
'*  RETURN:    Created tool pane                                        *
'*                                                                      *
'************************************************************************

Function AddToolPane( strID )

    Set AddToolPane = InsertToolPane(  strID, Empty )

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Function inserts tool pane in dialog before tool pane    *                       
'*             given ID                                                 *
'*                                                                      *
'*  NEED:      String with ID of new tool pane                          *
'*                                                                      *
'*  RETURN:    Created tool pane                                        *
'*                                                                      *
'************************************************************************

Function InsertToolPane( strID, nPane )

    Set InsertToolPane = divToolPane.CloneNode( FALSE )
    InsertToolPane.ID  = strID

    InsertToolPane.Style.Behavior  =  divToolPane.Style.Behavior    '*** Needed for IE8

    Do 

        g_objVBSXtensions.Sleep 50

    Loop While InsertToolPane.readyState <> "complete"


    If  IsEmpty(  nPane  )  Then

        divToolPane.ParentElement.AppendChild   InsertToolPane

    Else

        divToolPane.ParentElement.InsertBefore  InsertToolPane,         _
                                                divToolPane             _
                                                 .ParentElement         _
                                                  .Children( nPane )
    End If

End Function

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************
'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   L I F T   T Y P E   S E L E C T I O N            *
'*                                                                              *
'********************************************************************************

Sub LiftType_onClick( objEvt )

    Dim vTmp
    Dim nI
    Dim nRadio
    Dim nSelected
    Dim objMenuItem
    Dim objPopupMenu
    Set objPopupMenu = CreateObject( "ScriptingToolsSO.PopupMenu" )

    nSelected        = GetSelectedLiftMode()

    For Each objMenuItem In xmlLiftMenu.DocumentElement.ChildNodes

        If objMenuItem.NodeType = 1  Then

            If IsEmpty( nI ) Then

                objPopupMenu.Add   objMenuItem.Text
                nI               = -1

            Else

                vTmp = objPopupMenu.Add(  "&" & objMenuItem.Text  )

                If Not g_vLiftModes( nI ).bIsEnabled Then  objPopupMenu.Gray( vTmp ) = TRUE

                If nI = nSelected  Then  nRadio = vTmp

            End If

            nI = nI + 1

        End If

    Next

    objPopupMenu.Default( 1 ) = TRUE

    If Not IsEmpty( nRadio )  Then objPopupMenu.Radio  nRadio

    vTmp = objPopupMenu.Show( Document.Body.WindowHandle, objEvt.ScreenX, objEvt.ScreenY )

    If vTmp > 1  Then  SelectLiftMode vTmp - 2

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   H O M E    B U T T O N                           *
'*                                                                              *
'********************************************************************************

Sub HomeButton_onclick( objEvt )

    objEvt.CancelBubble = TRUE

    EnableStopButton            TRUE
    EnableInputs                FALSE

    ShowAnimated  buttonHomePos, FALSE, "#00FF00", "#008000", 500

    If Not     ( g_objMachineScript Is Nothing )  _
       And Not ( g_objMachine       Is Nothing )  Then      ' Only if dialog is shown by app

        g_objMachine.ClearStop
        g_objMachine.ClearErrorText

        Do
         
            If g_bHasTiltAxis Then
            If Not g_objMachine.TiltHeadHome()       Then Exit Do
            End If

          If g_bPLCBC640 Then
            If Not g_objMachine.LiftHeadHome( False) Then Exit Do
          Else
            If Not g_objMachine.LiftHeadHome( ) Then Exit Do            
          End If
        
            g_objMachineScript.onHeadHome

        Loop While False

        SetupLiftInputs
        SetupDiamInput
        SetupTiltInput
        DisplayRotPos
        
    Else

        g_objVBSXtensions.Sleep 5000    

    End If

    StopAnimation  buttonHomePos

    If g_vLiftModes( C_nLiftMeasModeAbsolute ).bIsEnabled  Then 

        SelectLiftMode   C_nLiftMeasModeAbsolute

    End If


    EnableStopButton            FALSE
    EnableInputs                TRUE

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   P O S I T I O N   U P D A T E   B U T T O N      *
'*                                                                              *
'********************************************************************************

Sub PosButton_onclick( objEvt )

    objEvt.CancelBubble = TRUE

    EnableStopButton            TRUE
    EnableInputs                FALSE

    If Not ( g_objMachine Is Nothing )  Then        ' Only if dialog is shown by app

        g_objMachine.ClearStop
        g_objMachine.ClearErrorText

        Dim objPos
        Set objPos = GetSelectedPositions()
        PositionTo   objPos,  TRUE,  TRUE

    Else

        g_objVBSXtensions.Sleep 1000    

    End If

    EnableStopButton            FALSE
    EnableInputs                TRUE

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   S T O P    B U T T O N                           *
'*                                                                              *
'********************************************************************************

Sub StopButton_onclick( objEvt )

    objEvt.CancelBubble = TRUE

    If Not ( g_objMachine Is Nothing )  Then        ' Only if dialog is shown by app

        g_objMachine.StopAll

    End If

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

'********************************************************************************
'*                                                                              *
'*     H A N D L E R    F O R    K E Y B O A R D    I N P U T                   *
'*                                                                              *
'********************************************************************************

Sub DocBasic_onKeyDown( objEvt )

    objEvt.CancelBubble = TRUE

    Select Case objEvt.KeyCode

        Case 116    objEvt.ReturnValue = 0
                    If Not buttonUpdatePos.IsDisabled  Then  PosButton_onclick  objEvt  : Exit Sub

        Case 117    objEvt.ReturnValue = 0
                    If Not buttonHomePos.IsDisabled    Then  HomeButton_onclick  objEvt : Exit Sub

        Case 118    objEvt.ReturnValue = 0
                    If Not buttonStopPos.IsDisabled    Then  HomeButton_onclick  objEvt : Exit Sub

    End Select

    If objEvt.AltKey  Then

        If objEvt.KeyCode = 115 Then        ' Cancel Alt+F4

            objEvt.KeyCode     = 0
            objEvt.ReturnValue = 0
            Exit Sub

        End If

    End If

    Dim objLiftMode
    Dim strSCKey
    Dim i

    If objEvt.CtrlKey  Then

        i        = C_nLiftMeasModeCenter
        strSCKey = UCase( Chr( objEvt.KeyCode ))

        For Each objLiftMode  In  g_vLiftModes

            If strSCKey = objLiftMode.strKey  Then

                SelectLiftMode  i
                objEvt.ReturnValue = 0
                Exit Sub

            End If

            i = i + 1

        Next

    End If



End Sub

'********************************************************************************

Sub DocPlus1_onKeyDown( objEvt )

    objEvt.CancelBubble = TRUE

    Select Case objEvt.KeyCode

        Case 119    objEvt.ReturnValue = 0
                    VideoButton_onclick  objEvt

    End Select

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************

Sub STANDARDDLG_BUTTONS_onClicked()

    If ( g_nGITCookie1 <> 0 )  Or  ( g_nGITCookie2 <> 0 )  Then

        With CreateObject( "ScriptingToolsSO.ProxyFactory")

            If g_nGITCookie1 <> 0  Then  .Unregister  g_nGITCookie1
            If g_nGITCookie2 <> 0  Then  .Unregister  g_nGITCookie2

        End With

    End If

    Document.Body.WindowHandle.InstallMessageFilter Nothing, 0, 0
    
    g_objWinTools.RepaintWindowTitle  g_objApplication.HWND, FALSE

    Window.Close

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   S H O W I N G   H E A D   P O S   I N F O        *
'*                                                                              *
'********************************************************************************

Sub HeadPosInfo_onShow( objEvt )

    Dim objMasterStyle
    Dim strInfo
    Dim nX
    Dim nY

    If g_objApplication Is Nothing  Then Exit Sub

    strInfo = "<TABLE>"                             _
            & "<THEAD><TR><TD colspan=3 "           _
            & "align=center STYLE=""Padding:7px; "  _
            & "Font-Weight:900"">"                  _
            & divHeadPosPane.HeaderText             _
            & "</TD></TR></THEAD>"                  _
            & "<COL></COL><COL align=right "        _
            & "STYLE=""Font-Weight:700"">"          _
            & "</COL><COL></COL>"                   _
            & "<TBODY><TR><TD NoWrap>"              _
            & tdLift.InnerText                      _
            & "</TD><TD NoWrap>"                    _
            & g_objMachineState.HeadHeight.Value    _
            & "</TD><TD NoWrap>mm</TD></TR>"        _
            & "<TR><TD NoWrap>"                     _
            & tdDiam.InnerText                      _
            & "</TD><TD NoWrap>"                    _
            & g_objMachineState.HeadDiameter.Value  _
            & "</TD><TD NoWrap>mm</TD></TR>"

    If g_bHasTiltAxis  Then  

        strInfo = strInfo                           _
                & "<TR><TD NoWrap>"                 _
                & tdTilt.InnerText                  _
                & "</TD><TD NoWrap>"                _
                & g_objMachineState.HeadTilt0.Value  _
                & "</TD><TD NoWrap>&deg;</TD></TR>"

    End If

    strInfo = strInfo                               _
            & "<TR><TD NoWrap>"                     _
            & tdRot.InnerText                       _
            & "</TD><TD NoWrap>"                    _
            & g_objMachineState.HeadRotation.Value  _
            & "</TD><TD NoWrap>&deg;</TD></TR>"     _
            & "</TBODY></TABLE>"



    If Not ( g_objHeadPosPopup Is Nothing )  Then  g_objHeadPosPopup.Hide

    Set g_objHeadPosPopup = Window.CreatePopup

    With  g_objHeadPosPopup.Document.Body

        With .Style

            .BackgroundColor = "infobackground"
            .Overflow        = "hidden"
            .Border          = "1px outset"
            .Padding         = "5px"

        End With

        .NoWrap    = TRUE
        .InnerHTML = strInfo

        Set objMasterStyle = Document.Body.CurrentStyle

        With .FirstChild.Style

            .FontFamily     = objMasterStyle.FontFamily
            .FontSize       = objMasterStyle.FontSize  
            .FontWeight     = 500

        End With

        nX = Window.ScreenLeft                      _
           + divHeadPosPane.ClientLeft
        nY = Window.ScreenTop                       _
           + divHeadPosPane.OffsetTop               _
           + divHeadPosPane.Header.OffsetHeight \ 2


        g_objHeadPosPopup.Show  objEvt.ScreenX, objEvt.ScreenY, 10, 10

        g_objHeadPosPopup.Show  nX - .ScrollWidth,  _
                                nY,                 _
                                .ScrollWidth,       _
                                .ScrollHeight

    End With

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   H I D I N G   H E A D   P O S   I N F O          *
'*                                                                              *
'********************************************************************************

Sub HeadPosInfo_onHide( objEvt )

    If Not ( g_objHeadPosPopup Is Nothing )  Then  g_objHeadPosPopup.Hide

    Set g_objHeadPosPopup = Nothing

End Sub

'********************************************************************************
'*                                                                              *
'*   S U P P R E S S I N G    C L O S E    M E S S A G E S                      *
'*                                                                              *
'********************************************************************************

Sub MsgHdl( objMessage )

    objMessage.Msg = 0

End Sub

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

'************************************************************************
'*                                                                      *
'*  TASK:      Enables/disables all elements having an attribute        *
'*             named '_BLOCKINPUT_' set. Uses counter, so equal number  *
'*             of enabling/disabling calls are needed to take effect    *
'*                                                                      *
'*  NEED:      If set to TRUE, marked elements will gain enabling       *
'*                state set at previous disabling invokation            *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub EnableInputs( bIsEnabled )

    Dim bEnabling
    Dim objElem

    If  bIsEnabled  Then

        If g_nDisablingCount > 0  Then

            g_nDisablingCount = g_nDisablingCount - 1

            If g_nDisablingCount = 0  Then  bEnabling = TRUE

        End If

    Else

        If g_nDisablingCount = 0  Then  bEnabling = FALSE

        g_nDisablingCount = g_nDisablingCount + 1

    End If


    If Not IsEmpty( bEnabling )  Then

        For Each objElem In Document.All

            If Not IsNull( objElem.GetAttribute( "_BLOCKINPUT_" ))  Then

                If bEnabling  Then

                    objElem.Disabled = objElem.GetAttribute( "_PREVENA_" )

                Else

                    objElem.SetAttribute  "_PREVENA_",  objElem.Disabled
                    objElem.Disabled = TRUE

                End If

            End If

        Next

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Enables/disables particular spin edit input              *
'*                                                                      *
'*  NEED:      HTML spinedit object to be handled                       *
'*             If set to TRUE, spin edit will be activated; otherwise   *
'*                a simple text field will be displayed                 *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub EnableInput( objInput, bHasTo )

    If bHasTo  Then

        objInput.Style.Visibility             = "visible"
        objInput.NextSibling.Style.Visibility = "hidden"

    Else

        objInput.Style.Visibility   = "hidden"

        With objInput.NextSibling

            .Style.Visibility       = "visible"
            .InnerText              = objInput.Value

        End With 

    End If

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

            g_nVideoTickerCookie = MySetInterval(  GetRef( "VideoTicker" ), 10 )
            g_nVideoTickerTime   = DateAdd( "s", -60, Now )

            g_objApplication.ImageProcessing.ClearStop

        Else

            g_nVideoTickerCookie = MySetInterval(  GetRef( "DummyVideoTicker" ), 10 )

        End If

        buttonVideo.Checked = TRUE

    Else

        If g_nVideoTickerCookie <> 0  Then

            MyClearInterval g_nVideoTickerCookie

            g_nVideoTickerCookie = 0

        End If

        buttonVideo.Checked = FALSE

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Returns kind of lift positioning currently               *
'*             selected by user                                         *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    Lift mode as defined in 'C_nLiftMeasModeXXXXX' constants *
'*                                                                      *
'************************************************************************

Function GetSelectedLiftMode

    GetSelectedLiftMode = g_nSelectedLiftMode

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Set kind of lift positioning currently                   *
'*             selected by user                                         *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    Lift mode as defined in 'C_nLiftMeasModeXXXXX' constants *
'*                                                                      *
'************************************************************************

Function SetSelectedLiftMode ( nMode )

    g_nSelectedLiftMode = nMode

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Selects given lift mode. Only the selected mode will     *
'*                gain a spin edit input; others will display text      *
'*                                                                      *
'*  NEED:      Lift mode as defined in 'C_nLiftMeasModeXXXXX' constants *
'*                or corresponding HTML radio input object              *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub SelectLiftMode( nMode )

    If      ( nMode >= C_nLiftMeasModeCenter )  _
       And  ( nMode <= C_nLiftMeasModeDive   )  Then

        If  g_vLiftModes( nMode ).bIsEnabled  Then

            g_nSelectedLiftMode = nMode
            DisplayLiftPos        g_vLiftModes( nMode )

        End If

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Selects lift mode not disabled currently                 *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub SelectUseableLiftMode

    Dim nMode
    Dim nSelected
    Dim objMode

    nSelected   =  GetSelectedLiftMode()
    nMode       =  C_nLiftMeasModeCenter

    If  Not  g_vLiftModes( nSelected ).bIsEnabled  Then

        For Each objMode In g_vLiftModes

            If objMode.bIsEnabled  Then

                SelectLiftMode  nMode
                Exit For

            End If

            nMode = nMode + 1

        Next

    Else

        DisplayLiftPos  g_vLiftModes( nSelected )

    End If
    
End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Enables/disables diameter and tilt spin edit inputs      *
'*                                                                      *
'*  NEED:      If set to TRUE, inputs will accept user inputs           *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub EnableDiamAndTilt( bIsEnabled )

    EnableInput  spineditDiam,  bIsEnabled 

    If g_bHasTiltAxis  Then  EnableInput  spineditTilt,  bIsEnabled 

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Enables/disables usage of tilt axis                      *
'*                                                                      *
'*  NEED:      If set to TRUE, tilt axis will be accessible             *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub SetHasTiltAxis( bHasTiltAxis )

    g_bHasTiltAxis = TRUE

    If Not bHasTiltAxis  Then  DisplayTiltPos Null

    g_bHasTiltAxis = bHasTiltAxis

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Adjusts layout of corresponding HTML text container      *
'*             to spin edit                                             *
'*                                                                      *
'*  NEED:      HTML spin edit object                                    *
'*             HTML SPAN object                                         *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub SetPosDisplay( objPosSpinEdit, objPosDisplay )

    With objPosDisplay.Style

        .PixelLeft   = 0
        .PixelTop    = 0
        .PixelWidth  = objPosSpinEdit.OffsetWidth
        .PixelHeight = objPosSpinEdit.OffsetHeight

    End With

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets given value to input of lift position               *
'*             If value is invalid (NULL), spin edit input will be      *
'*             replaced by non editable text displaying home position.  *
'*                                                                      *
'*  NEED:      Lift mode object to be displayed                         *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub DisplayLiftPos( objLiftMode )

    spineditLift.Min        = objLiftMode.nMin
    spineditLift.Max        = objLiftMode.nMax
    spanLiftType.InnerText  = objLiftMode.strKey
    spineditLift.UpdateProps

    If IsNull( objLiftMode.nValue )  Then

        spineditLift.Value    = objLiftMode.nMax
        EnableInput    spineditLift,  FALSE

    Else

        spineditLift.Value    = objLiftMode.nValue
        spanLiftPos.InnerText = objLiftMode.nValue 
        EnableInput    spineditLift,  TRUE

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets given value to input of measure head's diameter.    *
'*             If value is invalid (NULL), spin edit input will be      *
'*             replaced by non editable text displaying home position.  *
'*                                                                      *
'*  NEED:      Value of measure head's diameter [mm]                    *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub DisplayDiamPos( nValue )

    If IsNull( nValue )  Then

        spineditDiam.Value    = spineditDiam.Min
        EnableInput  spineditDiam, FALSE

    Else

        spineditDiam.Value    = nValue
        spanDiamPos.InnerText = nValue 
        EnableInput  spineditDiam, TRUE

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets given value to input of measure head's tilt angle.  *
'*             If value is invalid (NULL), spin edit input will be      *
'*             replaced by non editable text displaying home position.  *
'*             Invokation will take no effect, if no tilt axis present. *
'*                                                                      *
'*  NEED:      Value of measure head's tilt angle [°]                   *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub DisplayTiltPos( nValue )

    If Not g_bHasTiltAxis Or IsNull( nValue )  Then

        spineditTilt.Value    = 0
        EnableInput  spineditTilt, FALSE

    Else

        spineditTilt.Value    = nValue
        spanTiltPos.InnerText = nValue 
        EnableInput  spineditTilt, TRUE

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets given value to input of measure head's rotational   *
'*             position angle. If value is invalid (NULL),              *
'*             spin edit input will be replaced by non editable text    *
'*             displaying home position.                                *
'*                                                                      *
'*  NEED:      Value of measure head's rotational position angle [°]    *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub DisplayRotPos

    Dim  nValue
    nValue = g_objMachineState.HeadRotation.Value

    If IsNull( nValue )  Then

        spineditRot.Value    = 0
        EnableInput  spineditRot, FALSE

    Else

        spineditRot.Value    = nValue
        spanRotPos.InnerText = nValue 
        EnableInput  spineditRot, TRUE

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets given value of outer diameter of tire to            *
'*             tire data table                                          *
'*                                                                      *
'*  NEED:      Value of tire's outer diameter [mm]                      *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub DisplayOuterDiameter( nValue )

    If IsNull( nValue ) Or IsEmpty( nValue )  Then

        tdOuterDiamMM.InnerText   = " "
        tdOuterDiamInch.InnerText = " "

    Else

        tdOuterDiamMM.InnerText   =           nValue
        tdOuterDiamInch.InnerText = MMToInch( nValue )

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets given value of inner diameter of tire to            *
'*             tire data table                                          *
'*                                                                      *
'*  NEED:      Value of tire's inner diameter [mm]                      *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub DisplayInnerDiameter( nValue )

    If IsNull( nValue ) Or IsEmpty( nValue )  Then

        tdInnerDiamMM.InnerText   = " "
        tdInnerDiamInch.InnerText = " "

    Else

        tdInnerDiamMM.InnerText   =           nValue
        tdInnerDiamInch.InnerText = MMToInch( nValue )

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets given value of tire width to tire data table        *
'*                                                                      *
'*  NEED:      Value of tire's width [mm]                               *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub DisplayTireWidth( nValue )

    If IsNull( nValue ) Or IsEmpty( nValue )  Then

        tdTireWidthMM.InnerText   = " "
        tdTireWidthInch.InnerText = " "

    Else

        tdTireWidthMM.InnerText   =           nValue
        tdTireWidthInch.InnerText = MMToInch( nValue )

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets given value of tire's misplacement to               *
'*             tire data table                                          *
'*                                                                      *
'*  NEED:      Value of tire's misplacement [mm]                        *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub DisplayExcentricity( nValue )

    If IsNull( nValue ) Or IsEmpty( nValue )  Then

        tdExcentricityMM.InnerText  = " "

    Else

        tdExcentricityMM.InnerText  = nValue

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets up all spin edit inputs concerning positioning of   *
'*             lift axis. Depending of current axis position and        *
'*             tire data, some lift axis positioning mode(s)            *
'*             will be disabled.                                        *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub SetupLiftInputs

    Dim i
    Dim vTmp
    Dim nTireWidth
    Dim bTireTooSmall
    Dim objHeadHeight

    bTireTooSmall = False   ' RW 21.9.2006


    For Each vTmp In  g_vLiftModes

        vTmp.bIsEnabled  =  FALSE
        vTmp.nValue      =  Null

    Next

    Set objHeadHeight = g_objMachineState.HeadHeight

    If  IsNull( objHeadHeight.Value )  Then

        g_vLiftModes( C_nLiftMeasModeAbsolute ).bIsEnabled = TRUE

        spineditLift.Title  = "???"

    Else

        With  g_objMachineState.Tire        ' Check, if tire's inner diameter is large enough to
                                            ' move head inside; otherwise keep head outside

            If IsNumeric( .InnerDiameter )  Then

                If    .InnerDiameter                                                _
                    >= ( 2 * .Excentricity + g_objMachineState.HeadDiameter.Value )  Then  bTireTooSmall = False

            End If

        End With

        nTireWidth  =  g_objMachineState.Tire.Width

        With g_vLiftModes( C_nLiftMeasModeCenter )

            .bIsEnabled = Not bTireTooSmall

            If IsEmpty( nTireWidth )  Then

                .nMin    =  0
                .nMax    =  0
                .nValue  =  0

            Else

                vTmp     = nTireWidth - objHeadHeight.DiveMax
                
                If vTmp < 0 Then  
                
                    vTmp =      - ( nTireWidth \ 2 )

                Else

                    vTmp = vTmp - ( nTireWidth \ 2 )

                End If

                .nMin    =  vTmp
                .nMax    =  objHeadHeight.Max - ( nTireWidth \ 2 )   '  nTireWidth \ 2 + 1
                .nValue  =  objHeadHeight.Value - nTireWidth \ 2 

            End If

        End With

        vTmp        =  objHeadHeight.Min

        With g_vLiftModes( C_nLiftMeasModeAbsolute )

            If Not IsEmpty( nTireWidth )  Then

                .bIsEnabled = TRUE

                If  bTireTooSmall  Then

                    vTmp = nTireWidth

                    If objHeadHeight.Member.ExistProperty( "SensorDist" )  Then

                        vTmp = vTmp + objHeadHeight.SensorDist

                    End If

                Else                            ' Check min head hight limited by dive capability

                    vTmp = nTireWidth - objHeadHeight.DiveMax

                    If vTmp < objHeadHeight.Min  Then  vTmp = objHeadHeight.Min

                End If

                .nMin   = vTmp
                .nMax   = objHeadHeight.Max
                .nValue = objHeadHeight.Value

            Else

                .bIsEnabled = FALSE

            End If

        End With


        With g_vLiftModes( C_nLiftMeasModeRelative )

            If Not IsEmpty( nTireWidth )  Then

                vTmp = objHeadHeight.Max - nTireWidth - 20

                If vTmp < 0  Then

                    .bIsEnabled = FALSE

                Else

                    .bIsEnabled  = TRUE
                    .nMin        = 0
                    .nMax        = vTmp

                    vTmp = objHeadHeight.Value - nTireWidth

                    If  vTmp < 0  Then vTmp = 0

                    .nValue = vTmp

                End If

            Else

                .bIsEnabled  = FALSE

            End If

        End With

        With g_vLiftModes( C_nLiftMeasModeDive )

            .bIsEnabled = Not bTireTooSmall
            .nMin       = 0
            .nMax       = objHeadHeight.DiveMax
            vTmp        = 0

            If Not IsEmpty( nTireWidth )  Then

                vTmp = nTireWidth - objHeadHeight.Value

                If vTmp < 0  Then

                    vTmp = 0

                Else

                    If vTmp > objHeadHeight.DiveMax  Then  vTmp = objHeadHeight.DiveMax

                End If

            End If

            .nValue = vTmp

        End With

        SelectUseableLiftMode

    End If

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets up spin edit input for measure head's diameter.     *
'*             If no tire data present, only minimum head diameter      *
'*             will be displayed                                        *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub SetupDiamInput

   Dim nDiamMin
   Dim nDiamMax
   Dim objPos

    With g_objMachineState.HeadDiameter

        spineditDiam.Min  = .Min
        spineditDiam.Max  = .Max ' We have to initialize here, before call to GetSelectedPositions

        If IsEmpty( g_objMachineState.Tire.InnerDiameter ) Then

            DisplayDiamPos      Null

        Else

            Set objPos = GetSelectedPositions()

            If Not IsEmpty( objPos.Height )  Then

            If g_bPLCBC640 Then

               nDiamMin   = g_objMachine.EvalheadParam( "HeadDiamMin", objPos )
            
            Else
            
               nDiamMin = .Min

            End If

            nDiamMax   = g_objMachine.EvalHeadParam( "HeadDiamMax", objPos )

         End If

         If Not IsNull( nDiamMax ) And Not IsEmpty( nDiamMax ) And _
            Not IsNull( nDiamMin ) And Not IsEmpty( nDiamMin) Then

            spineditDiam.Min  = nDiamMin
                spineditDiam.Max  = nDiamMax

                DisplayDiamPos .Value

                EnableInput  spineditDiam, TRUE

            Else

                spineditDiam.Max   = .Min
                spineditDiam.Value = .Min

                EnableInput  spineditDiam, FALSE

            End If

        End If

        spineditDiam.UpdateProps

    End With

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets up spin edit input for measure head's rotational    *
'*             angle position                                           *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub SetupRotInput

    With g_objMachineState.HeadRotation

        If .Member.ExistProperty( "Positions" ) Then  
        
            spineditRot.Values   = .Positions
            spineditRot.ReadOnly = TRUE

        Else

            spineditRot.Min      = .Min
            spineditRot.Max      = .Max

        End If

        spineditRot.UpdateProps
        DisplayRotPos


    End With

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Sets up spin edit input for measure head's diameter.     *
'*             If no tire data present, only a head tilt angle of zero  *
'*             will be displayed                                        *
'*             Invokation will take no effect, if no tilt axis present. *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub SetupTiltInput

    If Not g_bHasTiltAxis  Then  Exit Sub

    Dim bInputEnabled 
    bInputEnabled = False

    With g_objMachineState.HeadTilt0

        Dim nTiltMin
        Dim nTiltMax

        Dim objPos
        Set objPos = GetSelectedPositions()

        If Not IsEmpty( objPos.Height )  Then

            'nTiltMin   = g_objMachine.EvalHeadParam( "HeadTiltMin", objPos )
            'nTiltMax   = g_objMachine.EvalHeadParam( "HeadTiltMax", objPos )
				nTiltMin		= g_objMachine.TiltMin
				nTiltMax		= g_objMachine.TiltMax

            If Not IsNull( nTiltMin ) And Not IsEmpty( nTiltMin )  And _
               Not IsNull( nTiltMax ) And Not IsEmpty( nTiltMax ) Then

                spineditTilt.Min  = nTiltMin
                spineditTilt.Max  = nTiltMax

                bInputEnabled = TRUE

            Else

                spineditTilt.Min  = 0
                spineditTilt.Max  = 0

                bInputEnabled = FALSE

            End If

        Else

            spineditTilt.Min  = .Min
            spineditTilt.Max  = .Max

            bInputEnabled = TRUE

        End If


        spineditTilt.UpdateProps

        If IsEmpty( g_objMachineState.Tire.InnerDiameter ) Then

            DisplayTiltPos      Null

        Else

            DisplayTiltPos      .Value
            EnableInput  spineditTilt, bInputEnabled

        End If

    End With

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Returns an object similar to 'GetHeadPosition'           *
'*             containing measure head's positions currently selected   *
'*             by user. Additionally two properties will be added.      *
'*             'HeightBase'  will be set to the selected lift axis      *
'*             positioning mode as defined in                           *
'*             'C_nLiftMeasModeXXXXX' constants                         *
'*             'HeightValue' will be set to spin edit's value           *
'*             corresponding to the selected lift positioning mode.     *
'*                                                                      *
'*  NEED:      ---                                                      *
'*                                                                      *
'*  RETURN:    DynamicObject containing measure head's position as      *
'*             selected by user                                         *
'*                                                                      *
'************************************************************************

Function GetSelectedPositions

    Set GetSelectedPositions = g_objMachine.GetHeadPosition

    With GetSelectedPositions.Member
        
        .Documentation( "Member" ) = "Contains destination head position"

        .AddProperty  "HeightValue",,,0
        .AddProperty  "HeightBase",,,0

    End With

    GetSelectedPositions.Diameter = spineditDiam.Value
    GetSelectedPositions.Rotation = spineditRot.Value

    If g_bHasTiltAxis  Then  

        GetSelectedPositions.Tilt = spineditTilt.Value

    End If

    Dim nTmp
    Dim nLim
    Dim nTireWidth
    nTireWidth  =  g_objMachineState.Tire.Width

    GetSelectedPositions.HeightBase = GetSelectedLiftMode

    Select Case  GetSelectedPositions.HeightBase

        Case C_nLiftMeasModeCenter
                    GetSelectedPositions.HeightValue  = spineditLift.Value

                    If Not IsEmpty( nTireWidth )  Then

                        GetSelectedPositions.Height = ( nTireWidth \ 2 ) + GetSelectedPositions.HeightValue

                    Else

                        GetSelectedPositions.Member.GetFullAccess.Item( "Height" ) = Empty

                    End If

        Case C_nLiftMeasModeAbsolute
                    GetSelectedPositions.Height       = spineditLift.Value
                    GetSelectedPositions.HeightValue  = spineditLift.Value

        Case C_nLiftMeasModeRelative
                    GetSelectedPositions.HeightValue  = spineditLift.Value

                    If Not IsEmpty( nTireWidth )  Then

                        nLim = g_objMachineState.HeadHeight.Max
                        nTmp = nTireWidth + GetSelectedPositions.HeightValue

                        If nTmp > nLim Then

                            GetSelectedPositions.Height = nLim

                        Else

                            GetSelectedPositions.Height = nTmp

                        End If

                    Else

                        GetSelectedPositions.Height = Empty

                    End If


        Case C_nLiftMeasModeDive
                    GetSelectedPositions.HeightValue  = spineditLift.Value

                    If Not IsEmpty( nTireWidth )  Then

                        nLim = g_objMachineState.HeadHeight.Min
                        nTmp = nTireWidth - GetSelectedPositions.HeightValue

                        If nTmp < nLim Then

                            GetSelectedPositions.Height = nLim

                        Else

                            GetSelectedPositions.Height = nTmp

                        End If

                    Else

                        GetSelectedPositions.Member.GetFullAccess.Item( "Height" ) = Empty

                    End If

    End Select

End Function

'************************************************************************
'*                                                                      *
'*  TASK:                                                               *
'*                                                                      *
'*  NEED:      Object containing positions (like returned by            *
'*                'GetSelectedPositions')                               *
'*             TRUE, if position spin edit inputs should be updated to  *
'*                new positions                                         *
'*             TRUE, if message box should inform user on errors        *
'*                                                                      *
'*  RETURN:    TRUE on success                                          *
'*                                                                      *
'*  EFFECTS:   Measure head may gain new position                       *
'*                                                                      *
'************************************************************************

Function PositionTo(  objPositions, bHasToUpdatePosInputs, bHasToShowErrors  )

    PositionTo = TRUE

    Dim vTmp
    Dim nRef
    Dim nDist
    Dim bNeedsUpdate
    Dim bHasToUpdateRot
    Dim bHasToUpdateLift
    Dim bHasToUpdateDiam
    Dim bHasToUpdateTilt
    Dim objCurPos
    Set objCurPos = g_objMachine.GetHeadPosition

    bHasToUpdateRot     = FALSE
    bHasToUpdateLift    = FALSE
    bHasToUpdateDiam    = FALSE
    bHasToUpdateTilt    = FALSE

    g_objMachine.ClearErrorText

    If objPositions.Rotation <> Round(objCurPos.Rotation)  Then

        ShowAnimated  tdRot, TRUE, "#00FF00", "#008000", 500

        PositionTo = g_objMachine.RotateHead(  objPositions.Rotation  )

        StopAnimation  tdRot

        bHasToUpdateRot = TRUE

        If bHasToShowErrors And Not PositionTo  Then  g_objMachine.DisplayErrorText

    End If

    ' HeadDiameter
    If PositionTo  And PositionToDiameterFirst(objPositions) Then

        Set objCurPos = g_objMachine.GetHeadPosition
        If objPositions.Diameter <> Round(objCurPos.Diameter)  Then

            ShowAnimated  tdDiam, TRUE, "#00FF00", "#008000", 500

            PositionTo = g_objMachine.HeadDiameter(  objPositions.Diameter  )

            StopAnimation  tdDiam

            bHasToUpdateDiam = TRUE

            If bHasToShowErrors And Not PositionTo  Then  g_objMachine.DisplayErrorText

        End If

    End If
    
    If PositionTo  Then

        If IsNull ( objPositions.Height ) And IsNull( objCurPos.Height ) Then
            bNeedsUpdate = False
        Else
            bNeedsUpdate = Not CBool( objPositions.Height = Round(objCurPos.Height) )
        End If

        If bNeedsUpdate  Then

            If g_bHasTiltAxis Then
            If objCurPos.Tilt <> 0 Then PositionTo = g_objMachine.TiltHead(0)
            End If
            
            If PositionTo Then
        
                nDist = objPositions.HeightValue
                nRef  = 0

                Select Case  objPositions.HeightBase

                    Case C_nLiftMeasModeAbsolute

                            nRef        = 2
                            If g_bPLCBC640 Then
                                nDist       = nDist - g_objMachine.State.Tire.Width * 0.5
                            End If

                    Case C_nLiftMeasModeCenter

                            nRef        = 2
                            If Not g_bPLCBC640 Then
                                nDist       = nDist - g_objMachine.State.Tire.Width * 0.5
                            End If

                    Case C_nLiftMeasModeRelative
     
                            nRef        =  1
                            PositionTo  =  Abs( PositionTo )

                    Case C_nLiftMeasModeDive

                            nRef        =  1
                            PositionTo  = -Abs( PositionTo )

                End Select

                ShowAnimated  tdLift, TRUE, "#00FF00", "#008000", 500
                

                If g_bPLCBC640 Then
                PositionTo   =   g_objMachine.LiftHead(  nDist, nRef + C_LiftTypeKeepPosFlag  )
               
                Else
                
                    Select Case  objPositions.HeightBase

                        Case C_nLiftMeasModeAbsolute

                                PositionTo   =   g_objMachine.LiftHead( nDist )      

                        Case C_nLiftMeasModeCenter

                                PositionTo   =   g_objMachine.LiftHeadToTire(  -nDist )      

                        Case C_nLiftMeasModeRelative
         
                                nDist  =  Abs( nDist )
                                PositionTo   =   g_objMachine.LiftHeadToTire(  -nDist )     

                        Case C_nLiftMeasModeDive

                                nDist  = -Abs( nDist )
                                PositionTo   =   g_objMachine.LiftHeadToTire(  -nDist )   

                    End Select
                End If
                                

                StopAnimation  tdLift

                bHasToUpdateLift = TRUE
                bHasToUpdateDiam = TRUE
                bHasToUpdateTilt = TRUE

            End If
            
            If bHasToShowErrors And Not PositionTo  Then  g_objMachine.DisplayErrorText

            If PositionTo  Then  DisplayTireWidth  g_objMachineState.Tire.Width

        End If

    End If

    ' HeadDiameter
    If PositionTo  Then

        Set objCurPos = g_objMachine.GetHeadPosition

        If objPositions.Diameter <> Round(objCurPos.Diameter)  Then

            ShowAnimated  tdDiam, TRUE, "#00FF00", "#008000", 500

            PositionTo = g_objMachine.HeadDiameter(  objPositions.Diameter  )

            StopAnimation  tdDiam

            bHasToUpdateDiam = TRUE

            If bHasToShowErrors And Not PositionTo  Then  g_objMachine.DisplayErrorText

        End If

    End If

    If PositionTo  And  g_bHasTiltAxis  Then

        If objPositions.Tilt <> Round(objCurPos.Tilt)  Then

            Dim nTilt
            nTilt = objPositions.Tilt

            ShowAnimated  tdTilt, TRUE, "#00FF00", "#008000", 2000

            Dim objPos
            Set objPos = g_objMachine.GetHeadPosition

            If Not IsEmpty( objPos.Height )  Then
        
                Dim nTiltMin
                Dim nTiltMax

                nTiltMin   = g_objMachine.EvalHeadParam( "HeadTiltMin", objPos )
                nTiltMax   = g_objMachine.EvalHeadParam( "HeadTiltMax", objPos )


                If Not IsNull( nTiltMin ) And Not IsEmpty( nTiltMin )  And _
                   Not IsNull( nTiltMax ) And Not IsEmpty( nTiltMax ) Then
                
                   If nTiltMin > nTilt Then nTilt = nTiltMin
                   If nTiltMax < nTilt Then nTilt = nTiltMax
                    
                Else
                	
                   nTilt = 0
                  
                End If

            End If

            PositionTo = g_objMachine.TiltHead(  nTilt  )

            StopAnimation  tdTilt

            bHasToUpdateTilt = TRUE

            If bHasToShowErrors And Not PositionTo  Then  g_objMachine.DisplayErrorText

        End If

    End If

    If bHasToUpdatePosInputs  Then  

        If bHasToUpdateRot   Then  DisplayRotPos
        If bHasToUpdateLift  Then  SetupLiftInputs
        If bHasToUpdateDiam  Then  SetupDiamInput

        If g_bHasTiltAxis  And  bHasToUpdateTilt  Then  SetupTiltInput

    End If

End Function

'************************************************************************

Function PositionToDiameterFirst( objPositions )

    PositionToDiameterFirst = False

    Dim objCurPos 
    Set objCurPos = g_objMachine.GetHeadPosition
    
    If objPositions.Diameter > g_objMachineState.Tire.OuterDiameter Then PositionToDiameterFirst = True

'  FIXME Check HeadHight    

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Starts color animation for given HTML element.           *
'*             Color will be vary between two given values within       *
'*             given time interval back and forth. Animation will       *
'*             be repeated until 'StopAnimation' is invoked.            *
'*             Only one animation at a time is allowed                  *
'*                                                                      *
'*  NEED:      HTML element to be animated                              *
'*             TRUE, if element's color should be animated; FALSE, if   *
'*                element's background color should be animated;        *
'*             Color 1 (format #RRGGBB)                                 *
'*             Color 2 (format #RRGGBB)                                 *
'*             Duration of single animation [ms]                        *
'*                                                                      *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub ShowAnimated( objElem, bIsForegroundCol, strCol1, strCol2, nDur )

    Set g_nAnimObject   = objElem
    g_bAnimIsFGColor    = bIsForegroundCol
    g_nAnimCol1         = strCol1
    g_nAnimCol2         = strCol2
    g_nAnimTickerCount  = 0

    g_nAnimTickerCookie = MySetInterval(  GetRef( "AnimationTicker"), nDur \ 20 )

    g_nAnimObject. RuntimeStyle.Color = strCol2

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Stops animation of given HTML element                    *
'*                                                                      *
'*  NEED:      HTML element an animation was started over by            *
'*                invoking 'ShowAnimated'                               *
'*                                                                      *
'*  RETURN:    ---                                                      *
'*                                                                      *
'************************************************************************

Sub StopAnimation( objElem )

    MyClearInterval g_nAnimTickerCookie

    objElem.RuntimeStyle.Color           = ""
    objElem.RuntimeStyle.BackgroundColor = ""
    g_nAnimTickerCookie                  = 0
    
End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Converting from millimeter to inches                     *
'*                                                                      *
'*  NEED:      Millimeter value                                         *
'*                                                                      *
'*  RETURN:    Inch value                                               *
'*                                                                      *
'************************************************************************

Function MMToInch( nValue )

    If Not IsNumeric( nValue )  Then Exit Function

    Dim n
    Dim nAbs
    Dim nInt

    n    = Round( CDbl( nValue ) / 25.4, 1 )
    nAbs = Abs( n    )
    nInt = Int( nAbs )

    If ( nAbs - Int( nAbs )) >= 0.5 Then nInt = nInt + 0.5

    MMToInch = nInt
    If n < 0 Then MMToInch = -nInt

End Function

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

'********************************************************************************
'*                                                                              *
'*  H A N D L E R   F O R   T I M E R   T I C K S                               *
'*                                                                              *
'********************************************************************************
'*                                                                              *
'* As you cannot reliably use multiple timers with IE6,                         *
'* we use one timer with schedules all events                                   *
'*                                                                              *
'********************************************************************************

const cMaxTimer = 10
Dim g_TimerCode(10)
Dim g_TimerInterval(10)

Dim g_nTimerTicker      ' Count Number of Ticks
g_nTimerTicker = 0

'********************************************************************************

Sub TimerTicker()

    g_nTimerTicker = g_nTimerTicker + 1

    Dim vCode
    Dim nInterval
    Dim i

    For i = 1 To cMaxTimer

        nInterval = g_TimerInterval(i)

        If nInterval > 0 Then

            Set vCode = g_TimerCode(i)
            If g_nTimerTicker mod nInterval = 0 Then
                Call vCode
            End If

        End If

    Next
        
End Sub

'********************************************************************************

Function MySetInterval( vCode, iMilliSeconds )


    Dim i
    For i = 1 To cMaxTimer

        If g_TimerInterval(i) = 0 Then
         
            Set g_TimerCode (i) = vCode
            g_TimerInterval(i) = iMilliSeconds / 10
            MySetInterval = i
            Exit Function
        End If
    Next


    If g_nTimer > cMaxTimer Then
        MsgBox "Too many timers ( >" & cMaxTimer & ")"
        MySetInterval = 0
        Exit Function
    End If


End Function

'********************************************************************************

Sub MyClearInterval( nID )

    If (nID < 1) Or (nID > cMaxTimer) Then 
        MsgBox "MyClearInterval: invalid ID " & nID
        Exit Sub
    End If

    g_TimerCode( nID ) = 0
    g_TimerInterval( nID ) = 0

End Sub

'********************************************************************************
'*                                                                              *
'*  H A N D L E R   F O R   T I M E R   T I C K S                               *
'*                                                                              *
'********************************************************************************

Sub MachineTicker()

    g_objMachine.IsEmergencyActive

End Sub

'********************************************************************************
'*                                                                              *
'*  H A N D L E R   F O R   T I M E R   T I C K S   O F   A N I M A T I O N     *
'*                                                                              *
'********************************************************************************

Sub AnimationTicker()

    Dim strCol

    If g_nAnimTickerCount <= 10  Then

        strCol = g_objHTMLHelper.MixColor( g_nAnimCol1, g_nAnimCol2, g_nAnimTickerCount / 10.0 )

    ElseIf g_nAnimTickerCount <= 20  Then

        strCol = g_objHTMLHelper.MixColor( g_nAnimCol1, g_nAnimCol2, ( 20 - g_nAnimTickerCount ) / 10.0 )

    Else

        strCol              = g_nAnimCol2
        g_nAnimTickerCount  = -1    

    End If

    If g_bAnimIsFGColor  Then

        g_nAnimObject.RuntimeStyle.Color           = strCol

    Else

        g_nAnimObject.RuntimeStyle.BackgroundColor = strCol

    End If

    g_nAnimTickerCount  = g_nAnimTickerCount + 1

End Sub

'********************************************************************************
'********************************************************************************
