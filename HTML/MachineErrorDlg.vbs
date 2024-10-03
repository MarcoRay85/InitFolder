'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   M a c h i n e E r r o r D l g             *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************

Option Explicit

Dim g_objWarnDIV
Dim g_strMessages
Dim g_nDlgWidth
Dim g_nDlgHeight
Dim g_nSpaceXY
Dim g_nPIScale
Dim g_Application

g_nPIScale = 0.8 * Atn( 1 )

Dim g_nTickCount

Dim g_strVML

g_strVML =  "<v:group coordorigin=""0,0"" CoordSize=""100 100"" "                             & _
            "         Style=""position:relative;width:100; height:100;margin:2,8""> "         & _
            "  <v:shape Style=""position:relative; left:0; top:0; width:100; height:100""> "  & _
            "    <v:fill type=""solid"" color=""yellow"" /> "                                 & _
            "    <v:stroke weight=""3pt"" color=""black"" linestyle=""single""/> "            & _
            "    <v:path v=""M15 85 "                                                         & _
                            "L85 85 "                                                         & _
                            "C95 85 95 85 90 75 "                                             & _
                            "L55 15 "                                                         & _
                            "C50 7 50 7 45 15 "                                               & _
                            "L10 75 "                                                         & _
                            "C5 85 5 85 15 85""/> "                                           & _
            "  </v:shape> "                                                                   & _
            "  <v:shape Style=""position:relative; left:0; top:0; width:100; height:100""> "  & _
            "    <v:fill type=""solid"" color=""black"" /> "                                  & _
            "    <v:path v=""M45 30 "                                                         & _
                            "C40 30 40 50 45 58 "                                             & _
                            "C50 65 50 65 55 58 "                                             & _
                            "C60 51 60 30 55 30 "                                             & _
                            "C50 27 50 27 45 30"" /> "                                        & _
            "  </v:shape> "                                                                   & _
            "  <v:oval fillcolor=""#000000"" "                                                & _
            "          Style=""position:relative; left:45; top:68; width:12; height:12""/> "  & _
            "</v:group>"

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim vArg

    If Document.Body.IsDialog Then

        varg = Window.DialogArguments
        Set g_Application = varg(1)
        g_strMessages = Split( varg(0), vbNewLine )

    Else

        g_strMessages = Array( "[Description 1]" & vbTab & "Test Error 1",   _
                               "[Description 2]" & vbTab & "Test Error 2",   _
                               "[Description 3]" & vbTab & "Test Error 3"    )

    End If

    With Document.Body.WindowHandle

        g_nDlgWidth   = .Width
        g_nDlgHeight  = .Height

    End With

    Document.Body.InsertAdjacentHTML                                                      _
      "afterBegin",                                                                       _
      "<div align=""center"" "                                                          & _
      "     Style=""position:absolute; left:0; top:0; width:100%; height:100%; "        & _
      "             z-Index:-1;filter:alpha(opacity=10) ""> "                           & _
      "</div>"

    Set g_objWarnDIV = Document.Body.FirstChild

    g_objWarnDIV.InnerHTML = g_strVML & g_strVML & g_strVML & g_strVML & g_strVML _
                           & g_strVML & g_strVML & g_strVML & g_strVML & g_strVML

    Dim nI
    Dim nPos
    Dim bMoreFlag
    Dim strLine
    Dim strOut

    nI        = 1
    bMoreFlag = FALSE

    For Each strLine In g_strMessages

        nPos = InStr( strLine, "]" & vbTab )

        If nPos > 0 Then strLine = Mid( strLine, nPos + 2 ) : bMoreFlag = TRUE

        strOut = strOut & strLine & vbNewLine
        nI     = nI + 1

    Next

    With preDisplay.Style

        .Position     = "absolute"
        .Left         = 0
        .Top          = 0
        .FontFamily   = "Arial, Swiss"
        .FontSize     = "16pt"
        .FontWeight   = "Bolder"

    End With

    preDisplay.InnerText = Left( strOut, Len( strOut ) - 1 )

    Dim nSpaceX
    Dim nSpaceY
    Dim nWidth
    Dim nHeight
    Dim objDummy
    Set objDummy = preDisplay.GetBoundingClientRect

    nWidth      = objDummy.Right  - objDummy.Left
    nHeight     = objDummy.Bottom - objDummy.Top
    g_nSpaceXY  = nHeight \ nI

    nSpaceX     = document.body.clientwidth     _
                - g_nSpaceXY

    nSpaceY     = STANDARDDLG_BUTTONS.OffsetTop _
                - g_nSpaceXY

    IncSize  nWidth - nSpaceX, nHeight - nSpaceY

    nSpaceX     = document.body.clientwidth     _
                - g_nSpaceXY

    nSpaceY     = STANDARDDLG_BUTTONS.OffsetTop _
                - g_nSpaceXY

    With preDisplay.Style

        .Left = ( nSpaceX - nWidth  ) \ 2
        .Top  = ( nSpaceY - nHeight ) \ 2

    End With

    g_nTickCount          = 0
    window.SetInterval  "onFlash", 100, "VBScript"

    STANDARDDLG_BUTTONS.CloseDlg  = FALSE
    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )

    If Not g_Application.Machine.Member.ExistMethod( "ShowPDF" ) Then
        Dim objButton
        Set objButton = STANDARDDLG_BUTTONS.GetButton( 2 )
        If Not objButton is Nothing Then objButton.style.visibility = "hidden"
    End If

    If bMoreFlag = FALSE Then STANDARDDLG_BUTTONS.GetButton( 1 ).Disabled = TRUE

End Sub

'********************************************************************************

Sub IncSize( nIncX, nIncY )


    With Document.Body

        If nIncX > 0 Then

            g_nDlgWidth =  g_nDlgWidth + nIncX
            If .IsDialog Then Window.DialogWidth =  g_nDlgWidth & "px"

        End If

        If nIncY > 0 Then

            g_nDlgHeight =  g_nDlgHeight + nIncY
            If .IsDialog Then Window.DialogHeight =  g_nDlgHeight & "px"

        End If

        If Not .IsDialog Then Window.ResizeTo g_nDlgWidth, g_nDlgHeight

        .WindowHandle.Center

        STANDARDDLG_BUTTONS.Style.Left = (    .ClientWidth                    _
                                            - STANDARDDLG_BUTTONS.OffsetWidth _
                                         ) \ 2

    End With

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************

Sub STANDARDDLG_BUTTONS_onClicked()

    If STANDARDDLG_BUTTONS.Value <> 258 And STANDARDDLG_BUTTONS.Value <> 261 Then Window.Close : Exit Sub

    If  STANDARDDLG_BUTTONS.Value = 258 Then STANDARDDLG_BUTTONS.GetButton( 1 ).Disabled = TRUE

    Window.Event.CancelBubble = TRUE

    If  STANDARDDLG_BUTTONS.Value = 258 Then ' MORE

        Dim objDispDetails
        Dim nPos
        Dim strLine
        Dim strOut

        For Each strLine In g_strMessages

            nPos = InStr( strLine, "]" & vbTab )

            If nPos > 0 Then

                strLine = Left( strLine, nPos )

            Else

                If Len( strLine ) > 0  Then strLine = ". . ."

            End If

            strOut = strOut & strLine & vbNewLine

        Next

        Set objDispDetails = preDisplay.CloneNode( FALSE )
        objDispDetails.ID  = "preDisplayDetails"

        With objDispDetails.Style

            .Left         = "auto"
            .PixelRight   = g_nSpaceXY

        End With

        preDisplay.style.PixelLeft = g_nSpaceXY

        preDisplay.InsertAdjacentElement "afterEnd",  objDispDetails
        objDispDetails.InnerText = Left( strOut, Len( strOut ) - 1 )

        Dim objDummy
        Set objDummy = preDisplay.GetBoundingClientRect
        nPos         = objDummy.Right - objDummy.Left       _
                     - Document.Body.ClientWidth

        Set objDummy = objDispDetails.GetBoundingClientRect
        nPos         = objDummy.Right - objDummy.Left       _
                     + nPos

        IncSize  nPos + 3 * g_nSpaceXY, 0

    ElseIf STANDARDDLG_BUTTONS.Value = 261 Then ' HELP

        strLine = g_strMessages(0)

        nPos = InStr( strLine, "]" & vbTab )

        If nPos > 0 Then strLine = Mid( strLine, nPos + 2 )

        g_Application.Machine.ShowPDF( strLine )

    Else

        MsgBox  "Internal error: Unknown Button was pressed!",,   _
                "Intact testing"
    End If

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   F L A S H I N G   E X C L A M A T I O N          *
'*                                                                              *
'********************************************************************************

Sub onFlash()

    g_nTickCount = g_nTickCount + 1

    g_objWarnDIV.Filters.Alpha.Opacity =  Round( 5.5 * sin( g_nTickCount * g_nPIScale )) + 7

End Sub

'********************************************************************************
'********************************************************************************
