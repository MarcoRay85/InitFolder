'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   M a c h i n e T e s t D l g               *
'*                                                                              *
'********************************************************************************


'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************
Option Explicit

Dim g_nTickCount
Dim g_nTickCookie

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

g_nTickCount          = 0
g_nTickCookie         = 0

'********************************************************************************

Sub onInitialize        ' Let window be visible if started as HTML page

    With Document.Body

        With .WindowHandle

            .Show                0
            .ModifyStyle         &H08000000, &H00080000
            .ModifyExtendedStyle &H00000008, 0
            .Center

        End With

        If Not .IsDialog Then SetDisplayMode TRUE

        .AttachEvent  "onmousemove", GetRef( "onMouseMoveOverWindow" )

        With .Style

            .Border = "0.5ex outset red"

        End With

    End With

End Sub

'********************************************************************************
'*                                                                              *
'*       M E T H O D    F O R   S E T T I N G   D I S P L A Y   S T A T E       *
'*                                                                              *
'********************************************************************************

Sub SetDisplayMode( bIsVisible )

    With Document.Body.WindowHandle

        If IsNull( bIsVisible ) Then

            Window.Close
            Exit Sub

        End If


        If bIsVisible Then

            .Show 5

            On Error Resume Next

                .ToForeground

            On Error Goto 0

            If g_nTickCookie = 0 Then g_nTickCookie = Window.SetInterval( "KeepOnTop()", 50, "VBScript" )

        Else

            .Show   0

            If g_nTickCookie <> 0 Then

                window.ClearInterval  g_nTickCookie
                g_nTickCookie       = 0
                g_nTickCount        = 0

            End If

        End If

    End With

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   M E S S A G E   F L A S H I N G                  *
'*                                                                              *
'********************************************************************************

Sub KeepOnTop

    Dim strColor

    With Document.Body

        With .WindowHandle

            If ( g_nTickCount Mod 10 ) = 0 Then

                Dim objTopWin
                Set objTopWin = .ForegroundWindow

                If objTopWin.Exists Then

                    If .ProcessID = objTopWin.ProcessID Then

                        On Error Resume Next

                            .ToForeground

                        On Error Goto 0

                    End If

                End If

            End If

        End With


        With .RuntimeStyle

            Select Case UCase(   .BackgroundColor )

                Case "#FF0000"   strColor = "F8"
                Case "#F80000"   strColor = "F0"
                Case "#F00000"   strColor = "E8"
                Case "#E80000"   strColor = "E0"
                Case "#E00000"   strColor = "D8"
                Case "#D80000"   strColor = "D0"
                Case "#D00000"   strColor = "C8"
                Case "#C80000"   strColor = "C0"
                Case "#C00000"   strColor = "B8"
                Case "#B80000"   strColor = "B0"
                Case "#B00000"   strColor = "A8"
                Case "#A80000"   strColor = "A0"
                Case "#A00000"   strColor = "A4"
                Case "#A40000"   strColor = "AC"
                Case "#AC0000"   strColor = "B4"
                Case "#B40000"   strColor = "BC"
                Case "#BC0000"   strColor = "C4"
                Case "#C40000"   strColor = "CC"
                Case "#CC0000"   strColor = "D4"
                Case "#D40000"   strColor = "DC"
                Case "#DC0000"   strColor = "E4"
                Case "#E40000"   strColor = "EC"
                Case "#EC0000"   strColor = "F4"
                Case "#F40000"   strColor = "FC"
                Case "#FC0000"   strColor = "FF"
                Case Else        strColor = "FF"

            End Select

            .Color           = "#" & strColor & strColor & "00"
            .BorderColor     = .Color
            .BackgroundColor = "#" & strColor & "0000"

        End With

        '.FirstChild.RuntimeStyle.Color = "#" & strColor & strColor & "00"

    End With

    g_nTickCount = g_nTickCount + 1

    If g_nTickCount > 20 Then g_nTickCount = 0

End Sub

'********************************************************************************
'*                                                                              *
'*             H A N D L E R    D R A G G I N G    W I N D O W                  *
'*                                                                              *
'********************************************************************************

Sub onMouseMoveOverWindow

    Dim objEvent
    Set objEvent = Window.Event

    If  ( objEvent.Button And 1 ) <> 0  Then

        With Document.Body.WindowHandle

            .Left = objEvent.ScreenX  - .Width  \ 2
            .Top  = objEvent.ScreenY  - .Height \ 2

        End With

    End If

End Sub

'********************************************************************************
'********************************************************************************
