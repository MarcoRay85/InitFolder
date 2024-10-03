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
Dim g_nCloseTime
Dim g_nStartTime

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

g_nTickCount          = 0
g_nTickCookie         = 0
g_nStartTime          = -1
g_nCloseTime          = -1
    
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

    g_nStartTime = -1
    g_nCloseTime = -1

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


Sub SetMessage(strMsg, strTitle, nTime)

    Document.Title  = strTitle
    divTitleText.InnerText = strTitle
    divMessageText.InnerText = strMsg
    g_nStartTime = Timer()
    g_nCloseTime = nTime
    
End Sub


'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   M E S S A G E   F L A S H I N G                  *
'*                                                                              *
'********************************************************************************

Sub KeepOnTop

    Dim strColor

    With Document.Body

        If True Then

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

        End If
        

        With .RuntimeStyle

            Select Case UCase(   .BackgroundColor )

                Case "#0060FF"   strColor = "F8"
                Case "#0060F8"   strColor = "F0"
                Case "#0060F0"   strColor = "E8"
                Case "#0060E8"   strColor = "E0"
                Case "#0060E0"   strColor = "D8"
                Case "#0060D8"   strColor = "D0"
                Case "#0060D0"   strColor = "C8"
                Case "#0060C8"   strColor = "C0"
                Case "#0060C0"   strColor = "B8"
                Case "#0060B8"   strColor = "B0"
                Case "#0060B0"   strColor = "A8"
                Case "#0060A8"   strColor = "A0"
                Case "#0060A0"   strColor = "A4"
                Case "#0060A4"   strColor = "AC"
                Case "#0060AC"   strColor = "B4"
                Case "#0060B4"   strColor = "BC"
                Case "#0060BC"   strColor = "C4"
                Case "#0060C4"   strColor = "CC"
                Case "#0060CC"   strColor = "D4"
                Case "#0060D4"   strColor = "DC"
                Case "#0060DC"   strColor = "E4"
                Case "#0060E4"   strColor = "EC"
                Case "#0060EC"   strColor = "F4"
                Case "#0060F4"   strColor = "FC"
                Case "#0060FC"   strColor = "FF"
                Case Else        strColor = "FF"

            End Select

            .Color           = "#" & strColor & strColor & strColor
            .BorderColor     = .Color
            .BackgroundColor = "#" & "0060" & strColor

        End With

        '.FirstChild.RuntimeStyle.Color = "#" & strColor & strColor & "00"

    End With

    g_nTickCount = g_nTickCount + 1

    If g_nTickCount > 20 Then g_nTickCount = 0
    
    If g_nCloseTime >= 0 Then
    
        'divMessageText.InnerText = "Timer() - g_nStartTime  = " & (Timer() - g_nStartTime) & ", g_nCloseTime = " & g_nCloseTime
    
        If (Timer() - g_nStartTime) > g_nCloseTime Then
            
            'MsgBox "Countdown beendet"
            
            SetDisplayMode FALSE
            
        End If
    
    Else
    
        'divMessageText.InnerText = "g_nCloseTime < 0"
    
    End If

End Sub

'********************************************************************************
'********************************************************************************


