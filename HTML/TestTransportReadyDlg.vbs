'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   T e s t T r a n s p o r t R e a d y D l g *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************
Option Explicit

Dim g_nTimeCounter
Dim g_objMachine
Dim g_vParam1
Dim g_vParam2
Set g_objMachine = Nothing

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    With spanStatus.Style

        .FontFamily  = "Arial"
        .FontSize    = "125%"
        .Position    = "relative"
        .SetExpression "Left",                              _
                       "(  spanMain.ClientWidth"            _
                    &  " - spanStatus.OffsetWidth ) \ 2",   _
                       "VBScript"
        .SetExpression "top",                               _
                       "(  spanMain.ClientHeight"           _
                     & " - spanStatus.OffsetHeight ) \ 2",  _
                       "VBScript"

    End With

    spanStatus.RuntimeStyle.Top = "1ex"

    With spanReason.Style

        .Overflow       = "auto"
        .Position       = "absolute"
        .PixelWidth     = spanMain.ClientWidth
        .PixelHeight    = spanMain.ClientHeight         _
                        - spanStatus.OffsetHeight       _
                        - spanStatus.OffsetTop     * 2
        .Top            = "auto"
        .PixelLeft      = 0
        .PixelBottom    = 0
        .Padding        = "0.5ex"
        .Border         = "1px solid"
        .BorderColor    = "buttonshadow buttonhighlight buttonhighlight buttonshadow"

    End With

    spanStatus.RuntimeStyle.Top = ""

    spanReason.RuntimeStyle.Display = "none"
    spanReason.InnerHTML            = "<PRE Style= ""Position:Relative;"                    _
                                    & "Font:normal normal 200 75% normal Arial""></PRE>"


    If Document.Body.IsDialog Then

        If Not IsArray( window.DialogArguments ) Then

            MsgBox  Document.URL    _
                  & vbNewLine       _
                  & "expects an ARRAY( Application (dynamic-) object, Param1, Param2 )!"

            window.Close
            Exit Sub

        End If

        Dim vTmp
        vTmp        = window.DialogArguments
        Set g_objMachine = vTmp( 0 ).Machine
        g_vParam2        = vTmp( 2 )
        g_vParam1        = vTmp( 1 )

    End If

    window.SetInterval "MonitorTask", 50, "VBScript"

    g_nTimeCounter = 0

End Sub

'********************************************************************************

Sub MonitorTask()

    Dim i
    Dim vTmp
    vTmp = g_nTimeCounter Mod 20

    If vTmp => 10 Then

        spanStatus.RuntimeStyle.Color = "#" & Hex(  55 + 10 * vTmp ) & "0000"

    Else

        spanStatus.RuntimeStyle.Color = "#" & Hex( 255 - 10 * vTmp ) & "0000"

    End If

    g_nTimeCounter = g_nTimeCounter + 1

    If ( g_nTimeCounter Mod 10 ) <> 0 Then Exit Sub

    If g_objMachine Is Nothing   Then                   ' Simulate messages

        If ( g_nTimeCounter Mod 20 ) = 0 Then

            vTmp = ""

            For i = 1 To g_nTimeCounter / 20

                vTmp = vTmp & vbNewLine & String( i, "*" )

            Next

            vTmp = "Cycle count: " & g_nTimeCounter _
                 & vbNewLine                        _
                 & String( g_nTimeCounter, "*" )    _
                 & vTmp

        Else

            vTmp = ""

        End If

    Else

        vTmp = g_objMachine.IsReadyToTransport( g_vParam1, g_vParam2 )

    End If

    If VarType( vTmp ) = vbBoolean Then If vTmp Then window.Close : Exit Sub

    If Len( vTmp ) > 0 Then

        spanStatus.RuntimeStyle.Top     = "1ex"
        spanReason.RuntimeStyle.Display = ""

        With  spanReason.FirstChild

            .InnerText = vTmp

            If .OffsetHeight < spanReason.ClientHeight Then

                .RuntimeStyle.PixelTop = ( spanReason.ClientHeight - .OffsetHeight ) \ 2

            Else

                .RuntimeStyle.PixelTop = 0

            End If

        End With

    Else

        spanStatus.RuntimeStyle.Top     = ""
        spanReason.RuntimeStyle.Display = "none"
        spanReason.FirstChild.InnerText = ""

    End If

    If ( g_nTimeCounter Mod 200 ) = 0 Then g_nTimeCounter = 0

End Sub

'********************************************************************************
'********************************************************************************


