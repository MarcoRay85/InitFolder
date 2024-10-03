'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   T e s t F e e d T i r e D l g             *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************
Option Explicit
                                                'Text resources used in code
Dim g_nTimeCounter
Dim g_objMachine
Set g_objMachine = Nothing

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    With pLoadText.Style

        .FontFamily = "Arial"
        .FontSize   = "125%"
        .Position   = "relative"
        .Width      = "100%"
        .Height     = "100%"
        .Top        = "1em"
        .Filter     = "progid:DXImageTransform.Microsoft.Matrix( M22=1.75 )"
        .BackgroundColor = "ThreeDFace"
        .ZIndex     = -1

    End With

    If Document.Body.IsDialog Then

        If TypeName( window.DialogArguments ) <> "DynamicObject" Then

            MsgBox "Dialog expects a Application (dynamic-) object!"

            window.Close
            Exit Sub

        End If

        Set g_objMachine = window.dialogArguments
        Set g_objMachine = g_objMachine.Machine

        g_objMachine.PrepareNextTire

    End If

    window.SetInterval "MonitorTask", 50, "VBScript"

    g_nTimeCounter = 0

End Sub

'********************************************************************************

Sub MonitorTask()

    Dim vTmp
    vTmp = g_nTimeCounter Mod 20

    If vTmp => 10 Then

        pLoadText.RuntimeStyle.Color = "#" & Hex(  55 + 10 * vTmp ) & "0000"

    Else

        pLoadText.RuntimeStyle.Color = "#" & Hex( 255 - 10 * vTmp ) & "0000"

    End If

    g_nTimeCounter = g_nTimeCounter + 1

    If ( g_nTimeCounter Mod 10 ) <> 0 Then Exit Sub

    If g_objMachine Is Nothing   Then

        vTmp = FALSE
        Window.Status = "L " &g_nTimeCounter

    Else

        vTmp = g_objMachine.HasTireToLoad()

    End If

    If IsNull( vTmp ) Then window.Close : Exit Sub

    If vTmp = TRUE    Then window.Close : Exit Sub

    If ( g_nTimeCounter Mod 200 ) <> 0 Then Exit Sub

    g_nTimeCounter = 0

    If ( g_objMachine Is Nothing )  Then

        vTmp = FALSE
        Window.Status = "P " &g_nTimeCounter

    Else

        g_objMachine.PrepareNextTire

    End If

End Sub

'********************************************************************************
'********************************************************************************
