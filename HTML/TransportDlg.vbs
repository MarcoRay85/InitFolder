'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   T r a n s p o r t D l g                   *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************

Option Explicit

Dim  g_nMin
Dim  g_nMax
Dim  g_nValue

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim  vTmp

    inputDistance.Precision = 0

    If Not Document.Body.IsDialog Then

        Randomize

        g_nMin        = -2000
        g_nMax        =  1000
        g_nValue      = g_nMin + Round(( g_nMax - g_nMin )* Rnd )

    Else

        With window.dialogArguments

            g_nMin    = .Item( "Min"       )
            g_nMax    = .Item( "Max"       )
            g_nValue  = .Item( "Value"     )

            If .Exists( "HighSpeed" ) Then

                If .Item( "HighSpeed" ) Then

                    inputSpeedFast.Checked = TRUE

                Else

                    inputSpeedSlow.Checked = TRUE

                End If

            End If

            If .Exists( "Reversed" ) Then

                If .Item( "Reversed" ) Then

                    inputDirReverse.Checked = TRUE

                Else

                    inputDirForward.Checked = TRUE

                End If

            End If

        End With

    End If

    If  g_nMin <> -g_nMax  Then

        If  g_nMax <= 0  Then

            vTmp      =   g_nMax
            g_nMax    =  -g_nMin
            g_nMin    =  -vTmp
            g_nValue  =   Abs(  g_nValue  )

            inputDirReverse.Checked = TRUE

        Else

            inputDirForward.Checked = TRUE

        End If

        inputDirForward.ParentElement.ParentElement.Disabled  =  TRUE

    Else

        g_nMin    =  0
        g_nValue  =  Abs(  g_nValue  )

    End If

    inputDistance.Min    =  g_nMin
    inputDistance.Max    =  g_nMax
    inputDistance.Value  =  g_nValue
    

    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )

    inputDistance.Focus

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************

Sub STANDARDDLG_BUTTONS_onClicked()

    If STANDARDDLG_BUTTONS.Value <> 1 Then

        Window.Close
        Exit Sub

    End If

    Dim bIsOK
    bIsOK = inputDistance.Validate

    STANDARDDLG_BUTTONS.CloseDlg = bIsOk

    If Document.Body.IsDialog And bIsOK Then

        With window.dialogArguments

            .Item(  "Distance"  )  =        inputDistance.value
            .Item(  "Fast"      )  =  CInt( inputSpeedFast.checked  )
            .Item(  "Forward"   )  =  CInt( inputDirForward.checked )

        End With

    End If

End Sub


'********************************************************************************
'********************************************************************************
