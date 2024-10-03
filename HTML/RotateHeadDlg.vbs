'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   R o t a t e H e a d D l g                 *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************
Option Explicit


'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim vTmp
    Dim bHasToHome
    bHasToHome = FALSE

    InputRotDegree.Precision = 0

    If Not Document.Body.IsDialog Then

        Randomize

        If Rnd > 0.5 Then

            InputRotDegree.Min      = 0
            InputRotDegree.Max      = 180
            InputRotDegree.Value    = Round( 180 * Rnd )

        Else

            vTmp                    = Array( 0, 40, 60, 80 )
            InputRotDegree.Values   = vTmp
            InputRotDegree.Value    = vTmp( Round( 3 * Rnd ))
            InputRotDegree.ReadOnly = TRUE


        End If

    Else

        With window.dialogArguments

            If .Exists( "Positions" ) Then

                InputRotDegree.Values   = .Item( "Positions" )
                InputRotDegree.ReadOnly = TRUE

            Else

                InputRotDegree.Min      = .Item( "Min"       )
                InputRotDegree.Max      = .Item( "Max"       )

                If .Item( "Min" ) = .Item( "Max" ) Then  InputRotDegree.ReadOnly = TRUE

            End If

            InputRotDegree.Value        = .Item( "Value"     )

        End With

    End If

    InputRotDegree.UpdateLook

    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )

    If Not bHasToHome Then InputRotDegree.Focus

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
    bIsOK = InputRotDegree.Validate

    STANDARDDLG_BUTTONS.CloseDlg = bIsOk

    If Document.Body.IsDialog And bIsOK Then

        Window.DialogArguments.Item( "Degree" ) = InputRotDegree.value

    End If

End Sub

'********************************************************************************
'********************************************************************************
