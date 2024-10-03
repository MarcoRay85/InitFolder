'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   S e t T e s t P r e s s u r e D l g       *
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

    ' inputVacuumStart.Precision = 0
    inputVacuum.Precision = 0

    If Not Document.Body.IsDialog Then

        Randomize

        ' inputVacuumStart.Min      = 0
        ' inputVacuumStart.Max      = 80
        ' inputVacuumStart.Value    = Round( 80 * Rnd )
        
        inputVacuum.Min      = 0
        inputVacuum.Max      = 80
        inputVacuum.Value    = Round( 80 * Rnd )

    Else

        With window.dialogArguments

            ' inputVacuumStart.Min      = .Item( "Min"       )
            ' inputVacuumStart.Max      = .Item( "Max"       )
            ' inputVacuumStart.Value    = .Item( "ValueStart")
            
            inputVacuum.Min      = .Item( "Min"       )
            inputVacuum.Max      = .Item( "Max"       )
            inputVacuum.Value    = .Item( "Value"     )

        End With

    End If

    ' inputVacuumStart.UpdateLook
    inputVacuum.UpdateLook

    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )

    inputVacuum.Focus

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
    bIsOK = inputVacuum.Validate

    STANDARDDLG_BUTTONS.CloseDlg = bIsOk

    If Document.Body.IsDialog And bIsOK Then

        ' Window.DialogArguments.Item( "PressureStart" ) = inputVacuumStart.Value
        Window.DialogArguments.Item( "Pressure" ) = inputVacuum.Value

    End If

End Sub


'********************************************************************************
'********************************************************************************
