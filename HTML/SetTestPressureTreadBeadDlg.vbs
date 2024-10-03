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
    inputVacuumTread.Precision = 0
	inputVacuumBead.Precision = 0

    If Not Document.Body.IsDialog Then

        Randomize

        ' inputVacuumStart.Min      = 0
        ' inputVacuumStart.Max      = 80
        ' inputVacuumStart.Value    = Round( 80 * Rnd )
        
        inputVacuumTread.Min      = 0
        inputVacuumTread.Max      = 80
        inputVacuumTread.Value    = Round( 80 * Rnd )

		inputVacuumBead.Min      = 0
        inputVacuumBead.Max      = 80
        inputVacuumBead.Value    = Round( 80 * Rnd )
		

    Else

        With window.dialogArguments

            ' inputVacuumStart.Min      = .Item( "Min"       )
            ' inputVacuumStart.Max      = .Item( "Max"       )
            ' inputVacuumStart.Value    = .Item( "ValueStart")
            
            inputVacuumTread.Min      = .Item( "Min"       )
            inputVacuumTread.Max      = .Item( "Max"       )
            inputVacuumTread.Value    = .Item( "ValueTread")

			inputVacuumBead.Min      = .Item( "Min"       )
            inputVacuumBead.Max      = .Item( "Max"       )
            inputVacuumBead.Value    = .Item( "ValueBead" )
		

        End With

    End If

    ' inputVacuumStart.UpdateLook
    inputVacuumTread.UpdateLook
	inputVacuumBead.UpdateLook
	
    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )

    inputVacuumTread.Focus

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
	bIsOK = False
    'bIsOK = inputVacuum.Validate
	If (inputVacuumTread.Validate = True AND inputVacuumBead.Validate = True) Then
		bIsOK = True
	End If

    STANDARDDLG_BUTTONS.CloseDlg = bIsOk

    If Document.Body.IsDialog And bIsOK Then

        ' Window.DialogArguments.Item( "PressureStart" ) = inputVacuumStart.Value
        Window.DialogArguments.Item( "PressureTread" ) = inputVacuumTread.Value
		Window.DialogArguments.Item( "PressureBead" ) = inputVacuumBead.Value
    End If

End Sub


'********************************************************************************
'********************************************************************************
