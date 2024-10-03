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

Dim  g_bIsWithUnload

g_bIsWithUnload  =  FALSE

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim  nAdd
    Dim  nHeight

    With spanFlash.Style

        .FontSize        = "110%"
        .Color           = "Yellow"
        .height          = "1.9em"
        .BackgroundColor = "Red"
        .Border          = "3px solid yellow"

    End With


    On Error Resume Next

        With window.dialogArguments

            g_bIsWithUnload  =  window.dialogArguments.Exists( "ClearChamber" )

        End With

    On Error Goto 0

    If  g_bIsWithUnload  Then

        nAdd     =  pContinue.OffsetTop
        nHeight  =  Window.DialogHeight
        nHeight  =  CInt( Left( nHeight, Instr( nHeight, "px") - 1 )) 

        'inputcheckClearChamber.checked = window.dialogArguments.Item( "ClearChamber" )
        'pClearQuery.RuntimeStyle.Display  =  "block"

        Window.DialogHeight  =  (  nHeight  +  pContinue.OffsetTop  -  nAdd   )  & "px"

    End If

    window.SetInterval "Flash()", 40, "VBScript"

End Sub

'********************************************************************************
  
Sub inputcheckClearChamber_onclick

     If  g_bIsWithUnload  Then
 
         window.dialogArguments.Item( "ClearChamber" ) = inputcheckClearChamber.checked

     End If

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   M E S S A G E   F L A S H I N G                  *
'*                                                                              *
'********************************************************************************

Sub Flash()

    Dim strColor

    With spanFlash.RuntimeStyle

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

End Sub

'********************************************************************************
'********************************************************************************
