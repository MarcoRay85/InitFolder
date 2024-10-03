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

    If Document.Body.IsDialog Then
        
        inpUserName.Value = window.dialogArguments.Item("username")
        
        STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )

        inpOldPassword.Focus()
    End If

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************

Sub STANDARDDLG_BUTTONS_onClicked()

    If STANDARDDLG_BUTTONS.Value = 1 Then

        window.dialogArguments.Add "oldpassword", inpOldPassword.Value
        window.dialogArguments.Add "newpassword1", inpNewPassword1.Value
        window.dialogArguments.Add "newpassword2", inpNewPassword2.Value
    Else

        Window.Close
        Exit Sub

    End If

End Sub


'********************************************************************************
'********************************************************************************
