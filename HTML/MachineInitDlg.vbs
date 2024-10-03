'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   M a c h i n e T e s t D l g               *
'*                                                                              *
'********************************************************************************

Option Explicit

const WM_CLOSE = &H0010

Dim g_objBtnHandler
Set g_objBtnHandler = Nothing


Set  Document.onKeyDown  =  GetRef( "onKeyDown" )


'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

	With Document.Body
		
		If  .IsDialog  Then  Set g_objBtnHandler = Window.DialogArguments

        If  Not ( g_objBtnHandler Is Nothing )  Then  '*** Visible dialog or hidden log generating window?

            With .WindowHandle

                .ModifyStyle  0, &H00080000
                .UpdateLook
                .Center
'                .InstallMessageFilter GetRef( "MsgHdl" ), WM_CLOSE, WM_CLOSE

            End With

        Else

            .WindowHandle.ToBottom

        End If

    End With

    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************

Sub STANDARDDLG_BUTTONS_onClicked()

    If Not ( g_objBtnHandler Is Nothing ) Then

        g_objBtnHandler  STANDARDDLG_BUTTONS.Value

    Else

        CloseDialog
        Exit Sub

    End If

End Sub

'********************************************************************************
'*                                                                              *
'*    F U N C T I O N    F O R    C L O S I N G    D I A L O G                  *
'*                                                                              *
'********************************************************************************

Sub CloseDialog

    If Not ( g_objBtnHandler Is Nothing ) Then

        On Error Resume Next

            With  Document.Body.WindowHandle

 '               .InstallMessageFilter Nothing, 0, 0
                .Show 0


            End With
        On Error Goto 0

        Set g_objBtnHandler = Nothing

        Window.SetTimeout  "Window.Close", 2000, "VBScript"

    Else

        Window.Close

    End If
    
End Sub

'********************************************************************************
'*                                                                              *
'*   S U P P R E S S I N G    C L O S E    M E S S A G E S                      *
'*                                                                              *
'********************************************************************************

Sub  onKeyDown

    With  Window.Event

        If     ( .KeyCode = 116 )                   _
           Or (( .KeyCode = 115 ) And .AltKey )     Then
           
            .CancelBubble   =  TRUE     '***  Suppress original action F5: reload
            .ReturnValue    =  FALSE    '***                       ALT+F4: close
            .KeyCode        =  0
            
        End If
        
    End With
    
End Sub


'********************************************************************************
'********************************************************************************

