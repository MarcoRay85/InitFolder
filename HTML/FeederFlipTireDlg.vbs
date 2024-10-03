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
Dim nTimeCounter
Dim nState
Dim nType
Dim objApp
Dim objHdl
Set objApp = Nothing
Set objHdl = Nothing

nState     = 1

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim vArgs

    If Not Document.Body.IsDialog Then

        If ( Second( Now ) Mod 2 ) > 0 Then

            pLiftText.Style.display = "block"

        Else

            pLoadText.Style.display = "block"

        End If

        Exit Sub

    End If


    vArgs         = Window.dialogArguments

    Set objApp    = vArgs( 0 )
    Set objHdl    = vArgs( 1 )
    nType         = vArgs( 2 )

    If nType = 1 Then

        pLiftText.Style.display = "block"

    Else

        pLoadText.Style.display = "block"

    End If

    Dim objWinHandle
    Set objWinHandle   = CreateObject( "ScriptingToolsSO.WindowHandle" )

    objWinHandle.ID    = Document
    Set objWinHandle   = objWinHandle.Parent

    If  UBound( vArgs ) > 2  Then

        objWinHandle.Title = vArgs( 3 )

    Else
    
        objWinHandle.Title = objWinHandle.Title & " (" & nType & ")"

    End If

    spanButton.Style.cssText = "color:green;font-size:120%"
    spanFlip.Style.cssText   = "color:green;font-size:120%"

    Window.SetInterval "MonitorTask", 500, "VBScript"
    Window.ReturnValue = FALSE


    nTimeCounter = 0

End Sub

'********************************************************************************

Sub MonitorTask()

    Dim objAnim
    Dim objInfo

    Set objInfo = objHdl()
    If    IsNull( objInfo.IsDropped     )  _
       Or IsNull( objInfo.IsHorizontal  )  _
       Or IsNull( objInfo.IsLifted      )  Then Window.Close

    If nType = 1 Then

        Set objAnim = spanButton
        If  objInfo.IsHorizontal  Or  objInfo.IsLifted  Then
        
            Window.ReturnValue = TRUE
            Window.Close

        End If

    Else

        Select Case nState

            Case 1  Set objAnim = spanFlip
                    If  objInfo.IsDropped    Then

                        nState = 2
                        spanFlip.Style.cssText   = "auto"
                        spanLoad.Style.cssText   = "color:green;font-size:120%"

                    End If

            Case 2  Set objAnim = spanLoad
                    If  objInfo.IsHorizontal  Or  objInfo.IsLifted  Then

                        Window.ReturnValue = TRUE
                        Window.Close

                    End If

        End Select

    End If

    With objAnim.style

        If UCase( .color ) = "GREEN" Then

            .color = "springgreen"

        Else

            .color = "green"

        End If

    End With

    nTimeCounter = nTimeCounter + 1

    If nTimeCounter < 20 Then Exit Sub

    nTimeCounter = 0

End Sub

'********************************************************************************
'********************************************************************************
