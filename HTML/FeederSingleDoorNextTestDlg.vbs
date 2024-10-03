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

Dim nTimeCounter
Dim objApp
Dim objHdl
Set objApp = Nothing
Set objHdl = Nothing

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

   Dim vArgs

   If Not Document.Body.IsDialog Then

      If ( Second( Now ) Mod 2 ) > 0 Then

         pTiltUpText.Style.display = "block"

      Else

         pFlipText.Style.display = "block"

      End If

      Exit Sub

   End If

   vArgs         = Window.dialogArguments

   Set objApp    = vArgs( 0 )
   Set objHdl    = vArgs( 1 )

    pTiltUpText.Style.display = "block"

   Dim objWinHandle
   Set objWinHandle   = CreateObject( "ScriptingToolsSO.WindowHandle" )

   objWinHandle.ID    = Document
   Set objWinHandle   = objWinHandle.Parent

   If  UBound( vArgs ) > 2  Then

      objWinHandle.Title = vArgs( 3 )

   Else
    
      objWinHandle.Title = objWinHandle.Title

   End If

   Window.SetInterval "MonitorTask", 500, "VBScript"
   Window.ReturnValue = FALSE

   nTimeCounter = 0

End Sub

'********************************************************************************

Sub MonitorTask()

   Dim objInfo

   Set objInfo = objHdl()

   If IsNull( objInfo.ExtState.PositionDown ) Then Window.Close
     
   If objInfo.ExtState.PositionDown Then            

      Window.ReturnValue = TRUE
      Window.Close

   End If

   nTimeCounter = nTimeCounter + 1

   If nTimeCounter < 20 Then Exit Sub

   nTimeCounter = 0

End Sub

'********************************************************************************
'********************************************************************************
