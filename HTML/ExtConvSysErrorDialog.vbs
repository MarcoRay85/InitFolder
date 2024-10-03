Option explicit

'**********************************************************************************************************************

Dim G_objEventSink

'**********************************************************************************************************************


Sub onInitialize

   With Document.Body

      With .WindowHandle

         .ModifyStyle  0, &H00080000
         .InstallMessageFilter GetRef( "MsgHdl" ), &H0010, &H0010

      End With

   End With
   
   
   Set G_objEventSink = window.dialogArguments.Item( "objEvtSink" )

   G_objEventSink.SetInit

End Sub

'**********************************************************************************************************************

Sub MsgHdl( objMessage )

    objMessage.Msg = 0

End Sub

'**********************************************************************************************************************

Sub btOk_onClick
   
   G_objEventSink.SetOK
   
End Sub

'**********************************************************************************************************************

Sub SetDisplayMode( bIsVisible )

	With Document.Body.WindowHandle

		If IsNull( bIsVisible ) Then

			Document.Body.WindowHandle.InstallMessageFilter Nothing, 0, 0
			Window.Close
			Exit Sub

		End If	
		
		G_objEventSink.SetInit

		If bIsVisible Then

			.Show 5

			On Error Resume Next

				.ToForeground

			On Error Goto 0

		Else

			.Show   0

		End If

	End With

End Sub

'**********************************************************************************************************************

Sub SetErrorText ( nModuleNr, nModuleType, nErrNrSeg1, nErrNrSeg2 )
	
	Dim strOutput

	If nModuleType = 2 Then		'nModuleType = 2 is a long module
		
		 If nErrNrSeg1 <> 0 And nErrNrSeg2 = 0 Then
		
			 strOutput = GetStringRes( "Module" ) & " " & nModuleNr & " ( " & GetStringRes( CStr( "ModuleType" & nModuleType ) ) _
			 & ", " & GetStringRes( "ModuleLongSeg1" ) & " ): " & GetStringRes( CStr( "ExtConvSysError" & nErrNrSeg1 ) )
				
		 ElseIf nErrNrSeg1 = 0 And nErrNrSeg2 <> 0 Then
		
			 strOutput = GetStringRes( "Module" ) & " " & nModuleNr & " ( " & GetStringRes( CStr( "ModuleType" & nModuleType ) ) _
			 & ", " & GetStringRes( "ModuleLongSeg2" ) & " ): " & GetStringRes( CStr( "ExtConvSysError" & nErrNrSeg2 ) )
			
		 Else
		
			 strOutput = _
			 GetStringRes( "Module" ) & " " & nModuleNr & " ( " & GetStringRes( CStr( "ModuleType" & nModuleType ) ) & " )" _
			 & vbNewLine & vbNewLine _
			 & GetStringRes( "ModuleLongSeg1" ) & ": " & GetStringRes( CStr( "ExtConvSysError" & nErrNrSeg1 ) ) & vbNewLine _
			 & GetStringRes( "ModuleLongSeg2" ) & ": " & GetStringRes( CStr( "ExtConvSysError" & nErrNrSeg2 ) )
		
		 End If
		
	 Else
		
		 strOutput = GetStringRes( "Module" ) & " " & nModuleNr & " ( " & GetStringRes( CStr( "ModuleType" & nModuleType ) ) _
		 & " ): " & GetStringRes( CStr( "ExtConvSysError" & nErrNrSeg1 ) )
		
	 End If
    
    Document.GetElementByID( "ExtConvSysErrorDialog_Output" ).InnerText = strOutput

End Sub

'**********************************************************************************************************************

Sub SetText( strMsg )
	
	Document.GetElementByID( "ExtConvSysErrorDialog_Output" ).InnerText = GetStringRes( CStr( strMsg ) )

End Sub

'**********************************************************************************************************************

Function GetStringRes( ID )
	Dim objNode

	Set objNode = XMLStringTable.DocumentElement.SelectSingleNode("String[@ID=""" & ID & """]")

	If Not objNode Is Nothing Then

		GetStringRes = objNode.Text
	  
	Else   

		GetStringRes = "THIS RESSOURCE DOESN'T EXIST" 
	  
	End If

End Function