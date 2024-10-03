'********************************************************************************
'*                                                                              *
'*    S C R I P T   F I L E    F O R    R o t a t e f l i p p e r D l g         *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************

Option Explicit
                                                'Text resources used in code
Dim g_objTermBlock
Dim g_objApplication
Dim g_bIsClampOpen
Dim g_bIsRetainerOpen

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim vTmp
	Dim nToResize

	nToResize = 0
	
    Set  g_objTermBlock  =  Nothing

    If  Document.Body.IsDialog  Then

        vTmp                    =  Window.DialogArguments
        Set g_objTermBlock      =  vTmp( 0 )
        Set g_objApplication    =  vTmp( 1 )

		If  Not ( g_objTermBlock Is Nothing )  Then
	
			g_objTermBlock.Call  "Flip1200_SetManualMode", TRUE
			
		End If
		
        buttonRotLeft.onmousedown       = GetRef( "RotateLeft"    )
        buttonRotLeft.onmouseup         = GetRef( "RotateStop"    )
        buttonRotLeft.onmouseout        = GetRef( "RotateStop"    )

        buttonRotRight.onmousedown      = GetRef( "RotateRight"   )
        buttonRotRight.onmouseup        = GetRef( "RotateStop"    )
        buttonRotRight.onmouseout       = GetRef( "RotateStop"    )

		buttonTransLoad.onmousedown     = GetRef( "TransportLoading" )
		buttonTransLoad.onmouseup       = GetRef( "TransportStop"    )
		buttonTransLoad.onmouseout      = GetRef( "TransportStop"    )
		
		buttonTransUnload.onmousedown   = GetRef( "TransportUnloading" )
		buttonTransUnload.onmouseup     = GetRef( "TransportStop"      )
		buttonTransUnload.onmouseout    = GetRef( "TransportStop"      )
		
        STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )

        g_bIsClampOpen  =  g_objTermBlock.Call(  "Flip1200_IsClampOpen"  )

    ElseIf ( Second( Now ) Mod 2 ) > 0 Then

        g_bIsClampOpen  =  TRUE

    End If

    If  Not IsEmpty(  g_bIsClampOpen  )  Then

		nToResize = nToResize + CInt( buttonRotLeft.OffsetHeight ) * 1.25
	
        buttonClamp.ParentElement.Style.Display = "block"
        UpdateClampButton

    End If

    If  CBool( g_objTermBlock.Call(  "Flip1200_HasAdaptiveRetainer"  ))  Then

	   	nToResize = nToResize + CInt( buttonRotLeft.OffsetHeight ) * 1.25

        UpdateRetainerButton

        buttonRetainer.ParentElement.Style.Display = "block"

    End If

	
	If CBool( g_objTermBlock.Call(  "Flip1200_HasPoweredRolls"  )) Then
	
		nToResize = nToResize + CInt( buttonRotLeft.OffsetHeight ) * 1.25
	
		UpdateTransportButtons
		
		buttonTransLoad.ParentElement.Style.Display = "block"
		buttonTransUnload.ParentElement.Style.Display = "block"
	
	End If

	If nToResize > 0 Then
	
		If  Document.Body.IsDialog  Then

			 Window.DialogHeight  =  Document.Body.WindowHandle.Height      _
								  +  CInt(  nToResize  )    				_
								  &  "px"

		Else

			Document.ParentWindow.ResizeBy  0, CInt(  nToResize  )

		End If

	End If
		
End Sub

'********************************************************************************

Sub STANDARDDLG_BUTTONS_onClicked()

    If  Not ( g_objTermBlock Is Nothing )  Then
	
        g_objTermBlock.Call  "Flip1200_SetManualMode", FALSE
        g_objTermBlock.Call  "Flip1200_ManRotate", NULL

    End If

End Sub

'********************************************************************************

Sub RotateLeft

    If  Not ( g_objTermBlock Is Nothing )  Then

        g_objTermBlock.Call  "Flip1200_ManRotate", 1
 
    End If	
	
End Sub

'********************************************************************************

Sub RotateRight

    If  Not ( g_objTermBlock Is Nothing )  Then

        g_objTermBlock.Call  "Flip1200_ManRotate", -1

    End If

End Sub

'********************************************************************************

Sub RotateStop

    If  Not ( g_objTermBlock Is Nothing )  Then

        g_objTermBlock.Call  "Flip1200_ManRotate", 0

    End If

	UpdateTransportButtons
	
End Sub

'********************************************************************************

Sub TransportLoading

	g_objTermBlock.Call  "Flip1200_StartPoweredRolls", TRUE
	
End Sub

'********************************************************************************

Sub TransportUnloading

	g_objTermBlock.Call  "Flip1200_StartPoweredRolls", FALSE

End Sub

'********************************************************************************

Sub TransportStop

	g_objTermBlock.Call  "Flip1200_StopPoweredRolls"

End Sub

'********************************************************************************

Sub UpdateTransportButtons

	If g_objTermBlock.Call( "Flip1200_IsHome" ) Then
		buttonTransLoad.InnerHTML = "&larr;["
		buttonTransUnload.InnerHTML = "&rarr;["		
		buttonTransLoad.Disabled = FALSE
		buttonTransUnload.Disabled = FALSE
	ElseIf g_objTermBlock.Call( "Flip1200_IsAt180Deg" ) Then
		buttonTransLoad.InnerHTML = "&rarr;["
		buttonTransUnload.InnerHTML = "&larr;["		
		buttonTransLoad.Disabled = FALSE
		buttonTransUnload.Disabled = FALSE
	Else
		buttonTransLoad.Disabled = TRUE
		buttonTransUnload.Disabled = TRUE
		buttonTransLoad.InnerHTML = "&harr;["
		buttonTransUnload.InnerHTML = "&harr;["		
	End If
	
End Sub

'********************************************************************************

Sub UpdateClampButton

    If  Not ( g_objTermBlock Is Nothing )  Then

        g_bIsClampOpen  =  g_objTermBlock.Call(  "Flip1200_IsClampOpen"  )

    End If

    If  IsNull( g_bIsClampOpen )  Then

        buttonClamp.InnerHTML = "|&nbsp;???&nbsp;|"

    ElseIf  g_bIsClampOpen  Then

        buttonClamp.InnerHTML = "|&rarr;&nbsp;&nbsp;&nbsp;&larr;|"

    Else

        buttonClamp.InnerHTML = "|&larr;&rarr;|"

    End If

End Sub

 '********************************************************************************

Sub UpdateRetainerButton

    If  Not ( g_objTermBlock Is Nothing )  Then

        g_bIsRetainerOpen  =  g_objTermBlock.Call(  "Flip1200_IsRetainerOpen"  )

    End If

    If  IsNull( g_bIsRetainerOpen )  Then

        buttonRetainer.InnerHTML = "|&nbsp;???&nbsp;|"

    ElseIf  g_bIsRetainerOpen  Then

        buttonRetainer.InnerHTML = "|&nbsp;=&nbsp;|"

    Else

        buttonRetainer.InnerHTML = "|&nbsp;<&nbsp;|"

    End If

End Sub

'********************************************************************************

Sub buttonClamp_onclick

    If  Not ( g_objTermBlock Is Nothing )  Then

        If  IsNull( g_bIsClampOpen )  Then

            g_objTermBlock.Call  "Flip1200_SetClampOpen",  FALSE

        Else

            g_objTermBlock.Call  "Flip1200_SetClampOpen",  Not  g_bIsClampOpen

        End If

    Else

        g_bIsClampOpen  =  Not  CBool(  g_bIsClampOpen  )

    End If

    UpdateClampButton

End Sub

'********************************************************************************

Sub buttonRetainer_onclick


    If  Not ( g_objTermBlock Is Nothing )  Then

 
        g_bIsRetainerOpen  =  g_objTermBlock.Call(  "Flip1200_IsRetainerOpen"  )


        If  g_bIsRetainerOpen Then

            g_objTermBlock.Call  "Flip1200_AdjustRetainer"

        Else

            g_objTermBlock.Call  "Flip1200_ReleaseRetainer"

        End If

    Else

        g_bIsRetainerOpen  =  Not  CBool(  g_bIsRetainerOpen  )

    End If

    UpdateRetainerButton

End Sub

'********************************************************************************
'********************************************************************************
