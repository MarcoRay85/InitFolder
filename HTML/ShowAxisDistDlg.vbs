'********************************************************************************
'*                                                                              *
'*        S C R I P T   F I L E    F O R  S h o w H e a d D i s t D l g         *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************
option explicit

	Dim objShellServices, objFSO, objWShell              
	Set objShellServices = CreateObject( "ScriptingToolsSO.ShellServices" )
	Set objFSO           = CreateObject( "Scripting.FileSystemObject"     )
	Set objWShell        = CreateObject( "WScript.Shell"                  )

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

sub window_onload

	
    tdLiftMinMM.InnerText   = window.dialogArguments.Item( "LiftMin" )
    tdLiftMinINCH.InnerText = fctMMToInch( window.dialogArguments.Item( "LiftMin" ))

    tdLiftMaxMM.InnerText   = window.dialogArguments.Item( "LiftMax" )
    tdLiftMaxINCH.InnerText = fctMMToInch( window.dialogArguments.Item( "LiftMax" ))


    tdShiftMinMM.InnerText   = window.dialogArguments.Item( "SHiftMin" )
    tdShiftMinINCH.InnerText = fctMMToInch( window.dialogArguments.Item( "SHiftMin" ))

    tdShiftMaxMM.InnerText   = window.dialogArguments.Item( "SHiftMax" )
    tdShiftMaxINCH.InnerText = fctMMToInch( window.dialogArguments.Item( "SHiftMax" ))

    if Not window.dialogArguments.Item( "bExistPneuRot"  ) then
    	  IdTrPneuRot.Style.visibility = "hidden"
	     tdRotMinMM.InnerText   = window.dialogArguments.Item( "RotMin" ) & "°"
	     tdRotMaxMM.InnerText   = window.dialogArguments.Item( "RotMax" ) & "°"
    Else
	     IdTrRot.Style.visibility = "hidden"
	     IdTrPneuRot.Style.visibility = "visible"
	     Dim nDegree, vDegrees
	     vDegrees        = window.dialogArguments.Item( "RotPositions" )
	     Dim strPneuRot
	     strPneuRot = ""

	  For Each nDegree in vDegrees
		
		strPneuRot = strPneuRot & CStr(nDegree) & "°, "
	                   
        Next
        Dim nLengthStr
        nLengthStr = len(strPneuRot)
        
        strPneuRot = left( strPneuRot, nLengthStr-2 )
	tdPneuRot.InnerText =	strPneuRot
    End IF

    if window.dialogArguments.Item( "bExistTiltAxis"  ) then
    	tdTiltMinMM.InnerText   = window.dialogArguments.Item( "TiltMin" ) & "°"
    	tdTiltMaxMM.InnerText   = window.dialogArguments.Item( "TiltMax" ) & "°"
    Else
           IdTrTilt.Style.visibility = "hidden"
    End IF

    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )

end sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************

sub STANDARDDLG_BUTTONS_onClicked()
	
    if STANDARDDLG_BUTTONS.Value = 1 Then
	
    ElseIF STANDARDDLG_BUTTONS.Value = 256 Then
    	if not fctPrintHTML() then 
   	    msgbox("Print failed!")
   	end if
    Else
    	msgbox("STANDARDDLG_BUTTONS.Value: " & STANDARDDLG_BUTTONS.Value)
    	
    End if
    window.close
end sub

'***********************************************************************
'****                                                               ****
'****       Converting from millimeter to inches                    ****
'****                                                               ****
'***********************************************************************

function fctMMToInch( nValue )

    Dim n
    Dim nAbs
    Dim nInt

    n    = Round( CDbl( nValue ) / 25.4, 1 )
    nAbs = Abs( n    )
    nInt = Int( nAbs )
    
    if ( nAbs - Int( nAbs )) >= 0.5 then nInt = nInt + 0.5

    fctMMToInch = nInt
    if n < 0 then fctMMToInch = -nInt

end function

'***********************************************************************
'****                                                               ****
'****       			   Print String                             ****
'****                                                               ****
'***********************************************************************

Function fctPrintHTML()

	fctPrintHTML = False

   window.print

	fctPrintHTML = True

   Exit Function

	Do 
	   Dim strMain, strTemplate
		strMain = objFSO.GetAbsolutePathName("")
		strMain = objFSO.BuildPath(strMain, "Html")
        	strTemplate = cTemplate
        	strMain  = strMain + strTemplate

		Dim objTestDocPage    
		Set objTestDocPage = ShowModelessDialog( strMain )
		
		objTestDocPage.idTimeDate.InnerText = Date & "	" & Time
		
		Dim objTestDoc     
		Set objTestDoc     = objTestDocPage.document  
		
		Dim nPageheight, bPortrait
		bPortrait = True
		If bPortrait = True then
			nPageheight = 24.8
		Else
			nPageheight = 16.5
		End If

		Dim strToPrint, strData, strHeader, strFooter  
	
		strHeader = objTestDocPage.idHEADER.InnerHTML
		strData = idDataTable.InnerHtml
		strFooter = objTestDocPage.idFooter.InnerHTML
		strToPrint = "<HTML>"&_
				"<Body>"  &_
				strHeader &_
				"<div style='height:" & nPageheight & "cm'>" &_
				strData   &_
				"</div>"&_
				strFooter &_
				"<Body>"  &_
				"</HTML>" 

		Dim objfile
		Set objfile = objFSO.createtextfile("e:\testFile.html")
		objfile.WriteLine( strToPrint )
        	objfile.Close

		window.ShowModalDialog "PrintTemplate.HTML", Array( strToPrint, bPortrait )


    Dim objPrintDriver
    Set objPrintDriver          = CreateObject( "IEHTMLToolsSO.TemplatedPrinting" )
    Set objPrintDriver.Document = objTestDoc
    objPrintDriver.Execute        "PrintTemplate.HTML"   _
                                & "?"                                  _
                                & 0_                
                                & "|"                                  
                                '& Application.ApplicationProxyID
    
    g_nPrintJobID               = g_nPrintJobID + 1

                                                    ' Needed to trigger printing....
    g_objDocWinHandle.PostMessage  &H0100, 0, 0
    g_objDocWinHandle.PostMessage  &H0101, 0, &HC0000000


		fctPrintHTML = True
		
	Loop While False
	
End Function

'********************************************************************************
'********************************************************************************
