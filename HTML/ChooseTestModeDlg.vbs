'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   R U N M A N U A L S E T T I N G           *
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



    Dim nNumberOfTireTypes
    nNumberOfTireTypes = window.dialogArguments.Item( "NumberOfTireTypes")

'    MsgBox window.dialogArguments.Item("NumberOfSelectedTire") & vbTab & nNumberOfTireTypes

    pSize.InnerText = window.dialogArguments.Item ( "Size" ) & ":"
	
    Dim i
    For i = 1 To nNumberOfTireTypes

      Dim testoption
      Set testoption =  document.CreateElement("Option")
      testoption.text = window.dialogArguments.Item ( "Option" & i)
      testoption.value = i
      TireTypeSelect.Add(testoption)

      
      If i =  window.dialogArguments.Item("NumberOfSelectedTire") Then
         testoption.selected = True
      End If

    Next


    STANDARDDLG_BUTTONS.CloseDlg  = FALSE
    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )

End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************

sub STANDARDDLG_BUTTONS_onClicked()

   If STANDARDDLG_BUTTONS.Value = 1 Then ' OK

      window.returnvalue = window.dialogArguments.Item ( "Option" & TireTypeSelect.SelectedIndex + 1)

	  window.dialogArguments.Add "Result", window.dialogArguments.Item ( "Option" & TireTypeSelect.SelectedIndex + 1)
	  
'    MsgBox "window.returnvalue=" & window.returnvalue & vbTab & "Option" & TireTypeSelect.SelectedIndex + 1
	  
   Else

      window.returnvalue = 0

   End If

   window.close

End Sub

'********************************************************************************
'********************************************************************************