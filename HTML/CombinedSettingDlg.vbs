'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   C o m b i n e d S e t t i n g D l g       *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************

Option Explicit

const WM_CLOSE                  = &H0010

Dim g_objApplication
Dim g_objMachine
Dim g_objMachineScript
Dim g_objMachineState
Dim g_objVBSXtensions
Dim g_objHTMLHelper
Dim g_objWinTools


Set g_objApplication            = Nothing
Set g_objMachine                = Nothing
Set g_objMachineScript          = Nothing
Set g_objMachineState           = Nothing
Set g_objHTMLHelper             = CreateObject( "IEHTMLToolsSO.HTMLHelper"      )
Set g_objVBSXtensions           = CreateObject( "ScriptingToolsSO.VBSXtensions" )
Set g_objWinTools               = CreateObject( "ScriptingToolsSO.Windowing"    )


'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim objStyle
    Dim objElem
    Dim vArgs
    Dim vTmp
    Dim nI
    Dim nLen

    With Document.Body

        With .WindowHandle

            .ModifyStyle  0, &H00080000
            .UpdateLook
            .InstallMessageFilter GetRef( "MsgHdl" ), WM_CLOSE, WM_CLOSE

        End With

    End With

    For Each objStyle In Document.StyleSheets       ' Match colors to different backgrounds in NT4 and W2K

        If objStyle.ID = "linkedstyleCustom"  Then

            nLen = objStyle.Rules.Length - 1

            For nI = 0 To nLen

                Set objElem = objStyle.Rules.Item( nI )

                If objElem.SelectorText = ".classTableHeader"  Then

                    vTmp = g_objHTMLHelper.AmpColor( "ThreeDFace", 0.85 )

                    objElem.Style.BackgroundColor = vTmp

                End If

                If objElem.SelectorText = ".classButtons"  Then

                    vTmp = 56 ' FIXME: move into css

                    With objElem.Style

                        .Width      = vTmp & "px"
                        .Height     = vTmp & "px"
                        .LineHeight = vTmp & "px"

                    End With

                End If

            Next


            Exit For

        End If

    Next

    If Not Document.Body.IsDialog Then

    Else
    
        With Document.ParentWindow

            vArgs  =  .DialogArguments

            Set g_objApplication    = vArgs( 0 )
            Set g_objMachineScript  = vArgs( 1 )
            .DialogTop              = vArgs( 2 )

        End With

        Set g_objMachine      = g_objApplication.Machine
        Set g_objMachineState = g_objMachine.State


        g_objWinTools.RepaintWindowTitle  g_objApplication.HWND, TRUE

   End If
   
   If Document.Body.IsDialog Then

        If IsObject( vArgs( 3 ))  Then

            Set vTmp = vArgs( 3 )

            If Not ( vTmp Is Nothing )  Then  vTmp  Me, GetRef( "Command" ), vArgs( 4 )

        End If

    End If

    divCombinedSettingPane.HeaderStyle  = "background-color:;"
    divToolPane.HeaderStyle     = "background-color:;"

    divCombinedSettingPane.SizeToContent "H"
    divCombinedSettingPane.CenterContent "X"

    divToolPane.SizeToContent     "H"
    divToolPane.CenterContent     "X"

    STANDARDDLG_BUTTONS.onClicked = GetRef( "STANDARDDLG_BUTTONS_onClicked" )
    STANDARDDLG_BUTTONS.WithoutAlt = False
    STANDARDDLG_BUTTONS.Default = 0

End Sub

'************************************************************************
'*                                                                      *
'*  TASK:      Function executes given string within context of         *
'*             script of dialog                                         *
'*                                                                      *
'*  NEED:      String with VBScript statements                          *
'*             Arbitrary parameter                                      *
'*                                                                      *
'*  RETURN:    If return value is needed, use 'Command'                 *
'*                                                                      *
'************************************************************************

Function Command( strExec, vParameter )

    Execute  strExec

End Function

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************

Sub STANDARDDLG_BUTTONS_onClicked()

    Document.Body.WindowHandle.InstallMessageFilter Nothing, 0, 0
    
    g_objWinTools.RepaintWindowTitle  g_objApplication.HWND, FALSE

    Window.Close

End Sub

'********************************************************************************
'*                                                                              *
'*   S U P P R E S S I N G    C L O S E    M E S S A G E S                      *
'*                                                                              *
'********************************************************************************

Sub MsgHdl( objMessage )

    objMessage.Msg = 0

End Sub

'********************************************************************************
'********************************************************************************
