'********************************************************************************
'*                                                                              *
'*        S C R I P T   F I L E    F O R  S h o w H e a d C h e c k D l g         *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************
Option Explicit

    Dim objShellServices, objFSO, objWShell              
    Set objShellServices = CreateObject( "ScriptingToolsSO.ShellServices" )
    Set objFSO           = CreateObject( "Scripting.FileSystemObject"     )
    Set objWShell        = CreateObject( "WScript.Shell"                  )

'********************************************************************************


Const C_strIIDPrefix            =  "IID_"
Const C_strToolIDPrefix         =  "spanTool"
Const C_strToolBtnClass         =  "classToolButton"
Const C_strAttibToolPath        =  "Path"

Const C_AttrNameAppObj          =  "_AppObj_"
Const C_AttrShareCnt            =  "_SHARES_"
Const C_AttrCanceled            =  "_CANCEL_"

Const C_ClassError              =  "classERROR"
Const C_LangVBScript            =  "VBScript"

Const C_strStateProp            =  "State"

Const C_nDefaultSizeFactor      =  0.75

Const WM_CLOSE                  =  &H0010
Const WM_SYSCOMMAND             =  &H0112
Const SC_CLOSE                  =  61536    '0xF060


Dim g_nCurPage
Dim g_nProgressDlgRefCnt
Dim g_nOutputSizeFactor

Dim g_strCurPageName
Dim g_strStartupScript


Dim Application
Dim g_objFSO
Dim g_objSTSOFactory
Dim g_objVBXens
Dim g_objObjectInfo
Dim g_objWinHandle
Dim g_objLogHost
Dim g_objPLCList
Dim g_objEPCLogin

Dim g_objBeforeQuitHdl
Dim g_objHTMLProgressDlg


Set g_objFSO                =  CreateObject(  "Scripting.FileSystemObject"      )
Set g_objEPCLogin           =  CreateObject(  "EPCClient.EPCLogin"              )

Set g_objBeforeQuitHdl      =  Nothing
Set g_objHTMLProgressDlg    =  Nothing

Set g_objPLCList            =  Nothing

g_nProgressDlgRefCnt        =  0 
g_nOutputSizeFactor         =  C_nDefaultSizeFactor
    

Dim g_objMachine
Dim g_objMachineScript
Dim g_objMachineState
Dim g_objVBSXtensions
Dim g_objHTMLHelper
Dim g_objWinTools


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

            Next

            Exit For

        End If

    Next

 
End Sub

'***********************************************************************
'****                                                               ****
'****       Handler for form initialisation                         ****
'****                                                               ****
'***********************************************************************

Sub window_onload()     'Initialisation

End Sub

'***********************************************************************
'****                                                               ****
'****       Handler for form unloading                              ****
'****                                                               ****
'***********************************************************************

Sub Window_OnUnload()

    Dim vTmp
 
End Sub

'********************************************************************************
'*                                                                              *
'*     H A N D L E R   F O R   B U T T O N    C L I C K S                       *
'*                                                                              *
'********************************************************************************

Sub STANDARDDLG_BUTTONS_onClicked()
    
    If STANDARDDLG_BUTTONS.Value = 1 Then
    
    ElseIF STANDARDDLG_BUTTONS.Value = 256 Then
        if not fctPrintHTML() then 
        msgbox("Print failed!")
    End If
    Else
     '   msgbox("STANDARDDLG_BUTTONS.Value: " & STANDARDDLG_BUTTONS.Value)
        
    End if
    window.close

End Sub

'***********************************************************************
'****                                                               ****
'****       Handler for suppressing WM_CLOSE                        ****
'****                                                               ****
'***********************************************************************

Sub MsgHdl( objMessage )

    If  objMessage.wParam <> SC_CLOSE  Then  Exit Sub

    If  Not ( g_objBeforeQuitHdl Is Nothing )  Then

        If  Not g_objBeforeQuitHdl()  Then

            objMessage.Msg = 0
            Exit Sub

        End If

    End If
    
End Sub


'***********************************************************************
'***********************************************************************
