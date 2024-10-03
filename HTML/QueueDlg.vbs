'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   Q U E U E  D l g                          *
'*                                                                              *
'********************************************************************************
    Option Explicit

    Dim objRedrawHDL
    
    Dim g_Args 
    g_Args = window.DialogArguments 
    Dim vTmp 
    Set vTmp = g_args(1) 
    Set objRedrawHDL = g_args(2) 
    vtmp me
    Window.ExecScript  g_Args( 0 ), "VBScript"      

Sub onInitialize
    Window.SetInterval "MonitorTask", 500, "VBScript"
End Sub

Sub MonitorTask
   objRedrawHDL()
End Sub

