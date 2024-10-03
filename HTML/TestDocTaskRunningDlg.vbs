'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   T e s t D o c S t o r i n g D l g         *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************
Option Explicit

Dim g_objQueryHdl
Set g_objQueryHdl = Document.CreateStyleSheet

g_objQueryHdl.AddRule  ".classIntact",            _
                       "Font-Family : Arial;    " _
                    &  "Font-Weight : 900%;     " _
                    &  "Font-Size   : 150%;     " _
                    &  "Color       : lime;     " _
                    &  "Position    : relative; " _
                    &  "Bottom      : -0.1ex;   "

g_objQueryHdl.AddRule  ".classButton",                _
                       "Font-Family : Arial;        " _
                    &  "Font-Size   :  70%;         " _
                    &  "Border      : 2px outset;   " _
                    &  "Padding     : 0 1ex 0 1ex ; " _
                    &  "Position    : relative;     " _
                    &  "Bottom      : 0.5ex;        " _


Set g_objQueryHdl = Nothing

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim vTmp

    If Document.Body.IsDialog  Then

        vTmp              = Window.DialogArguments
        Set g_objQueryHdl = vTmp( 0 )

        SetTaskKind  vTmp( 1 )

        Window.SetInterval "onQueryFinished", 500, "VBScript"

    Else

        Randomize

        SetTaskKind  Round( Rnd )

    End If

End Sub

'********************************************************************************
'*                                                                              *
'*          Q U E R Y I N G    F O R   S T O R I N G   C O M P L E T E D        *
'*                                                                              *
'********************************************************************************

Sub onQueryFinished

    If g_objQueryHdl Then Window.Close

End Sub

'********************************************************************************
'********************************************************************************
