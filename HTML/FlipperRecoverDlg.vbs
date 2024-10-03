'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   F l i p p e r R e c o v e r D l g         *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************

Option Explicit
                                                'Text resources used in code
Dim g_nStateRetainer
Dim g_nStateRotor
Dim g_nStateConveyor
Dim g_objRetainerHdl
Dim g_objRotorHdl
Dim g_objConveyorHdl

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim vTmp

    If  Document.Body.IsDialog  Then

        vTmp                    =  Window.DialogArguments
        Set g_objRetainerHdl    =  vTmp( 0 )
        Set g_objRotorHdl       =  vTmp( 1 )
        Set g_objConveyorHdl    =  vTmp( 2 )

        Window.SetInterval "MonitorTask", 50, "VBScript"

    End If

    buttonRetOpen.onmousedown       = GetRef( "RetainerOpen"  )
    buttonRetOpen.onmouseup         = GetRef( "RetainerStop"  )
    buttonRetOpen.onmouseout        = GetRef( "RetainerStop"  )

    buttonRetClose.onmousedown      = GetRef( "RetainerClose" )
    buttonRetClose.onmouseup        = GetRef( "RetainerStop"  )
    buttonRetClose.onmouseout       = GetRef( "RetainerStop"  )

    buttonRotLeft.onmousedown       = GetRef( "RotateLeft"    )
    buttonRotLeft.onmouseup         = GetRef( "RotateStop"    )
    buttonRotLeft.onmouseout        = GetRef( "RotateStop"    )

    buttonRotRight.onmousedown      = GetRef( "RotateRight"   )
    buttonRotRight.onmouseup        = GetRef( "RotateStop"    )
    buttonRotRight.onmouseout       = GetRef( "RotateStop"    )

    buttonConveyLeft.onmousedown    = GetRef( "ConveyLeft"    )
    buttonConveyLeft.onmouseup      = GetRef( "ConveyStop"    )
    buttonConveyLeft.onmouseout     = GetRef( "ConveyStop"    )

    buttonConveyRight.onmousedown   = GetRef( "ConveyRight"   )
    buttonConveyRight.onmouseup     = GetRef( "ConveyStop"    )
    buttonConveyRight.onmouseout    = GetRef( "ConveyStop"    )


End Sub

'********************************************************************************

Sub MonitorTask()

    If  Not IsEmpty( g_nStateRetainer )  Then  g_objRetainerHdl(  g_nStateRetainer  ) : If  g_nStateRetainer = 0  Then  g_nStateRetainer = Empty
    If  Not IsEmpty( g_nStateRotor    )  Then  g_objRotorHdl(     g_nStateRotor     ) : If  g_nStateRotor    = 0  Then  g_nStateRotor    = Empty
    If  Not IsEmpty( g_nStateConveyor )  Then  g_objConveyorHdl(  g_nStateConveyor  ) : If  g_nStateConveyor = 0  Then  g_nStateConveyor = Empty

End Sub

'********************************************************************************

Sub RetainerOpen

    g_nStateRetainer = -1

End Sub

'********************************************************************************

Sub RetainerClose

    g_nStateRetainer = 1

End Sub

'********************************************************************************

Sub RetainerStop

    g_nStateRetainer = 0

End Sub

'********************************************************************************

Sub RotateLeft

    g_nStateRotor = -1

End Sub

'********************************************************************************

Sub RotateRight

    g_nStateRotor = 1

End Sub

'********************************************************************************

Sub RotateStop

    g_nStateRotor = 0

End Sub

'********************************************************************************

Sub ConveyLeft

    g_nStateConveyor = 1

End Sub

'********************************************************************************

Sub ConveyRight

    g_nStateConveyor = -1

End Sub

'********************************************************************************

Sub ConveyStop

    g_nStateConveyor = 0

End Sub

'********************************************************************************
'********************************************************************************
