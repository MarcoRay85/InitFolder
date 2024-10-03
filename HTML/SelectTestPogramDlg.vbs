'********************************************************************************
'*                                                                              *
'*   S C R I P T   F I L E    F O R   S e l e c t T e s t P r o g r a m D l g   *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************

Option Explicit

Dim g_strPath
Dim g_strCloseAccel

onInitialize

'********************************************************************************
'*                                                                              *
'*                                  M A I N                                     *
'*                                                                              *
'********************************************************************************

Sub onInitialize

    Dim bHasToDisplay
    Dim strSelected
    Dim strAttrib
    Dim objXMLDOM
    Dim objXMLChild
    Dim objElem
    Dim objRow
    Dim objFocusElem

    Set objFocusElem  = Nothing

    With window.DialogArguments

        g_strPath       =  .Item( "XMLPath"          )
        g_strCloseAccel =  .Item( "CloseAccel"       )
        strSelected     =  .Item( "Selected"         )
        Set objXMLDOM   =  .Item( "TestProgXMLDOM"   )

        .Item( "Selected" ) = NULL

    End With

    Document.Body.RuntimeStyle.cssText = "Background-Color:threedface;Overflow:auto;Visibility:hidden"

    Document.Body.InsertAdjacentHTML                        _
        "afterBegin",                                       _
        "<TABLE ID      =""tableMain"" "                    _
      &        "cellpadding=0 "                             _
      &        "cellspacing=0 "                             _
      &        "STYLE =""Font-Family:Arial;"                _
      &                 "Font-Size:18pt;"                   _
      &                 "Font-Weight:600;"" >"              _
      &   "<COLGROUP ALIGN=""center"" ></COLGROUP>"         _
      &   "<COLGROUP ></COLGROUP>"                          _
      & "</TABLE>"

    strAttrib = objXMLDOM.DocumentElement.GetAttribute( "HTMLStyle" )
    If Not IsNull( strAttrib ) Then tableMain.RuntimeStyle.CSSText = strAttrib

    For Each objXMLChild In objXMLDOM.DocumentElement.ChildNodes

        Set objRow = tableMain.InsertRow

        objRow.SetAttribute "_XML_", objXMLChild
        objRow.Style.Padding = "0.4ex"

        objRow.TabIndex     =  1
        objRow.onClick      =  GetRef( "onItemClicked" )
        objRow.onFocus      =  GetRef( "onItemFocus"   )
        objRow.onBlur       =  GetRef( "onItemBlur"    )
        objRow.onMouseOver  =  GetRef( "onItemMOver"   )
        objRow.onMouseDown  =  GetRef( "onItemMOver"   )
        strAttrib           =  objXMLChild.GetAttribute( "Info" )
        If IsNull( strAttrib ) Then strAttrib =  objXMLChild.GetAttribute( "Path" )
        objRow.Title        =  strAttrib

        If strSelected = objXMLChild.GetAttribute( "Name" )  Then Set objFocusElem = objRow

        Set objElem = objRow.InsertCell
        strAttrib   =  objXMLChild.GetAttribute( "Accel" )

        If IsNumeric( strAttrib ) Then

            If ( CInt( strAttrib ) > 0 ) And ( CInt( strAttrib ) < 10 ) Then

                objElem.InnerHTML = "<DIV STYLE=""Font-Size:60%; border: 2 outset;"  _
                                  &   "padding: 1 3 1 3;"                             _
                                  &   "background-color:buttonface"" >"              _
                                  & CInt( strAttrib )                                _
                                  & "</DIV>"

            End If

        End If

        strAttrib           =  objXMLChild.GetAttribute( "HTMLStyle" )
        If Not IsNull( strAttrib ) Then objElem.RuntimeStyle.CSSText = strAttrib

        Set objElem = objRow.InsertCell
        objElem.NoWrap    =  TRUE
        objElem.InnerText = objXMLChild.GetAttribute( "Name" )

        If Not IsNull( strAttrib ) Then objElem.Style.CSSText = strAttrib

    Next

    With Document.Body

        strAttrib = .ClientHeight - tableMain.OffsetHeight

        If strAttrib > 0 Then

            .Style.OverflowY = "hidden"

            Window.DialogHeight = ( GetNumericPart( Window.DialogHeight ) - strAttrib     ) & "px"
            Window.DialogTop    = ( GetNumericPart( Window.DialogTop    ) + strAttrib \ 2 ) & "px"

        End If

        strSelected = 0

        Set strAttrib = tableMain.GetBoundingClientRect

        strAttrib = .ClientWidth - strAttrib.Right + strAttrib.Left

        If strAttrib > 0 Then

            .Style.OverflowX = "hidden"

            Window.DialogWidth = ( GetNumericPart( Window.DialogWidth ) - strAttrib     ) & "px"
            Window.DialogLeft  = ( GetNumericPart( Window.DialogLeft  ) + strAttrib \ 2 ) & "px"

        End If

        .RuntimeStyle.Visibility = "visible"

    End With

    If objFocusElem Is Nothing Then

        tableMain.Rows.Item( 0 ).Focus

    Else

        Window.SetTimeout "SetFocus( " & objFocusElem.sourceIndex & " )", 100

    End If

End Sub

'********************************************************************************
'*                                                                              *
'*                         K E Y       H A N D L E R                            *
'*                                                                              *
'********************************************************************************

Sub document_onkeydown

    Dim objElem

    With Window.Event

        .CancelBubble = TRUE

        Select Case .KeyCode

            Case g_strCloseAccel
                        Window.Close
                        .ReturnValue = FALSE
                                                                ' [RETURN]
            Case 13     Set objElem = Document.ActiveElement
                        If Not ( objElem Is Nothing ) Then objElem.Click
                        .ReturnValue = FALSE

            Case 27     Window.Close                            ' [ESC]
                        .ReturnValue = FALSE

            Case 32     Set objElem = Document.ActiveElement    ' [SPACE]
                        If Not ( objElem Is Nothing ) Then objElem.Click
                        .ReturnValue = FALSE
                                                                ' [PG-UP]
            Case 35     tableMain.Rows.Item( tableMain.Rows.Length - 1 ).Focus
                        .ReturnValue = FALSE
            Case 36     tableMain.Rows.Item( 0 ).Focus          ' [PG-DOWN]
                        .ReturnValue = FALSE

            Case 38     Set objElem = Document.ActiveElement    ' [UP]
                        If Not ( objElem Is Nothing ) Then Set objElem = objElem.PreviousSibling
                        If Not ( objElem Is Nothing ) Then objElem.Focus
                        .ReturnValue = FALSE

            Case 40     Set objElem = Document.ActiveElement    ' [DOWN]
                        If Not ( objElem Is Nothing ) Then Set objElem = objElem.NextSibling
                        If Not ( objElem Is Nothing ) Then objElem.Focus
                        .ReturnValue = FALSE

        End Select

    End With

End Sub

'********************************************************************************

Sub document_onkeypress

    Dim strKey
    Dim objElem

    With Window.Event

        .CancelBubble = TRUE

        strKey = UCase( Chr( .KeyCode ))

        If ( strKey > "0" ) And ( strKey <= "9" ) Then

            For Each objElem In tableMain.Rows

                If objElem.Cells.Item( 0 ).InnerText = strKey Then objElem.Click : Exit For

            Next

        ElseIf  ( strKey >= "A" ) And ( strKey <= "_" ) Then

            Set objElem = Document.ActiveElement

            If Not ( objElem Is Nothing ) Then

                If UCase( objElem.TagName ) = "TR" Then

                    If UCase( Left( objElem.Cells.Item( 1 ).InnerText, 1 )) = strKey Then

                        Set objElem = objElem.NextSibling

                        If Not ( objElem Is Nothing ) Then

                            If UCase( Left( objElem.Cells.Item( 1 ).InnerText, 1 )) = strKey Then objElem.Focus : Exit Sub

                        End If

                    End If

                End If

            End If


            For Each objElem In tableMain.Rows

                If UCase( Left( objElem.Cells.Item( 1 ).InnerText, 1 )) = strKey Then objElem.Focus : Exit Sub

            Next

        End If

    End With

End Sub

'********************************************************************************
'*                                                                              *
'*                         F O C U S   H A N D L E R                            *
'*                                                                              *
'********************************************************************************

Sub onItemFocus

    Window.Event.CancelBubble = TRUE

    Dim objTR
    Set objTR = Window.Event.SrcElement

    Do Until objTR Is Nothing

        If UCase( objTR.TagName ) = "TR" Then

            objTR.RuntimeStyle.Color           = "captiontext"
            objTR.RuntimeStyle.BackgroundColor = "activecaption"

            Exit Do

        End If

        Set objTR = objTR.ParentElement

        If UCase( objTR.TagName ) = "TABLE" Then Exit Do

    Loop

End Sub

'********************************************************************************
'*                                                                              *
'*                         B L U R   H A N D L E R                              *
'*                                                                              *
'********************************************************************************

Sub onItemBlur

    Window.Event.CancelBubble = TRUE
    With Window.Event.SrcElement

        .RuntimeStyle.Color           = ""
        .RuntimeStyle.BackgroundColor = ""

    End With

End Sub

'********************************************************************************
'*                                                                              *
'*                 M O U S E    O V E R    H A N D L E R                        *
'*                                                                              *
'********************************************************************************

Sub onItemMOver

    Window.Event.CancelBubble = TRUE

    Dim objTR
    Set objTR = Window.Event.SrcElement

    Do Until objTR Is Nothing

        If UCase( objTR.TagName ) = "TR" Then

            Window.Event.ReturnValue = FALSE
            objTR.Focus

            Window.SetTimeout "SetFocus( " & objTR.sourceIndex & " )", 1

            Exit Do

        End If

        Set objTR = objTR.ParentElement

        If UCase( objTR.TagName ) = "TABLE" Then Exit Do

    Loop

End Sub

'********************************************************************************
'*                                                                              *
'*                         C L I C K   H A N D L E R                            *
'*                                                                              *
'********************************************************************************

Sub onItemClicked

    Dim objXMLItem

    Window.Event.CancelBubble = TRUE

    Dim objTR
    Set objTR = Window.Event.SrcElement

    Do Until objTR Is Nothing

        If UCase( objTR.TagName ) = "TR" Then

            If IsObject( objTR.GetAttribute( "_XML_" )) Then

                Set objXMLItem = objTR.GetAttribute( "_XML_" )

                If Not ( objXMLItem Is Nothing ) Then

                    window.DialogArguments.Item( "Selected" ) = objXMLItem.GetAttribute( "Name" )
                    Window.Close

                End If

            End If

            Exit Do

        End If

        Set objTR = objTR.ParentElement

        If UCase( objTR.TagName ) = "TABLE" Then Exit Do

    Loop

End Sub

'********************************************************************************
'*                                                                              *
'*                        H E L P E R    F U N C T I O N S                      *
'*                                                                              *
'********************************************************************************

Function GetNumericPart( strMeasUnit )

    Dim nPos
    nPos = 0

    Do While IsNumeric( Left( strMeasUnit, nPos + 1 ))

        nPos= nPos+ 1

    Loop

    GetNumericPart = CInt( Left( strMeasUnit, nPos ))

End Function

'********************************************************************************

Sub SetFocus( nItemPos )

    Document.All( nItemPos ).Focus

End Sub

'********************************************************************************
'********************************************************************************
