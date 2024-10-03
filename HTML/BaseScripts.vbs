'********************************************************************************
'*                                                                              *
'*       B A S E   S C R I P T   F I L E    F O R   A L L   P A G E S           *
'*                                                                              *
'********************************************************************************

'********************************************************************************
'*                                                                              *
'*                                  G L O B A L S                               *
'*                                                                              *
'********************************************************************************

Option Explicit

Dim Base_strUserLang


'************************************************************************
'*                                                                      *
'*  TASK:      Retrieves XML resource of given ID attempting to         *
'*             fulfill user's language settings                         *
'*                                                                      *
'*  NEED:      Symbolic name of resource                                *
'*                                                                      *
'*  RETURN:    Matching XMLNode or nothing if no such node available    *
'*                                                                      *
'*  EFFECTS:   HTML DOM has to contain an XML island named 'xmlRes'     *
'*             of given format (symbolic message name is 'Msg')         *
'*                                                                      *
'*                  <Resources>                                         *
'*                    <Msg>                                             *
'*                      <EN>Example</EN>                                *
'*                      <DE>Beispiel</DE>                               *
'*                    </Msg>                                            *
'*                  </Resources>                                        *
'*                                                                      *
'************************************************************************

Function Base_GetXMLResNode( strID )

    Dim vTmp
    Dim strLangID
    Dim objRes
    Set Base_GetXMLResNode = xmlRes.DocumentElement.SelectSingleNode( strID )

    If  IsEmpty( Base_strUserLang )  Then  
    
        On Error Resume Next
        
            Base_strUserLang = Application.LanguageID

        On Error Goto 0

        If  IsEmpty( Base_strUserLang )  Then  Base_strUserLang = UCase( Window.ClientInformation.UserLanguage )

    End If


    If  Not ( Base_GetXMLResNode Is Nothing )  Then

        strLangID  =  Base_strUserLang

        Set objRes =  Base_GetXMLResNode.SelectSingleNode( strLangID )

        If  Not ( objRes Is Nothing )  Then  
        
            Set Base_GetXMLResNode = objRes

        Else

            vTmp      = Split( strLangID, "-" )
            strLangID = vTmp( 0 )
        
            Set objRes = Base_GetXMLResNode.SelectSingleNode( strLangID )

            If  Not ( objRes Is Nothing )  Then  

                Set Base_GetXMLResNode = objRes

            Else

                For Each objRes In Base_GetXMLResNode.ChildNodes

                    If objRes.NodeType = 1  Then

                        Set Base_GetXMLResNode = objRes
                        Exit For

                    End If

                Next

            End If

        End If

    End If

End Function

'************************************************************************
'*                                                                      *
'*  TASK:      Retrieves string resource of given ID attempting to      *
'*             fulfill user's language settings                         *
'*                                                                      *
'*  NEED:      Symbolic name of resource                                *
'*                                                                      *
'*  RETURN:    String containing requested resource                     *
'*                                                                      *
'*  EFFECTS:   HTML DOM has to contain an XML island named 'xmlRes'     *
'*             of given format                                          *
'*                                                                      *
'*                  <Resources>                                         *
'*                    <Msg>                                             *
'*                      <EN>Example</EN>                                *
'*                      <DE>Beispiel</DE>                               *
'*                    </Msg>                                            *
'*                  </Resources>                                        *
'*                                                                      *
'*             If 'GetText( "Msg" )' is invoked, it will return         *
'*             "Beispiel", if user language is "DE" or "Example"        *
'*             otherwise. First localized resource is take, if no       *
'*             language identifier fits to user language                *
'*                                                                      *
'************************************************************************

Function Base_GetText( strID )

    Base_GetText = "dummy"
    Exit Function
    
    Dim objNode
    Set objNode = Base_GetXMLResNode( strID )

    If  Not ( objNode Is Nothing )  Then

        Base_GetText = objNode.Text

    Else

        MsgBox "Resource '" & strID & "' doesn't exist!"

    End If

End Function

'********************************************************************************
'********************************************************************************
