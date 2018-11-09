Sub test_sharepoint()
    
    Dim responseFormat As Integer, jsonFormat As Boolean
    responseFormat = MsgBox("Retrieve in JSON (Y), or in XML (N) data format?", vbYesNo + vbQuestion, "Response Data Format")
    If responseFormat = vbYes Then
        jsonFormat = True
    Else
        jsonFormat = False
    End If
        
    Dim spoIDCRLTokenCookie As String, webTitle As String, ceIssues As String
    
    spoIDCRLTokenCookie = ""
        
    Dim logonMode As Integer
    logonMode = MsgBox("Logon with specific account (instead of direct)?", vbYesNo + vbQuestion, "Logon mode")
        
    If logonMode = vbYes Then
        ' Get credentials
        Dim CurUserName As String, CurPassword As String
    
        CurUserName = InputBox("username")
        CurPassword = InputBox("password")
    
        If Len(CurUserName) = 0 Or Len(CurPassword) = 0 Then
            MsgBox "Need to enter username + password"
        Else
            spoIDCRLTokenCookie = O365SPO_ActiveAuthentication.GetO365SPO_SamlAuthentication(CurUserName, CurPassword)
        End If
    Else
        spoIDCRLTokenCookie = O365SPO_ActiveAuthentication.GetO365SPO_SamlAuthenticationIntegrated
    End If
    
    If Len(spoIDCRLTokenCookie) > 0 Then
        Dim msgBoxTitle
        
        If logonMode = vbYes Then
            msgBoxTitle = "Determined Username Active Authentication Cookie"
        Else
            msgBoxTitle = "Determined Integrated Active Authentication Cookie"
        End If
        
        webTitle = GetO365SPO_WebTitle(spoIDCRLTokenCookie, jsonFormat)
        MsgBox webTitle, vbApplicationModal, msgBoxTitle

        listItems = GetO365SPO_ListItems(spoIDCRLTokenCookie, jsonFormat)
        MsgBox listItems, vbApplicationModal, msgBoxTitle
    Else
        MsgBox "Failed to retrieve O365 authentication cookie"
    End If
End Sub

Function GetO365SPO_WebTitle(spoIDCRLTokenCookie As String, jsonFormat As Boolean) As String
    Dim responseNode As String, selectFields As String
    If jsonFormat Then
        responseNode = "d"
        selectFields = "Title"
    Else
        responseNode = "d:Title"
        selectFields = ""
    End If
    GetO365SPO_WebTitle = O365SPO_InvokeRestMethod("https://<YOUR-TENANT>.sharepoint.com/_api/web/Title", responseNode, selectFields, spoIDCRLTokenCookie, jsonFormat, False)
End Function

Function GetO365SPO_ListItems(spoIDCRLTokenCookie As String, jsonFormat As Boolean) As String
    Dim GetCeIssuesRequest As String
    GetCeIssuesRequest = "https://<YOUR-TENANT.sharepoint.com/teams/<YOUR-SITE>/_api/web/lists/GetByTitle('<YOUR-LIST>')/Items?$select=Id,field_2,field_X"
    Dim responseNode As String, selectFields As String
    If jsonFormat Then
        responseNode = "d.results"
        selectFields = "Id|field_2|field_X"
    Else
        responseNode = "//m:properties"
        selectFields = "d:Id|d:field2|d:field_X"
    End If
    GetO365SPO_ListItems = O365SPO_InvokeRestMethod(GetCeIssuesRequest, responseNode, selectFields, spoIDCRLTokenCookie, jsonFormat, True)
End Function

Private Function O365SPO_InvokeRestMethod(requestUrl As String, responseNode As String, selectFields As String, spoIDCRLTokenCookie As String, jsonFormat As Boolean, includeFieldNameInOutput As Boolean) ' As String
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    Request.Open "GET", requestUrl, False
    If Len(spoIDCRLTokenCookie) > 0 Then
        Request.setRequestHeader "Cookie", "SPOIDCRL=" & spoIDCRLTokenCookie
    End If
    If jsonFormat Then
        Request.setRequestHeader "Accept", "application/json;odata=verbose"
    End If
    
    Request.send

    If Request.Status = 200 Then
        Dim fields() As String
        Dim result As String, resultPart As String, resultLine As String
        Dim i As Integer, j As Integer
        
        result = ""
    
        If Len(selectFields) > 0 Then
            fields = Split(selectFields, "|")
        End If
            
        If Not jsonFormat Then
            Set XmlDocument = CreateObject("Msxml2.DOMDocument")
            With XmlDocument
                .async = False
                .validateOnParse = False
                .resolveExternals = False
                .SetProperty "SelectionNamespaces", "xmlns=""http://www.w3.org/2005/Atom"" xmlns:d=""http://schemas.microsoft.com/ado/2007/08/dataservices"" xmlns:m=""http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"" xmlns:georss=""http://www.georss.org/georss"" xmlns:gml=""http://www.opengis.net/gml"""

                Loaded = .LoadXML(Request.responseText)
            End With
            If Loaded Then
                If Len(selectFields) = 0 Then
                    O365SPO_InvokeRestMethod = XmlDocument.SelectSingleNode(responseNode).Text
                Else
                    Dim nodeValue As String
                    
                    result = ""

                    Set ListEntries = XmlDocument.SelectNodes(responseNode)
                    For Each ListEntry In ListEntries
                        For j = 0 To UBound(fields)

                            nodeValue = ExtractXmlNode(ListEntry.xml, fields(j), True)
                            resultPart = nodeValue
                            If includeFieldNameInOutput Then
                                resultPart = fields(j) & ":" & nodeValue
                            End If
                            
                            If UBound(fields) > 0 Then
                                result = result & "[" & resultPart & "]"
                            Else
                                result = resultPart
                            End If
                        
                        Next
                    Next
                    
                    O365SPO_InvokeRestMethod = result
                End If
            End If  
        Else
            Dim jsonNodes() As String
            jsonNodes = Split(responseNode, ".")

            
            Dim jsonObject As Object, jsonNode As Object
            Dim elemObject As Object
            Dim propValue As Variant
            Dim keys() As String

            
            JsonParser.InitScriptEngine
            Set jsonRoot = JsonParser.DecodeJsonString(Request.responseText)
            
            Set jsonNode = jsonRoot
            For i = 0 To UBound(jsonNodes)
                Set jsonNode = JsonParser.GetObjectProperty(jsonNode, jsonNodes(i))
            Next i
            
            If jsonNodes(UBound(jsonNodes)) <> "results" Then
                For j = 0 To UBound(fields)
                    propValue = JsonParser.GetProperty(jsonNode, fields(j))
                    If IsNull(propValue) Then
                        propValue = "[Null]"
                    End If
                    resultPart = propValue
                    If includeFieldNameInOutput Then
                        resultPart = fields(j) & ":" & propValue
                    End If
                    
                    If UBound(fields) > 0 Then
                        result = result & "[" & resultPart & "]"
                    Else
                        result = resultPart
                    End If
                    Debug.Print Space(2) & fields(j) & ": " & propValue
                Next
            Else
                keys = JsonParser.GetKeys(jsonNode)
            
                For i = 0 To UBound(keys)
                    resultLine = ""
                    
                    If JsonParser.GetPropertyType(jsonNode, keys(i)) = jptValue Then
                        propValue = JsonParser.GetProperty(jsonNode, keys(i))
                        If IsNull(propValue) Then
                            propValue = "[Null]"
                        End If
                    Else
                        Set elemObject = JsonParser.GetObjectProperty(jsonNode, keys(i))
                        For j = 0 To UBound(fields)
                            propValue = JsonParser.GetProperty(elemObject, fields(j))
                            If IsNull(propValue) Then
                                propValue = "[Null]"
                            End If
                            
                            resultPart = propValue
                            If includeFieldNameInOutput Then
                                resultPart = fields(j) & ":" & propValue
                            End If
                            
                            If UBound(fields) > 0 Then
                                resultLine = resultLine & "[" & resultPart & "]"
                            Else
                                resultLine = resultPart
                            End If
                        Next
                        
                        If Len(result) <> 0 Then
                            result = result & vbCrLf
                        End If
                        result = result & resultLine
                    End If
            
                Next i
            End If
    
            O365SPO_InvokeRestMethod = result
            
        End If
    Else
        O365SPO_InvokeRestMethod = "[Error]: " & Request.Status & " - " & Request.responseText
    End If
End Function