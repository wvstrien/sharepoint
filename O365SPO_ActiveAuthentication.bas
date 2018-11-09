Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As LongPtr
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr

Public Function GetO365SPO_SamlAuthenticationIntegrated() As String
    Dim samlAssertion As String, binarySecurityToken As String, spoIDCRLTokenCookie As String
    
    samlAssertion = GetO365SPO_SAMLAssertionIntegrated()
    If Len(samlAssertion) > 0 Then
        binarySecurityToken = GetO365SPO_BinarySecurityToken(samlAssertion)
        spoIDCRLTokenCookie = GetO365SPO_SPOIDCRLTokenAsCookie(binarySecurityToken)
    End If
    
    GetO365SPO_SamlAuthenticationIntegrated = spoIDCRLTokenCookie
End Function

Public Function GetO365SPO_SamlAuthentication(CurUserName As String, CurPassword As String) As String
    Dim samlAssertion As String, binarySecurityToken As String, spoIDCRLTokenCookie As String
    
    samlAssertion = GetO365SPO_SAMLAssertion(CurUserName, CurPassword)
    If Len(samlAssertion) > 0 Then
        binarySecurityToken = GetO365SPO_BinarySecurityToken(samlAssertion)
        spoIDCRLTokenCookie = GetO365SPO_SPOIDCRLTokenAsCookie(binarySecurityToken)
    End If
    
    GetO365SPO_SamlAuthentication = spoIDCRLTokenCookie
End Function

Private Function O365SPO_CreateGuidString()
    Dim guid As GUID_TYPE
    Dim strGuid As String
    Dim retValue As LongPtr
    Const guidLength As Long = 39 'registry GUID format with null terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}

    retValue = CoCreateGuid(guid)
    If retValue = 0 Then
        strGuid = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(guid, StrPtr(strGuid), guidLength)
        If retValue = guidLength Then
            ' valid GUID as a string
            O365SPO_CreateGuidString = strGuid
        End If
    End If
End Function

Private Function O365SPO_RandomString(Length As Integer)

    Dim CharacterBank As Variant
    Dim x As Long
    Dim str As String
    
      If Length < 1 Then
        MsgBox "Length variable must be greater than 0"
        Exit Function
      End If
    
    CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
      "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
      "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "!", "@", _
      "#", "$", "-", "^", "*", "A", "B", "C", "D", "E", "F", "G", "H", _
      "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
      "W", "X", "Y", "Z")
  
    'Randomly Select Characters One-by-One
    For x = 1 To Length
        Randomize
        str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
    Next x
    
    O365SPO_RandomString = str
End Function

Private Function O365SPO_ExtractXmlNode(xml As String, name As String, valueOnly As Boolean) As String
    Dim nodeValue As String
    nodeValue = Mid(xml, InStr(xml, "<" & name))
    If valueOnly Then
        nodeValue = Mid(nodeValue, InStr(nodeValue, ">") + 1)
        O365SPO_ExtractXmlNode = Left(nodeValue, InStr(nodeValue, "</" & name) - 1)
    Else
        O365SPO_ExtractXmlNode = Left(nodeValue, InStr(nodeValue, "</" & name) + Len(name) + 2)
    End If
End Function

Private Function GetO365SPO_SAMLAssertionIntegrated() As String
    Dim CustomStsUrl As String, CustomStsSAMLRequest, stsMessage As String
    
    CustomStsUrl = "https://sts.<YOUR-TENANT>.com/adfs/services/trust/2005/windowstransport"
    CustomStsSAMLRequest = "<?xml version=""1.0"" encoding=""UTF-8""?><s:Envelope xmlns:s=""http://www.w3.org/2003/05/soap-envelope"" xmlns:a=""http://www.w3.org/2005/08/addressing"">" & _
            "<s:Header>" & _
                "<a:Action s:mustUnderstand=""1"">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>" & _
                "<a:MessageID>urn:uuid:[[messageID]]</a:MessageID>" & _
                "<a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo>" & _
                "<a:To s:mustUnderstand=""1"">[[mustUnderstand]]</a:To>" & _
            "</s:Header>"
    CustomStsSAMLRequest = CustomStsSAMLRequest & _
            "<s:Body>" & _
                "<t:RequestSecurityToken xmlns:t=""http://schemas.xmlsoap.org/ws/2005/02/trust"">" & _
                    "<wsp:AppliesTo xmlns:wsp=""http://schemas.xmlsoap.org/ws/2004/09/policy"">" & _
                        "<wsa:EndpointReference xmlns:wsa=""http://www.w3.org/2005/08/addressing"">" & _
                        "<wsa:Address>urn:federation:MicrosoftOnline</wsa:Address></wsa:EndpointReference>" & _
                    "</wsp:AppliesTo>" & _
                    "<t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>" & _
                    "<t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>" & _
                "</t:RequestSecurityToken>" & _
            "</s:Body>" & _
        "</s:Envelope>"

    
    stsMessage = Replace(CustomStsSAMLRequest, "[[messageID]]", Mid(O365SPO_CreateGuidString(), 2, 36))
    stsMessage = Replace(stsMessage, "[[mustUnderstand]]", CustomStsUrl)

    ' Create HTTP Object ==> make sure to use "MSXML2.XMLHTTP" iso "MSXML2.ServerXMLHTTP.6.0"; as the latter does not send the NTLM
    ' credentials as Authorization header.
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    ' Get SAML:assertion
    Request.Open "POST", CustomStsUrl, False
    Request.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8"
    Request.send (stsMessage)
    
    If Request.Status = 200 Then
         GetO365SPO_SAMLAssertionIntegrated = O365SPO_ExtractXmlNode(Request.responseText, "saml:Assertion", False)
    End If
    
End Function


Private Function GetO365SPO_SAMLAssertion(CurUserName As String, CurPassword As String) As String
    Dim CustomStsUrl As String, CustomStsSAMLRequest, stsMessage As String

    CustomStsUrl = "https://sts.<YOUR-TENANT>.com/adfs/services/trust/2005/usernamemixed"
    CustomStsSAMLRequest = "<s:Envelope xmlns:s=""http://www.w3.org/2003/05/soap-envelope"" xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"" xmlns:saml=""urn:oasis:names:tc:SAML:1.0:assertion"" xmlns:wsp=""http://schemas.xmlsoap.org/ws/2004/09/policy"" xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"" xmlns:wsa=""http://www.w3.org/2005/08/addressing"" xmlns:wssc=""http://schemas.xmlsoap.org/ws/2005/02/sc"" xmlns:wst=""http://schemas.xmlsoap.org/ws/2005/02/trust"">" & _
            "<s:Header>" & _
                "<wsa:Action s:mustUnderstand=""1"">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action>" & _
                "<wsa:To s:mustUnderstand=""1"">https://sts.<YOUR-TENANT>.com/adfs/services/trust/2005/usernamemixed</wsa:To>" & _
                "<wsa:MessageID>[[messageID]]</wsa:MessageID>" & _
                "<ps:AuthInfo xmlns:ps=""http://schemas.microsoft.com/Passport/SoapServices/PPCRL"" Id=""PPAuthInfo"">" & _
                    "<ps:HostingApp>Managed IDCRL</ps:HostingApp>" & _
                    "<ps:BinaryVersion>6</ps:BinaryVersion>" & _
                    "<ps:UIVersion>1</ps:UIVersion>" & _
                    "<ps:Cookies></ps:Cookies>" & _
                    "<ps:RequestParams>AQAAAAIAAABsYwQAAAAxMDMz</ps:RequestParams>" & _
                "</ps:AuthInfo>" & _
                "<wsse:Security>" & _
                    "<wsse:UsernameToken wsu:Id=""user"">" & _
                        "<wsse:Username>[[username]]</wsse:Username>" & _
                            "<wsse:Password>[[password]]</wsse:Password>" & _
                    "</wsse:UsernameToken>" & _
                    "<wsu:Timestamp Id=""Timestamp"">" & _
                        "<wsu:Created>[[createdTime]]</wsu:Created>" & _
                        "<wsu:Expires>[[expiresTime]]</wsu:Expires>" & _
                    "</wsu:Timestamp>" & _
                "</wsse:Security>" & _
            "</s:Header>"
    
    CustomStsSAMLRequest = CustomStsSAMLRequest & _
            "<s:Body>" & _
                "<wst:RequestSecurityToken Id=""RST0"">" & _
                    "<wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType>" & _
                    "<wsp:AppliesTo>" & _
                        "<wsa:EndpointReference>" & _
                            "<wsa:Address>urn:federation:MicrosoftOnline</wsa:Address>" & _
                        "</wsa:EndpointReference>" & _
                    "</wsp:AppliesTo>" & _
                    "<wst:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</wst:KeyType>" & _
                "</wst:RequestSecurityToken>" & _
            "</s:Body>" & _
        "</s:Envelope>"
    
    Dim createdTime, expiresTime As String
    createdTime = UtcConverter.ConvertToIso(Now)
    expiresTime = UtcConverter.ConvertToIso(DateAdd("h", "3", Now))
    
    stsMessage = Replace(CustomStsSAMLRequest, "[[messageID]]", O365SPO_RandomString(15))
    stsMessage = Replace(stsMessage, "[[username]]", CurUserName)
    stsMessage = Replace(stsMessage, "[[password]]", CurPassword)
    stsMessage = Replace(stsMessage, "[[createdTime]]", createdTime)
    stsMessage = Replace(stsMessage, "[[expiresTime]]", expiresTime)
        
    ' Create HTTP Object
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    ' Get SAML:assertion
    Request.Open "POST", CustomStsUrl, False
    Request.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8"
    Request.send (stsMessage)
    
    If Request.Status = 200 Then
         GetO365SPO_SAMLAssertion = O365SPO_ExtractXmlNode(Request.responseText, "saml:Assertion", False)
    End If

End Function

Private Function GetO365SPO_BinarySecurityToken(samlAssertion As String) As String
    Dim spoSecurityTokenRequest As String, spoSecurityMessage As String
    spoSecurityTokenRequest = "<S:Envelope " & _
            "xmlns:S=""http://www.w3.org/2003/05/soap-envelope"" " & _
            "xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"" " & _
            "xmlns:wsp=""http://schemas.xmlsoap.org/ws/2004/09/policy"" " & _
            "xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"" " & _
            "xmlns:wsa=""http://www.w3.org/2005/08/addressing"" " & _
            "xmlns:wst=""http://schemas.xmlsoap.org/ws/2005/02/trust"">" & _
            "<S:Header>" & _
                "<wsa:Action S:mustUnderstand=""1"">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action>" & _
                "<wsa:To S:mustUnderstand=""1"">https://login.microsoftonline.com/rst2.srf</wsa:To>" & _
                "<ps:AuthInfo "
    spoSecurityTokenRequest = spoSecurityTokenRequest & _
                    "xmlns:ps=""http://schemas.microsoft.com/LiveID/SoapServices/v1"" Id=""PPAuthInfo"">" & _
                    "<ps:BinaryVersion>5</ps:BinaryVersion>" & _
                    "<ps:HostingApp>Managed IDCRL</ps:HostingApp>" & _
                "</ps:AuthInfo>" & _
                "<wsse:Security>" & _
                "[[samlAssertion]]" & _
                "</wsse:Security>" & _
            "</S:Header>" & _
            "<S:Body>" & _
                "<wst:RequestSecurityToken " & _
                    "xmlns:wst=""http://schemas.xmlsoap.org/ws/2005/02/trust"" Id=""RST0"">" & _
                    "<wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType>" & _
                    "<wsp:AppliesTo>" & _
                        "<wsa:EndpointReference>" & _
                            "<wsa:Address>sharepoint.com</wsa:Address>" & _
                        "</wsa:EndpointReference>" & _
                    "</wsp:AppliesTo>" & _
                    "<wsp:PolicyReference URI=""MBI""></wsp:PolicyReference>" & _
                "</wst:RequestSecurityToken>" & _
            "</S:Body>" & _
        "</S:Envelope>"

    spoSecurityMessage = Replace(spoSecurityTokenRequest, "[[samlAssertion]]", samlAssertion)
        
    ' Create HTTP Object
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    ' Get wsse:BinarySecurityToken
    Request.Open "POST", "https://login.microsoftonline.com/RST2.srf", False
    Request.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8"
    Request.send (spoSecurityMessage)
    
    If Request.Status = 200 Then
         GetO365SPO_BinarySecurityToken = O365SPO_ExtractXmlNode(Request.responseText, "wsse:BinarySecurityToken", True)
    End If
End Function

Private Function GetO365SPO_SPOIDCRLTokenAsCookie(binarySecurityToken As String) As String
    ' Create HTTP Object - make sure to here use ServerXMLHTTP iso XMLHTTP, as the latter does not allow access to cookie returned in response.
    Dim Request As Object
    Set Request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    ' Get wsse:BinarySecurityToken
    Dim binarySecurityTokenHeader As String
    binarySecurityTokenHeader = "BPOSIDCRL " & binarySecurityToken
    Request.Open "GET", "https://<YOUR-TENANT>.sharepoint.com/_vti_bin/idcrl.svc/", False
    Request.setRequestHeader "X-IDCRL_ACCEPTED", "t"
    Request.setRequestHeader "Authorization", binarySecurityTokenHeader

    Request.send

    If Request.Status = 200 Then
         Dim setCookie As String, spoIDCRLTokenCookie As String
         setCookie = Request.GetResponseHeader("Set-Cookie")
         spoIDCRLTokenCookie = Mid(setCookie, InStr(setCookie, "SPOIDCRL=") + Len("SPOIDCRL="))
         spoIDCRLTokenCookie = Left(spoIDCRLTokenCookie, InStr(spoIDCRLTokenCookie, ";") - 1)
         GetO365SPO_SPOIDCRLTokenAsCookie = spoIDCRLTokenCookie
    End If
End Function

