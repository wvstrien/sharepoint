
Option Explicit
 
Public Enum JsonPropertyType
    jptObject
    jptValue
End Enum
 
Private ScriptEngine As Object 'ScriptControl (ref: Microsoft Script Control 1.0)
 
Public Sub InitScriptEngine()
    Set ScriptEngine = CreateObject("MSScriptControl.ScriptControl") 'New ScriptControl
    ScriptEngine.Language = "JScript"
    ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
End Sub
 
Public Function DecodeJsonString(ByVal JSonString As String)
    Set DecodeJsonString = ScriptEngine.Eval("(" + JSonString + ")")
End Function
 
Public Function GetProperty(ByVal jsonObject As Object, ByVal PropertyName As String) 'As Variant
    GetProperty = ScriptEngine.Run("getProperty", jsonObject, PropertyName)
End Function
 
Public Function GetObjectProperty(ByVal jsonObject As Object, ByVal PropertyName As String) 'As Object
    Set GetObjectProperty = ScriptEngine.Run("getProperty", jsonObject, PropertyName)
End Function
 
Public Function GetPropertyType(ByVal jsonObject As Object, ByVal PropertyName As String) As JsonPropertyType
    On Error Resume Next
    Dim o As Object
    Set o = GetObjectProperty(jsonObject, PropertyName)
    If Err.Number Then
        GetPropertyType = jptValue
        Err.Clear
        On Error GoTo 0
    Else
        GetPropertyType = jptObject
    End If
End Function
 
Public Function GetKeys(ByVal jsonObject As Object) As String()
    Dim Length As Integer
    Dim KeysArray() As String
    Dim KeysObject As Object
    Dim Index As Integer
    Dim key As Variant
 
    Set KeysObject = ScriptEngine.Run("getKeys", jsonObject)
    Length = GetProperty(KeysObject, "length")
    ReDim KeysArray(Length - 1)
    Index = 0
    For Each key In KeysObject
        KeysArray(Index) = key
        Index = Index + 1
    Next
    GetKeys = KeysArray
End Function
