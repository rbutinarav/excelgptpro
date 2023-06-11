Attribute VB_Name = "Module1"
Public Function GetSetupValue(parameterName As String) As String
    Dim setupSheet As Worksheet
    Set setupSheet = ThisWorkbook.Sheets("Setup")
    
    Dim cell As Range
    For Each cell In setupSheet.Range("A:A")
        If cell.Value = parameterName Then
            GetSetupValue = cell.Offset(0, 1).Value
            Exit Function
        End If
    Next cell
    
    ' If the parameter is not found, return an empty string
    GetSetupValue = ""
End Function

Function JsonEscape(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, "/", "\/")
    s = Replace(s, Chr(8), "\b")
    s = Replace(s, Chr(12), "\f")
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(13), "\r")
    s = Replace(s, Chr(9), "\t")
    JsonEscape = s
End Function


Public Function OpenAI(prompt As String, Optional engine As String, Optional temperature As String, Optional max_tokens As String) As String
    ' Get default parameters from the Setup sheet if not provided
    If engine = "" Then engine = GetSetupValue("DEFAULT_ENGINE")
    If temperature = "" Then temperature = CDbl(GetSetupValue("DEFAULT_TEMPERATURE"))
    If max_tokens = "" Then max_tokens = CInt(GetSetupValue("DEFAULT_MAX_TOKENS"))

    Dim api_key As String: api_key = GetSetupValue("AZURE_OPENAI_KEY")
    Dim api_version As String: api_version = GetSetupValue("AZURE_API_VERSION")
    Dim api_endpoint As String: api_endpoint = GetSetupValue("AZURE_OPENAI_ENDPOINT")

    ' Prepare the API request
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Construct the URL for the request
    Dim url As String
    url = api_endpoint & "/openai/deployments/" & engine & "/completions?api-version=" & api_version

    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "api-key", api_key
    
    ' Construct the data to send in the request
    Dim data As String
    data = "{""prompt"": """ & JsonEscape(prompt) & """, ""max_tokens"": " & max_tokens & ", ""temperature"": " & temperature & "}"
    
    xmlhttp.send (data)
    
    ' Parse the response
    Dim response As String
    response = xmlhttp.responseText
    
    ' Extract the text from the response
    Dim startPos As Integer: startPos = InStr(response, "text"":""") + 7
    
    Dim endPos As Integer: endPos = InStr(startPos, response, """,""index") - 2
    Dim response_text As String: response_text = Mid(response, startPos, endPos - startPos + 1)
    
    ' Remove leading and trailing white spaces and new lines
    response_text = Trim(response_text)
    response_text = Replace(response_text, vbNewLine, "")
    
    Do While Left(response_text, 4) = "\n\n"
    response_text = Mid(response_text, 5)
    Loop
    
    OpenAI = response_text
End Function




Public Function RangeToJSON(rng As Range) As String
    Dim cell As Range
    Dim json As String
    
    json = "{"
    For Each cell In rng
        json = json & """" & cell.Value & """: """ & cell.Offset(0, 1).Value & ""","
    Next cell
    json = Left(json, Len(json) - 1) ' Remove the trailing comma
    json = json & "}"
    
    RangeToJSON = json
End Function

Public Function TableRangeToJSON(rng As Range) As String
    Dim row As Range
    Dim cell As Range
    Dim i As Integer
    Dim headers() As String
    Dim json As String

    ' Get headers from the first row
    ReDim headers(rng.Columns.Count - 1)
    For Each cell In rng.Rows(1).Cells
        headers(cell.Column - rng.Column) = cell.Value
    Next cell

    ' Start JSON array
    json = "["
    
    ' Loop over each row
    For Each row In rng.Offset(1, 0).Resize(rng.Rows.Count - 1).Rows
        ' Start a new JSON object
        json = json & "{"
        
        ' Add each cell in the row to the JSON object
        For i = 0 To UBound(headers)
            json = json & """" & headers(i) & """: """ & row.Cells(1, i + 1).Value & ""","
        Next i
        
        ' Remove trailing comma and close the JSON object
        json = Left(json, Len(json) - 1) & "},"
    Next row

    ' Remove trailing comma and close the JSON array
    json = Left(json, Len(json) - 1) & "]"
    
    TableRangeToJSON = json
End Function

