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

Function TrimNewlines(s As String) As String
    While Left(s, 2) = vbCrLf
        s = Mid(s, 3)
    Wend

    While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Wend

    TrimNewlines = s
End Function


Function ReplaceMultipleNewlines(s As String) As String
    ' Replace multiple newlines with a single newline
    Do While InStr(s, vbCrLf & vbCrLf) > 0
        s = Replace(s, vbCrLf & vbCrLf, vbCrLf)
    Loop

    ' Trim leading and trailing newlines or whitespaces
    Do While Left(s, 2) = vbCrLf Or Left(s, 1) = " " Or Left(s, 1) = Chr(9)
        s = Mid(s, IIf(Left(s, 2) = vbCrLf, 3, 2))
    Loop

    Do While Right(s, 2) = vbCrLf Or Right(s, 1) = " " Or Right(s, 1) = Chr(9)
        s = Left(s, Len(s) - IIf(Right(s, 2) = vbCrLf, 2, 1))
    Loop

    ReplaceMultipleNewlines = s
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
    On Error GoTo ErrorHandler ' Add this line to enable error handling
    ' Get default parameters from the Setup sheet if not provided
    If engine = "" Then engine = GetSetupValue("DEFAULT_ENGINE")
    If temperature = "" Then temperature = CDbl(GetSetupValue("DEFAULT_TEMPERATURE"))
    If max_tokens = "" Then max_tokens = CInt(GetSetupValue("DEFAULT_MAX_TOKENS"))

    Dim api_key As String
    Dim api_version As String
    Dim api_endpoint As String
    Dim api_type As String: api_type = GetSetupValue("API_TYPE")

    ' Check which API to use and set the key, version and endpoint accordingly
    If api_type = "Azure" Then
        api_key = GetSetupValue("AZURE_OPENAI_KEY")
        api_version = GetSetupValue("AZURE_API_VERSION")
        api_endpoint = GetSetupValue("AZURE_OPENAI_ENDPOINT")
    ElseIf api_type = "OpenAI" Then
        api_key = GetSetupValue("OPENAI_KEY")
        api_version = "" ' OpenAI does not use a version parameter
        If engine = "gpt-4" Or engine = "gpt-3.5-turbo" Or engine = "gpt-3.5-turbo-16k" Then
            api_endpoint = "https://api.openai.com/v1/chat/completions"
            
        Else
            api_endpoint = "https://api.openai.com/v1/engines/" & engine & "/completions"
        End If
    Else
        ' Invalid API type
        OpenAI = "Invalid API type"
        Exit Function
    End If

    ' Prepare the API request
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

    ' Construct the URL for the request
    Dim url As String
    If api_type = "Azure" Then
        url = api_endpoint & "/openai/deployments/" & engine & "/completions?api-version=" & api_version
    ElseIf api_type = "OpenAI" Then
        url = api_endpoint
    End If

    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "Content-Type", "application/json"

    ' Set the API key in the headers
    If api_type = "Azure" Then
        xmlhttp.setRequestHeader "api-key", api_key
    ElseIf api_type = "OpenAI" Then
        xmlhttp.setRequestHeader "Authorization", "Bearer " & api_key
    End If

    ' Construct the data to send in the request
    Dim data As String
    prompt = JsonEscape(prompt)

    If engine = "gpt-4" Or engine = "gpt-3.5-turbo" Or engine = "gpt-3.5-turbo-16k" Then
        'For chat models, construct the payload according to chat models requirements
        data = "{""model"": """ & engine & """, ""max_tokens"": " & max_tokens & ", ""temperature"": " & temperature & ", ""messages"": [{""role"": ""system"", ""content"": ""You are a helpful assistant.""},{""role"": ""user"", ""content"": """ & prompt & """}]}"
    Else
        'For completion models, construct the payload according to completion models requirements
        data = "{""prompt"": """ & prompt & """, ""max_tokens"": " & max_tokens & ", ""temperature"": " & temperature & "}"
    End If


    xmlhttp.send (data)

    ' Parse the response
    Dim response As String
    response = xmlhttp.responseText

    ' Extract the text from the response
    Dim startPos As Integer
    Dim endPos As Integer
    Dim response_text As String
    If api_type = "Azure" Then
        startPos = InStr(response, "text"":""") + 7
        endPos = InStr(startPos, response, """,""index") - 1
        response_text = Mid(response, startPos, endPos - startPos + 1)
    ElseIf api_type = "OpenAI" Then
        If engine = "gpt-4" Or engine = "gpt-3.5-turbo" Or engine = "gpt-3.5-turbo-16k" Then
            'OpenAI's chat models response structure might be different, adjust as needed
            startPos = InStr(response, """content"": """) + 12
            endPos = InStr(startPos, response, """") - 1
            response_text = Mid(response, startPos, endPos - startPos + 1)
        Else
            'OpenAI's completion models response structure might be different, adjust as needed
            startPos = InStr(response, """text"": """) + 9
            endPos = InStr(startPos, response, """") - 1
            response_text = Mid(response, startPos, endPos - startPos + 1)
        End If
    End If

    ' Convert JSON newlines to VBA newlines
    response_text = Replace(response_text, "\r\n", vbCrLf)
    response_text = Replace(response_text, "\n", vbCrLf)

    response_text = ReplaceMultipleNewlines(response_text)

    OpenAI = response_text
    
    Exit Function
    
ErrorHandler:
    OpenAI = "OpenAI call error"
    
End Function

Public Function MatrixRangeToJSON(tableRange As Range, entityHeadersRange As Range, entityName As String) As String
    Dim row As Range
    Dim cell As Range
    Dim i As Integer
    Dim entityHeaders() As String
    Dim json As String
    Dim startEntityCol As Integer
    Dim endEntityCol As Integer
    Dim detailsCol As Integer

    ' Get headers from the entityHeadersRange
    ReDim entityHeaders(entityHeadersRange.Columns.Count - 1)
    For Each cell In entityHeadersRange.Cells
        entityHeaders(cell.Column - entityHeadersRange.Column) = cell.Value
    Next cell

    ' Start JSON array
    json = "["

    ' Define start and end of entity columns
    startEntityCol = entityHeadersRange.Column - tableRange.Column + 1
    endEntityCol = startEntityCol + entityHeadersRange.Columns.Count - 1

    ' Define start of details columns
    If startEntityCol = 1 Then
        detailsCol = endEntityCol + 1
    Else
        detailsCol = 1
    End If

    ' Loop over each column (entity)
    For i = 0 To UBound(entityHeaders)
        ' Start a new JSON object for each entity
        json = json & "{""" & entityName & """: {"

        ' Add entity name
        json = json & """Name"": """ & entityHeaders(i) & """, "

        ' Add details from the same column
        For Each row In tableRange.Offset(1, 0).Resize(tableRange.Rows.Count - 1).Rows
            ' Add detail only if the cell is not empty
            If Trim(row.Cells(1, startEntityCol + i).Value) <> "" Then
                json = json & """" & tableRange.Cells(row.row, detailsCol).Value & """: """ & row.Cells(1, startEntityCol + i).Value & """, "
            End If
        Next row

        ' Remove trailing comma and close the JSON object for the entity
        json = Left(json, Len(json) - 2) & "}},"
    Next i

    ' Remove trailing comma and close the JSON array
    json = Left(json, Len(json) - 1) & "]"

    MatrixRangeToJSON = json
End Function


Public Function SimpleRangeToJSON(rng As Range, Optional hasHeaders As Boolean = False) As String
    Dim row As Range
    Dim cell As Range
    Dim i As Integer
    Dim headers() As String
    Dim json As String

    ' Get headers from the first row if they exist, else generate generic headers
    ReDim headers(rng.Columns.Count - 1)
    If hasHeaders Then
        For Each cell In rng.Rows(1).Cells
            headers(cell.Column - rng.Column) = cell.Value
        Next cell
    Else
        For i = 0 To UBound(headers)
            headers(i) = "Field" & (i + 1)
        Next i
    End If

    ' Start JSON array
    json = "["

    ' Determine the start row based on whether or not there are headers
    Dim startRow As Integer
    If hasHeaders Then
        startRow = 2
    Else
        startRow = 1
    End If

    ' Loop over each row
    For Each row In rng.Offset(startRow - 1, 0).Resize(rng.Rows.Count - (startRow - 1)).Rows
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

    SimpleRangeToJSON = json
End Function


Public Function SimpleRangeToJSONL(rng As Range, Optional hasHeaders As Boolean = True) As String
    Dim row As Range
    Dim cell As Range
    Dim i As Integer
    Dim headers() As String
    Dim json As String

    ' Get headers from the first row if they exist, else generate generic headers
    ReDim headers(rng.Columns.Count - 1)
    If hasHeaders Then
        For Each cell In rng.Rows(1).Cells
            headers(cell.Column - rng.Column) = cell.Value
        Next cell
    Else
        For i = 0 To UBound(headers)
            headers(i) = "Field" & (i + 1)
        Next i
    End If

    ' Start JSONL string
    json = ""

    ' Determine the start row based on whether or not there are headers
    Dim startRow As Integer
    If hasHeaders Then
        startRow = 2
    Else
        startRow = 1
    End If

    ' Loop over each row
    For Each row In rng.Offset(startRow - 1, 0).Resize(rng.Rows.Count - (startRow - 1)).Rows
        ' Start a new JSON object
        json = json & "{"

        ' Add each cell in the row to the JSON object
        For i = 0 To UBound(headers)
            json = json & """" & headers(i) & """: """ & row.Cells(1, i + 1).Value & ""","
        Next i

        ' Remove trailing comma and close the JSON object, then add newline
        json = Left(json, Len(json) - 1) & "}" & vbNewLine
    Next row

    ' Remove trailing newline
    json = Left(json, Len(json) - 2)

    SimpleRangeToJSONL = json
End Function


Public Function RangeToDelimitedText(rng As Range, Optional delimiter As String = "|", Optional replacement As String = "/") As String
    Dim row As Range
    Dim cell As Range
    Dim txt As String
    
    ' Loop over each row
    For Each row In rng.Rows
        ' Loop over each cell in the row
        For Each cell In row.Cells
            ' Replace any instances of the delimiter in the cell's value with the replacement character
            Dim cellValue As String
            cellValue = Replace(cell.Value, delimiter, replacement)
            ' Add the cell's value to the string, followed by the delimiter
            txt = txt & cellValue & delimiter
        Next cell
        ' Replace the trailing delimiter with a newline character
        txt = Left(txt, Len(txt) - Len(delimiter)) & vbNewLine
    Next row

    ' Remove the trailing newline character
    txt = Left(txt, Len(txt) - 1)

    RangeToDelimitedText = txt
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
