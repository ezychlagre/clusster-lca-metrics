' Const urlGetLcaMetrics As String = "http://91.134.23.173:5020/lca-api/"
Const urlGetLcaMetrics As String = "http://localhost:45021"
Dim sessionId As String



Function Ping()

' Dim url As String
Dim httpRequest As Object
Dim requestBody As String
Dim response As String

url = urlGetLcaMetrics + "/ping"

' requestBody = "{""key1"": ""value1"", ""key2"": ""value2""}"  ' Corps JSON

' Créer la requête HTTP

Set httpRequest = CreateObject("MSXML2.XMLHTTP")

' Ouvrir la connexion et préparer l'en-tête

httpRequest.Open "GET", url, False
httpRequest.setRequestHeader "accept:", "application/json"

' Envoyer le corps
' httpRequest.send requestBody
httpRequest.send

' Traiter la réponse

If httpRequest.status = 200 Then
    response = httpRequest.responseText
    MsgBox "Réponse (200) : " & response
Else
    MsgBox "Échec : code " & httpRequest.status & " - " & httpRequest.statusText
End If

End Function

Function create_session()

Dim url As String
Dim httpRequest As Object
Dim requestBody As String
Dim response As String

url = urlGetLcaMetrics + "/create_session"


    part_type = ThisWorkbook.Worksheets("Sheet1").Cells(2, 1)
    machine_id = ThisWorkbook.Worksheets("Sheet1").Cells(2, 2)
    escription = ThisWorkbook.Worksheets("Sheet1").Cells(2, 3)
' the peak_power is not used for CLUSSTER project
quantity = ThisWorkbook.Worksheets("Sheet1").Cells(2, 5)
Name = ThisWorkbook.Worksheets("Sheet1").Cells(2, 6)
quantity = ThisWorkbook.Worksheets("Sheet1").Cells(2, 5)


requestBody = "{ ""parts"": [ { ""part_type"": """ & part_type & """" & _
                ", ""machine_id"": """ & machine_id & """" & _
                ", ""description"": """ & escription & """" & _
                ", ""peak_power"": " & "0" & _
                ", ""quantity"": " & quantity & _
                ", ""name"": """ & Name & """" & _
                ", ""die_surface_mm2"": " & "0" & _
                ", ""litho_nm"": " & "0" & _
                ", ""size_gb"": " & "0" & _
                ", ""technology"": """ & "string" & """" & _
                ", ""casing"": """ & "string" & """" & _
                "} ] }"

MsgBox "body: " & requestBody

' Créer la requête HTTP

Set httpRequest = CreateObject("MSXML2.XMLHTTP")

' Ouvrir la connexion et préparer l'en-tête

httpRequest.Open "POST", url, False
httpRequest.setRequestHeader "accept", "application/json"
httpRequest.setRequestHeader "content-Type", "application/json"

' Envoyer le corps
httpRequest.send requestBody
' httpRequest.send

' Traiter la réponse

If httpRequest.status = 200 Then
    response = httpRequest.responseText
    MsgBox "Réponse (200) : " & response
    
    ' get the session id
    parts = Split(response, """")
    sessionId = parts(3)
    MsgBox "Session Id: " & sessionId
    
    
Else
    MsgBox "Échec : code " & httpRequest.status & " - " & httpRequest.statusText
End If


End Function

Function PreCheck()

' Dim url As String
Dim httpRequest As Object
Dim requestBody As String
Dim response As String

url = urlGetLcaMetrics + "/precheck/" + sessionId

Set httpRequest = CreateObject("MSXML2.XMLHTTP")

' Ouvrir la connexion et préparer l'en-tête

httpRequest.Open "GET", url, False
httpRequest.setRequestHeader "accept:", "application/json"

' Envoyer la requete
httpRequest.send

' Traiter la réponse

If httpRequest.status = 200 Then
    response = httpRequest.responseText
    MsgBox "Réponse (200) : " & response
Else
    MsgBox "Échec : code " & httpRequest.status & " - " & httpRequest.statusText
End If

End Function

Function launch()

' Dim url As String
Dim httpRequest As Object
Dim requestBody As String
Dim response As String

url = urlGetLcaMetrics + "/launch/" + sessionId

Set httpRequest = CreateObject("MSXML2.XMLHTTP")

' Ouvrir la connexion et préparer l'en-tête

httpRequest.Open "GET", url, False
httpRequest.setRequestHeader "accept:", "application/json"

' Envoyer la requete
httpRequest.send

' Traiter la réponse

If httpRequest.status = 200 Then
    response = httpRequest.responseText
    MsgBox "Réponse (200) : " & response
Else
    MsgBox "Échec : code " & httpRequest.status & " - " & httpRequest.statusText
End If

End Function

Function status()

' Dim url As String
Dim httpRequest As Object
Dim requestBody As String
Dim response As String

url = urlGetLcaMetrics + "/status/" + sessionId

Set httpRequest = CreateObject("MSXML2.XMLHTTP")

' Ouvrir la connexion et préparer l'en-tête

httpRequest.Open "GET", url, False
httpRequest.setRequestHeader "accept:", "application/json"

' Envoyer la requete
httpRequest.send

' Traiter la réponse

If httpRequest.status = 200 Then
    response = httpRequest.responseText
    MsgBox "Réponse (200) : " & response
Else
    MsgBox "Échec : code " & httpRequest.status & " - " & httpRequest.statusText
End If

End Function






