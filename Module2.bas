Attribute VB_Name = "Module2"
Option Compare Database

Public Sub ExportHTML()
Dim adoCN As adoDB.Connection
Dim adoRS As adoDB.Recordset
Dim i As Integer
Dim SQL As String
Dim filename As String

Set adoCN = CurrentProject.Connection

Set adoRS = New adoDB.Recordset
SQL = "SELECT * FROM T_月間一覧表"

adoRS.Open SQL, adoCN, adOpenKeyset, adLockOptimistic

Dim ado As Object
Set ado = CreateObject("ADODB.Stream")

ado.Charset = "UTF-8"

ado.Open

ado.WriteText "<!DOCTYPE html>", adWriteLine
ado.WriteText "<html lang=""ja"">", adWriteLine
ado.WriteText "<head>", adWriteLine
ado.WriteText "<meta charset=""utf-8"">", adWriteLine
ado.WriteText "<title>ExportHTML</title>", adWriteLine
ado.WriteText "<link ref=""stylesheet"" href=""css/styles.css"">", adWriteLine
ado.WriteText "</head>", adWriteLine
ado.WriteText "<body>", adWriteLine
ado.WriteText "<h1>ExportHTML</h1>", adWriteLine
Do Until adoRS.EOF
    ado.WriteText "<p>" & adoRS("日付") & " " & adoRS("ジョブ名") & "</p>", adWriteLine
    adoRS.MoveNext
Loop
ado.WriteText "</body>", adWriteLine
ado.WriteText "</html>", adowriteline

filename = "exporthtml"
ado.SaveToFile "C:\Users\inner.WINDOWS10\Desktop\" & filename & ".html", 2

ado.Close
adoCN.Close

Set adoCN = Nothing

End Sub
