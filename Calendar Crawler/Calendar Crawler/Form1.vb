Imports HtmlAgilityPack
Imports Newtonsoft.Json.Linq

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim doc As New HtmlDocument()
        doc.LoadHtml(TextBox1.Text.Trim)
        Dim node As HtmlNode = doc.DocumentNode

        Debug.Print("----------------------------------------------------------------------")

        Dim _table As HtmlNodeCollection = node.SelectNodes("//*[@id=""ContentPlaceHolder1_TabContainer1_TabPanelBsGov_div_ContentBsGov""]/table/tr")

        Dim jsonArray As JArray = New JArray

        For i = 1 To _table.Count - 1
            Dim _scheduleTable As String = node.SelectNodes("//*[@id=""ContentPlaceHolder1_TabContainer1_TabPanelBsGov_TDContent_" & (i - 1) & "9""]")(0).InnerText
            If (Not _scheduleTable.Contains("(教)")) Then
                Continue For
            End If

            Dim _jsonObject As JObject = New JObject
            Dim _jsonArray As JArray = New JArray

            Dim _week As String = node.SelectNodes("//*[@id=""ContentPlaceHolder1_TabContainer1_TabPanelBsGov_TDContent_" & (i - 1) & "1""]")(0).InnerText
            If Not (_week.Equals("寒") Or _week.Equals("暑") Or _week.Equals("預備週")) Then
                _week = "第" + _week + "週"
            End If

            _jsonObject.Add(New JProperty("week", _week))
            For j = 1 To _scheduleTable.Split("*(教)").Count - 1
                Dim _event As String = _scheduleTable.Split("*(教)")(j).Substring(3).Replace(")-", ") ")
                Dim _splitIndex As Integer = _event.IndexOf(" ")
                _event = _event.Substring(0, _splitIndex).Replace("-", " ~ ") + _event.Substring(_splitIndex)
                _jsonArray.Add(_event)
            Next
            _jsonObject.Add(New JProperty("events", _jsonArray))
            jsonArray.Add(_jsonObject)
        Next
        Debug.Print(jsonArray.ToString)
    End Sub
End Class
