Imports HtmlAgilityPack
Imports Newtonsoft.Json.Linq
Imports System.IO

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim doc As New HtmlDocument()
            doc.LoadHtml(TextBox1.Text.Trim)
            Dim node As HtmlNode = doc.DocumentNode

            Debug.Print("----------------------------------------------------------------------")

            Dim _table As HtmlNodeCollection = node.SelectNodes("//*[@id=""ContentPlaceHolder1_TabContainer1_TabPanelBsGov_div_ContentBsGov""]/table/tr")

            Dim jsonArray As JArray = New JArray
            Dim yms As String = node.SelectSingleNode("//*[@id=""ContentPlaceHolder1_TabContainer1_TabPanelBsGov_DropDownListForBsGovSchYearSem""]/option[@selected='selected']").Attributes.Item("value").Value.Replace("-", "0")

            For i = 1 To _table.Count - 1
                Dim _scheduleTable As String = node.SelectSingleNode("//*[@id=""ContentPlaceHolder1_TabContainer1_TabPanelBsGov_TDContent_" & (i - 1) & "9""]").InnerText
                If (Not _scheduleTable.Contains("(教)")) Then
                    Continue For
                End If

                Dim _jsonObject As JObject = New JObject
                Dim _jsonArray As JArray = New JArray

                Dim _week As String = node.SelectSingleNode("//*[@id=""ContentPlaceHolder1_TabContainer1_TabPanelBsGov_TDContent_" & (i - 1) & "1""]").InnerText
                If Not (_week.Equals("寒") Or _week.Equals("暑") Or _week.Equals("預備週")) Then
                    _week = "第" + _week + "週"
                End If

                ' Check JArray has contain JObject
                Dim _index As Integer = -1
                For j = 0 To jsonArray.Count - 1
                    If (jsonArray(j)("week").ToString.Equals(_week)) Then
                        _jsonArray = jsonArray(j)("events")
                        _index = j
                        Exit For
                    End If
                Next

                For j = 1 To _scheduleTable.Split("*(教)").Count - 1
                    Dim _event As String = _scheduleTable.Split("*(教)")(j).Substring(3).Replace(")-", ") ")
                    Dim _splitIndex As Integer = _event.IndexOf(" ")
                    _event = _event.Substring(0, _splitIndex).Replace("-", " ~ ") + _event.Substring(_splitIndex)
                    _jsonArray.Add(_event)
                Next

                ' Only add JObject when key not existed
                If (_index = -1) Then
                    _jsonObject.Add(New JProperty("week", _week))
                    _jsonObject.Add(New JProperty("events", _jsonArray))
                    jsonArray.Add(_jsonObject)
                End If
            Next
            Debug.Print(jsonArray.ToString)
            SaveFile(Application.StartupPath & "\" & yms & ".json", jsonArray.ToString)
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.Critical, "Error [Html]")
        End Try
    End Sub

    Private Sub SaveFile(ByVal strPathName As String, ByVal strContent As String)
        Try
            Dim swWriter As StreamWriter = New StreamWriter(strPathName, False, System.Text.Encoding.Unicode)
            swWriter.Write(strContent)
            swWriter.Close()
            MsgBox("已輸出至：" + vbCrLf + strPathName, MsgBoxStyle.Information, "Success")
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.Critical, "Error [Save File]")
        End Try
    End Sub
End Class
