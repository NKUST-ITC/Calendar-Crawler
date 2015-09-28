Imports HtmlAgilityPack

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim doc As New HtmlDocument()
        doc.LoadHtml(TextBox1.Text.Trim)
        Dim node As HtmlNode = doc.DocumentNode

        Debug.Print("----------------------------------------------------------------------")

        Dim _table As HtmlNodeCollection = node.SelectNodes("//*[@id=""ContentPlaceHolder1_TabContainer1_TabPanelBsGov_div_ContentBsGov""]/table/tr")

        For i = 1 To _table.Count - 1
            Dim _scheduleTable As String = node.SelectNodes("//*[@id=""ContentPlaceHolder1_TabContainer1_TabPanelBsGov_TDContent_" & (i - 1) & "9""]")(0).InnerText
            If (Not _scheduleTable.Contains("(教)")) Then
                Continue For
            End If
            Debug.Print(node.SelectNodes("//*[@id=""ContentPlaceHolder1_TabContainer1_TabPanelBsGov_TDContent_" & (i - 1) & "1""]")(0).InnerText)
            For j = 1 To _scheduleTable.Split("*(教)").Count - 1
                If (j = _scheduleTable.Split("*(教)").Count - 1) Then
                    Debug.Print("""" + _scheduleTable.Split("*(教)")(j).Substring(3).Replace(")-", ") ") + """")
                Else
                    Debug.Print("""" + _scheduleTable.Split("*(教)")(j).Substring(3).Replace(")-", ") ") + """,")
                End If
            Next
        Next
    End Sub
End Class
