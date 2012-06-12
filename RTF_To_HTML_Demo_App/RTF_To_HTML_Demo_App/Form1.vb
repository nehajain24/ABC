Public Class Form1

    Public Function sRTF_To_HTML(ByVal sRTF As String) As String

        Dim MyWord As Microsoft.Office.Interop.Word.Application
        Dim oDoNotSaveChanges As Object = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges
        Dim sReturnString As String = ""
        Dim sConvertedString As String = ""
        Try
            MyWord = CreateObject("Word.application")
            MyWord.Visible = False
            MyWord.Documents.Add()


            Dim doRTF As New System.Windows.Forms.DataObject
            doRTF.SetData("Rich Text Format", sRTF)
            Clipboard.SetDataObject(doRTF)
            MyWord.Windows(1).Selection.Paste()
            MyWord.Windows(1).Selection.WholeStory()
            MyWord.Windows(1).Selection.Copy()
            sConvertedString = Clipboard.GetData(System.Windows.Forms.DataFormats.Html)
            'Remove some leading text that shows up in the email
            sConvertedString = sConvertedString.Substring(sConvertedString.IndexOf("<html"))
            'Also remove multiple Â characters that somehow got inserted 
            sConvertedString = sConvertedString.Replace("Â", "")
            sReturnString = sConvertedString
            If Not MyWord Is Nothing Then
                MyWord.Quit(oDoNotSaveChanges)
                MyWord = Nothing
            End If
        Catch ex As Exception
            If Not MyWord Is Nothing Then
                MyWord.Quit(oDoNotSaveChanges)
                MyWord = Nothing
            End If
            MsgBox("Error converting Rich Text to HTML")
        End Try
        Return sReturnString
    End Function


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.TextBox1.Text = Me.sRTF_To_HTML(Me.RichTextBox1.Rtf)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        'neha jain
    End Sub
End Class
