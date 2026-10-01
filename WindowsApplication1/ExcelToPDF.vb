Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class ExcelToPDF

    Private Sub Button_convert_Click(sender As Object, e As EventArgs) Handles Button_convert.Click
        Dim folderPath As String = Application.StartupPath

        Dim excelFiles() As String = Directory.GetFiles(folderPath, "*.xlsx")

        If excelFiles.Length = 0 Then
            MessageBox.Show("No Excel files found in:" & vbCrLf & folderPath,
                            "No Files",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return
        End If

        Dim excelApp As Excel.Application = Nothing

        Try
            excelApp = New Excel.Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            For Each excelFile As String In excelFiles

                'Ignore temporary Excel files
                If Path.GetFileName(excelFile).StartsWith("~$") Then
                    Continue For
                End If

                Dim workbook As Excel.Workbook = Nothing

                Try
                    workbook = excelApp.Workbooks.Open(excelFile)

                    Dim pdfFile As String =
                        Path.Combine(
                            Path.GetDirectoryName(excelFile),
                            Path.GetFileNameWithoutExtension(excelFile) & ".pdf"
                        )

                    workbook.ExportAsFixedFormat(
                        Excel.XlFixedFormatType.xlTypePDF,
                        pdfFile
                    )

                    workbook.Close(False)
                    workbook = Nothing

                Catch ex As Exception

                    If workbook IsNot Nothing Then
                        workbook.Close(False)
                        workbook = Nothing
                    End If

                    MessageBox.Show(
                        "Failed to convert:" & vbCrLf &
                        Path.GetFileName(excelFile) & vbCrLf & vbCrLf &
                        ex.Message,
                        "Conversion Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    )

                End Try

            Next

            MessageBox.Show(
                "Conversion completed." & vbCrLf & vbCrLf &
                "Folder:" & vbCrLf & folderPath,
                "Done",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

        Catch ex As Exception

            MessageBox.Show(
                "Could not start Microsoft Excel." & vbCrLf & vbCrLf &
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        Finally

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                excelApp = Nothing
            End If

        End Try
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub
End Class
