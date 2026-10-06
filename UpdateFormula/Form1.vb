Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Runtime.InteropServices

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim folderPath As String = Application.StartupPath

        Dim files() As String = Directory.GetFiles(folderPath, "*.xlsx")

        If files.Length = 0 Then
            MessageBox.Show("No Excel files found.")
            Return
        End If

        Dim excelApp As New Excel.Application
        excelApp.Visible = False
        excelApp.DisplayAlerts = False

        Dim updatedCount As Integer = 0

        Try
            For Each filePath As String In files

                ' Skip temporary Excel files
                If Path.GetFileName(filePath).StartsWith("~$") Then
                    Continue For
                End If

                Dim workbook As Excel.Workbook = Nothing
                Dim worksheet As Excel.Worksheet = Nothing

                Try
                    workbook = excelApp.Workbooks.Open(filePath)

                    ' Change this if you need a specific sheet
                    worksheet = CType(workbook.Worksheets(1), Excel.Worksheet)

                    ' Set formulas
                    worksheet.Range("D15").Formula = "=(D24-D16)/1.05"
                    worksheet.Range("D22").Formula = "=D15*0.05"

                    ' Save
                    workbook.Save()

                    updatedCount += 1

                Catch ex As Exception
                    MessageBox.Show(
                        "Error updating:" & vbCrLf &
                        filePath & vbCrLf & vbCrLf &
                        ex.Message,
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    )

                Finally
                    If workbook IsNot Nothing Then
                        workbook.Close(SaveChanges:=False)
                        Marshal.ReleaseComObject(workbook)
                    End If

                    If worksheet IsNot Nothing Then
                        Marshal.ReleaseComObject(worksheet)
                    End If
                End Try

            Next

        Finally
            excelApp.Quit()
            Marshal.ReleaseComObject(excelApp)
        End Try

        MessageBox.Show(
            updatedCount & " Excel file(s) updated successfully.",
            "Complete",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )

    End Sub

End Class
