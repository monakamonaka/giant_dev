Imports System.Reflection.Metadata

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim cmds As String() = System.Environment.GetCommandLineArgs()
            'コマンドライン引数を列挙する
            MsgBox(cmds(0))
            MsgBox(cmds(1))
            CNVPDF(cmds(0), cmds(1))

        Catch ex As Exception

        End Try

        '  End

    End Sub
    Private Sub CNVPDF(F1 As String, F2 As String)
        Try


            Select Case LCase(System.IO.Path.GetExtension(F1))
                Case ".doc", ".docx", ".dotm"
                    Const wdExportFormatPDF = 17
                    Const wdExportOptimizeForPrint = 0
                    Const wdExportAllDocument = 0
                    Const wdExportDocumentContent = 0
                    Const wdExportCreateWordBookmarks = 2
                    With CreateObject("Word.Application")
                        .Visible = True
                        .Documents.Open(F1)
                        .Documents.ExportAsFixedFormat(F2, wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportAllDocument, , , wdExportDocumentContent, False, False, wdExportCreateWordBookmarks, True, True, False)
                        .Close

                    End With

                Case ".xls", ".xlsx", ".xlsm"

                    With CreateObject("Excel.Application")
                        .Visible = True
                        .Workbooks.open(F1)
                        .ActiveSheet.ExportAsFixedFormat(0, F2, 0, 1, 0,,, 0)
                        .ActiveWorkbook.Close
                        .Close
                        .Quit

                    End With

                Case ".ppt", ".pptx", ".pptm"
                    Const ppSaveAsPDF = 32

                    With CreateObject("PowerPoint.Application")
                        .Visible = True
                        .Presentations.Open(F1)
                        .SaveAs(F2, ppSaveAsPDF, -1)
                        .Close
                        .Quit
                    End With
                Case Else

            End Select
        Catch ex As Exception

        End Try

    End Sub

End Class
