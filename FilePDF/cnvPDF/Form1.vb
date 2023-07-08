Imports Microsoft.Office.Interop

Imports System
Imports Microsoft.Office
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CNVPDF("C:\Users\miura\Desktop\DEMO\test.pptx", "C:\Users\miura\Desktop\DEMO\33.pdf")
        'CNVPDF("C:\Users\miura\Desktop\DEMO\辞書ネタ1.xlsx", "C:\Users\miura\Desktop\DEMO\33.pdf")


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

                        .Documents.Open(F1)
                        .Documents.ExportAsFixedFormat(F2, wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportAllDocument, , , wdExportDocumentContent, False, False, wdExportCreateWordBookmarks, True, True, False)
                        .Close

                    End With

                Case ".xls", ".xlsx", ".xlsm"

                    With CreateObject("Excel.Application")

                        .Workbooks.open(F1)
                        .ActiveSheet.ExportAsFixedFormat(0, F2, 0, 1, 0,,, 0)
                        .ActiveWorkbook.Close
                        .Close
                        .Quit

                    End With

                Case ".ppt", ".pptx", ".pptm"
                    Const ppSaveAsPDF = 32

                    Dim ppApp As New PowerPoint.Application
                    'ppApp.Visible = True

                    'With CreateObject("PowerPoint.Application")
                    With ppApp
                        .Visible = True
                        .Presentation.Open(F1)
                        .SaveAs(F2, ppSaveAsPDF, -1)
                        .Close
                        .Quit()
                    End With
                Case Else

            End Select
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class
