Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Xml
Public Class pdfWriter
  Public Shared Function generatePDF(ByVal PathFileName As String, ByVal OutputPath As String) As Boolean
    Dim mRet As Boolean = True
    Dim extn As String = IO.Path.GetExtension(PathFileName).ToUpper
    Select Case extn
      Case ".XLS", ".XLSX", ".CSV"
        mRet = generateXLPDF({PathFileName}.ToList, OutputPath)
      Case ".DOC", ".DOCX", ".TXT"
        mRet = generateDOCPDF({PathFileName}.ToList, OutputPath)
      Case ".PPT", ".PPTX"
        mRet = generatePPTPDF({PathFileName}.ToList, OutputPath)
      Case ".PDF"
        mRet = generatePDFPDF({PathFileName}.ToList, OutputPath)
    End Select
    Return mRet
  End Function

  Public Shared Function generateXLPDF(ByVal PathFileList As List(Of String), ByVal OutputPath As String) As Boolean
    Dim mRet As Boolean = True
    Dim xlAp As Microsoft.Office.Interop.Excel.ApplicationClass = Nothing
    Dim xlWb As Workbook = Nothing
    Dim xlWs As Worksheet = Nothing
    Try
      xlAp = New Microsoft.Office.Interop.Excel.ApplicationClass
      For Each lstFile As String In PathFileList
        Dim FileName As String = IO.Path.GetFileName(lstFile)
        Dim Extn As String = IO.Path.GetExtension(FileName).ToUpper
        Select Case Extn
          Case ".XLS", ".XLSX", ".CSV"
          Case Else
            Continue For
        End Select
        Dim OutputFile As String = OutputPath & "\" & IO.Path.GetFileNameWithoutExtension(FileName) & ".PDF"
        Try
          xlWb = xlAp.Workbooks.Open(lstFile)
          If xlWb IsNot Nothing Then
            If IO.File.Exists(OutputFile) Then
              IO.File.Delete(OutputFile)
            End If
            'select all sheets, [pending -check if there are any line written]
            'xlWb.Sheets({1, 2}).Select()
            xlAp.ActiveSheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, OutputFile, XlFixedFormatQuality.xlQualityStandard, True, True, OpenAfterPublish:=False)
            xlWb.Close(False)
            xlWb = Nothing
          End If
        Catch ex As Exception
          mRet = False
        End Try
      Next
      xlAp.Quit()
    Catch ex As Exception
      mRet = False
    End Try
    Return mRet
  End Function

  Public Shared Function generateDOCPDF(ByVal PathFileList As List(Of String), ByVal OutputPath As String) As Boolean
    Dim mRet As Boolean = True
    Dim docAp As Microsoft.Office.Interop.Word.ApplicationClass
    Dim docWD As Document = Nothing
    Try
      docAp = New Microsoft.Office.Interop.Word.ApplicationClass
      For Each lstFile As String In PathFileList
        Dim FileName As String = IO.Path.GetFileName(lstFile)
        Dim Extn As String = IO.Path.GetExtension(FileName).ToUpper
        Select Case Extn
          Case ".DOC", ".DOCX", ".TXT"
          Case Else
            Continue For
        End Select
        Dim OutputFile As String = OutputPath & "\" & IO.Path.GetFileNameWithoutExtension(FileName) & ".PDF"
        Try
          docWD = docAp.Documents.Open(lstFile)
          If docWD IsNot Nothing Then
            If IO.File.Exists(OutputFile) Then
              IO.File.Delete(OutputFile)
            End If
            docWD.ExportAsFixedFormat(OutputFile, WdExportFormat.wdExportFormatPDF) ', False, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 1, 100, WdExportItem.wdExportDocumentContent, False, True, WdExportCreateBookmarks.wdExportCreateNoBookmarks, True, True, False, Nothing)
            docAp.ActiveDocument.Close(False)
            docWD = Nothing
          End If
        Catch ex As Exception
          mRet = False
        End Try
      Next
      docAp.Quit()
    Catch ex As Exception
      mRet = False
    End Try
    Return mRet
  End Function

  Public Shared Function generatePPTPDF(ByVal PathFileList As List(Of String), ByVal OutputPath As String) As Boolean
    Dim mRet As Boolean = True
    Dim pptAp As Microsoft.Office.Interop.PowerPoint.ApplicationClass = Nothing
    Dim pptPr As Presentation = Nothing
    Try
      pptAp = New Microsoft.Office.Interop.PowerPoint.ApplicationClass
      For Each lstFile As String In PathFileList
        Dim FileName As String = IO.Path.GetFileName(lstFile)
        Dim Extn As String = IO.Path.GetExtension(FileName).ToUpper
        Select Case Extn
          Case ".PPT", ".PPTX"
          Case Else
            Continue For
        End Select
        Dim OutputFile As String = OutputPath & "\" & IO.Path.GetFileNameWithoutExtension(FileName) & ".PDF"
        Try
          pptPr = pptAp.Presentations.Open(lstFile, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse)
          If pptPr IsNot Nothing Then
            If IO.File.Exists(OutputFile) Then
              IO.File.Delete(OutputFile)
            End If
            pptPr.SaveAs(OutputFile, PpFixedFormatType.ppFixedFormatTypePDF)  ', PpFixedFormatIntent.ppFixedFormatIntentScreen, MsoTriState.msoTriStateMixed, PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst, PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoFalse, Nothing, PpPrintRangeType.ppPrintAll, "", False, True, True, True, False, Type.Missing)
            pptAp.ActivePresentation.Close()
            pptPr.Close()
            pptPr = Nothing
          End If
        Catch ex As Exception
          mRet = False
        End Try
      Next
      pptAp.Quit()
    Catch ex As Exception
      mRet = False
    End Try
    Return mRet
  End Function
  Public Shared Function generatePDFPDF(ByVal PathFileList As List(Of String), ByVal OutputPath As String) As Boolean
    Dim mRet As Boolean = True
    Try
      For Each lstFile As String In PathFileList
        Dim FileName As String = IO.Path.GetFileName(lstFile)
        Dim Extn As String = IO.Path.GetExtension(FileName).ToUpper
        Select Case Extn
          Case ".PDF"
          Case Else
            Continue For
        End Select
        Dim OutputFile As String = OutputPath & "\" & FileName
        Try
          If IO.File.Exists(OutputFile) Then
            IO.File.Delete(OutputFile)
          End If
          IO.File.Copy(lstFile, OutputFile)
        Catch ex As Exception
          mRet = False
        End Try
      Next
    Catch ex As Exception
      mRet = False
    End Try
    Return mRet
  End Function
End Class
