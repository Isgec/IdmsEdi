Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Environment
Imports JobProcess
Imports EDICommon

Public Class cadUtility
  Implements IMessageFilter
  Implements IDisposable

  Public Property jpConfig As ConfigFile = Nothing

  Private acAppComObj As AcadApplication = Nothing
  Private acDocComObj As AcadDocument = Nothing
  Private SleepTime As Integer = 1000
  Public Property jp As JobProcess.JobProcessor = Nothing

  Public Sub LoadCad()
    Dim FileName As String = jpConfig.StartupPath & "\jpConfig.xml"
    jpConfig = ConfigFile.Serialize(jpConfig, FileName)
    jp.msg("Creating CAD Object.")
    Try
      acAppComObj = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("AutoCAD.Application.20.1"), True), AcadApplication)
    Catch exception As Exception
      jp.msg("Error : " & exception.Message)
      Thread.Sleep(SleepTime)
      jp.msg("2nd Try Creating CAD Object")
      acAppComObj = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("AutoCAD.Application.20.1"), True), AcadApplication)
    End Try
    If acAppComObj IsNot Nothing Then
      Try
        If Not acAppComObj.Application.GetAcadState.IsQuiescent Then
          Thread.Sleep(SleepTime)
        End If
        'acAppComObj.Visible = jpConfig.AutoCADVisible
        With acAppComObj
          .Width = 100
          .Height = 60
          .WindowTop = 1
          .WindowLeft = 1
          .WindowState = AcWindowState.acMin
        End With
      Catch ex As Exception
      End Try
    Else
      jp.msg("Could NOT Create CAD Object.")
    End If
  End Sub
  Public Sub UnloadCad()
    jpConfig = Nothing
    If (Not acDocComObj Is Nothing) Then
      jp.msg("Closing Document.")
      Try
        acAppComObj.ActiveDocument.Close(False)
        jp.msg("Document Closed.")
      Catch ex As Exception
        jp.msg("Error in Closing Document:" & ex.Message)
      End Try
    End If
    If (Not acAppComObj Is Nothing) Then
      jp.msg("Closing CAD Application.")
      Try
        acAppComObj.Quit()
        jp.msg("CAD Application Closed.")
      Catch ex As Exception
        jp.msg("Error Closing CAD Application: " & ex.Message)
      End Try
    End If
    acAppComObj = Nothing
    acDocComObj = Nothing
  End Sub
  Public Function ExtractPDF(ByVal Job As jobFile) As jobFile
    If acAppComObj Is Nothing Then
      LoadCad()
    End If
    If (Not acAppComObj Is Nothing) Then
      Dim FilePath As String = IIf(Job.ForOtherFile, jpConfig.TempFolderPath & "\" & Job.OtherFileName, jpConfig.TempFolderPath & "\" & Job.FileName)
      Dim DLLPath As String = (jpConfig.StartupPath & "\cadConvert.dll").Replace("\", "\\")
      jp.msg("Opening CAD Document:" & FilePath)
      Try
        acDocComObj = acAppComObj.Documents.Open(FilePath, Nothing, Nothing)
        jp.msg("Document Opened.")
      Catch ex As Exception
        jp.msg("Error Opening Document:" & ex.Message)
      End Try
      Thread.Sleep(SleepTime)
      jp.msg("Loading DLL.")
      Try
        'acAppComObj.Visible = jpConfig.AutoCADVisible
        acDocComObj.SendCommand("(command ""SECURELOAD"" ""0"") ")
        acDocComObj.SendCommand("(command ""NETLOAD"" """ & DLLPath & """) ")
        acDocComObj.SendCommand("(command ""SECURELOAD"" ""1"") ")
        'acAppComObj.Visible = jpConfig.AutoCADVisible
        jp.msg("DLL Loaded.")
      Catch ex As Exception
        jp.msg("Error Loading DLL:" & ex.Message)
      End Try
      Job = jobFile.Serialize(jpConfig.JobPathWorking, Job)
      Thread.Sleep(SleepTime)
      'acAppComObj.Visible = jpConfig.AutoCADVisible
      jp.msg("Extracting PDF.")
      Try
        acDocComObj.SendCommand("ExtractPDF  ")
        jp.msg("PDF Extracted")
      Catch ex As Exception
        jp.msg("Error Extracting PDF: " & ex.Message)
      End Try
      'acAppComObj.Visible = jpConfig.AutoCADVisible
      jp.msg("Closing Document")
      Try
        acAppComObj.ActiveDocument.Close(False)
        'acDocComObj.Close()
        jp.msg("Document Closed.")
      Catch ex As Exception
        jp.msg("Error closing Document: " & ex.Message)
      End Try
      acDocComObj = Nothing
      Job = jobFile.DeSerialize(Job)
    Else
      Job.IsError = True
      Job.ErrorList.Add(New lgErrors("Could Not Load AutoCAD"))
    End If
    Return Job
  End Function

  Public Function ExtractBOMPDF(ByVal Job As jobFile, Optional ByVal BOM As Boolean = True, Optional ByVal PDF As Boolean = True) As jobFile
    If acAppComObj Is Nothing Then
      LoadCad()
    End If
    If (Not acAppComObj Is Nothing) Then
      Dim FilePath As String = jpConfig.TempFolderPath & "\" & Job.FileName
      Dim DLLPath As String = (jpConfig.StartupPath & "\cadConvert.dll").Replace("\", "\\")
      jp.msg("Opening CAD Document:" & FilePath)
      Try
        acDocComObj = acAppComObj.Documents.Open(FilePath, Nothing, Nothing)
        jp.msg("Document Opened.")
      Catch ex As Exception
        jp.msg("Error Opening Document:" & ex.Message)
      End Try
      Thread.Sleep(SleepTime)
      jp.msg("Loading DLL.")
      Try
        'acAppComObj.Visible = jpConfig.AutoCADVisible
        acDocComObj.SendCommand("(command ""SECURELOAD"" ""0"") ")
        acDocComObj.SendCommand("(command ""NETLOAD"" """ & DLLPath & """) ")
        acDocComObj.SendCommand("(command ""SECURELOAD"" ""1"") ")
        'acAppComObj.Visible = jpConfig.AutoCADVisible
        jp.msg("DLL Loaded.")
      Catch ex As Exception
        jp.msg("Error Loading DLL:" & ex.Message)
      End Try
      Job = jobFile.Serialize(jpConfig.JobPathWorking, Job)
      Thread.Sleep(SleepTime)
      If BOM And PDF Then
        jp.msg("Extracting BOM & PDF.")
        Try
          acDocComObj.SendCommand("ExtractBOMPDF  ")
          jp.msg("BOM & PDF Extracted")
        Catch ex As Exception
          jp.msg("Error Extracting BOM & PDF:" & ex.Message)
        End Try
      ElseIf Not BOM And PDF Then
        jp.msg("Extracting PDF.")
        Try
          acDocComObj.SendCommand("ExtractPDF  ")
          jp.msg("PDF Extracted")
        Catch ex As Exception
          jp.msg("Error Extracting PDF: " & ex.Message)
        End Try
      ElseIf BOM And Not PDF Then
        jp.msg("Extracting BOM")
        Try
          acDocComObj.SendCommand("ExtractBOM  ")
          jp.msg("BOM Extracted")
        Catch ex As Exception
          jp.msg("Error Extracting BOM:" & ex.Message)
        End Try
      End If
      'acAppComObj.Visible = jpConfig.AutoCADVisible
      jp.msg("Closing Document")
      Try
        acAppComObj.ActiveDocument.Close(False)
        'acDocComObj.Close()
        jp.msg("Document Closed.")
      Catch ex As Exception
        jp.msg("Error closing Document: " & ex.Message)
      End Try
      acDocComObj = Nothing
      Job = jobFile.DeSerialize(Job)
    Else
      Job.IsError = True
      Job.ErrorList.Add(New lgErrors("Could Not Load AutoCAD"))
    End If
    Return Job
  End Function

  Public Function HandleInComingCall(dwCallType As Integer, hTaskCaller As IntPtr, dwTickCount As Integer, lpInterfaceInfo As IntPtr) As Integer Implements IMessageFilter.HandleInComingCall
    Return 0
  End Function

  Public Function RetryRejectedCall(hTaskCallee As IntPtr, dwTickCount As Integer, dwRejectType As Integer) As Integer Implements IMessageFilter.RetryRejectedCall
    Return &H3E8
  End Function

  Public Function MessagePending(hTaskCallee As IntPtr, dwTickCount As Integer, dwPendingType As Integer) As Integer Implements IMessageFilter.MessagePending
    Return 1
  End Function
  Sub New()
    Dim filter As IMessageFilter = Nothing
    cadUtility.CoRegisterMessageFilter(Me, filter)
  End Sub

  <DllImport("ole32.dll")>
  Private Shared Function CoRegisterMessageFilter(ByVal lpMessageFilter As IMessageFilter, <Out> ByRef lplpMessageFilter As IMessageFilter) As Integer
  End Function

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
        jpConfig = Nothing
        jp = Nothing
      End If

      ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' TODO: set large fields to null.
    End If
    disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    ' TODO: uncomment the following line if Finalize() is overridden above.
    ' GC.SuppressFinalize(Me)
  End Sub
#End Region
End Class
