Imports Autodesk.Connectivity.WebServices
Imports Autodesk.DataManagement.Client.Framework.Currency
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities
Imports Autodesk.DataManagement.Client.Framework.Vault.Settings
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports EDICommon
Public Class vltUtility
  Implements IDisposable

  Public Property jpConfig As ConfigFile = Nothing
  Public Property jp As JobProcess.JobProcessor = Nothing
  Public Function DownloadxData(ByVal Job As jobFile, ByVal xData As DrawingData) As jobFile
    jp.msg("Login in Vault.")
    If VaultActivity.VaultUtil.Login(Job.VaultDB) Then
      jp.msg("Searching " & xData.DrawingFileName)
      Dim tmpFiles() As Autodesk.Connectivity.WebServices.File = VaultActivity.VaultUtil.SearchFileInFolder(xData.DrawingFileName, Job.VaultFolderID, VaultActivity.VaultUtil.SrchOper.Equal)
      'If tmpFiles.Count <= 0 Then
      '  '====Wait to get Stable=====
      '  jp.msg("Waiting 2nd attempt xData")
      '  Threading.Thread.Sleep(1000)
      '  '===========================
      '  tmpFiles = VaultActivity.VaultUtil.SearchFileInFolder(xData.DrawingFileName, Job.VaultFolderID, VaultActivity.VaultUtil.SrchOper.Equal)
      'End If
      If tmpFiles.Count > 0 Then
        Dim tmpfile As Autodesk.Connectivity.WebServices.File = tmpFiles(0)
        If Not tmpfile.CheckedOut Then
          Try
            jp.msg("Downloading " & xData.DrawingFileName)
            VaultActivity.VaultUtil.DownloadFile(tmpfile.Id, jpConfig.TempFolderPath)
          Catch ex As Exception
            Job.IsError = True
            Job.ErrorList.Add(New lgErrors(ex.Message))
          End Try
        Else
          Job.IsError = True
          Job.ErrorList.Add(New lgErrors("File is Checked out."))
        End If
        VaultActivity.VaultUtil.LogOut()
      Else
        Job.IsError = True
        Job.ErrorList.Add(New lgErrors("File Not Found in Vault [Folder]."))
      End If
    Else
      Job.IsError = True
      Job.ErrorList.Add(New lgErrors("Not able to login vault."))
    End If
    Return Job
  End Function
  Public Function DownloadFile(ByVal Job As jobFile) As jobFile
    jp.msg("Login in Vault.")
    If VaultActivity.VaultUtil.Login(Job.VaultDB) Then
      Try
        jp.msg("Downloading " & Job.FileName)
        VaultActivity.VaultUtil.DownloadFile(Convert.ToInt64(Job.FileID), jpConfig.TempFolderPath)
      Catch ex As Exception
        Job.IsError = True
        Job.ErrorList.Add(New lgErrors(ex.Message))
      End Try
      VaultActivity.VaultUtil.LogOut()
    Else
      Job.IsError = True
      Job.ErrorList.Add(New lgErrors("Not able to login vault."))
    End If
    Return Job
  End Function

  Public Function DownloadMain(ByVal Job As jobFile) As jobFile
    jp.msg("Login in Vault.")
    If VaultActivity.VaultUtil.Login(Job.VaultDB) Then
      jp.msg("Searching " & Job.FileName)
      Dim tmpFile As Autodesk.Connectivity.WebServices.File = Nothing
      Try
        tmpFile = VaultActivity.VaultUtil.GetLatestFileByID(Job.FileID)
      Catch ex As Exception
        jp.msg("Error: " & ex.Message)
        jp.msg("Waiting to Find Again -Main")
        Threading.Thread.Sleep(5000)
        Try
          tmpFile = VaultActivity.VaultUtil.GetLatestFileByID(Job.FileID)
        Catch ex1 As Exception
          jp.msg("Could not find again, skipping this file.-Main")
          Job.IsError = True
          Job.ErrorList.Add(New lgErrors("Not able to find file. Err:" & ex1.Message))
        End Try
      End Try
      If tmpFile IsNot Nothing Then
        If Not tmpFile.CheckedOut Then
          Try
            jp.msg("Downloading " & Job.FileName)
            VaultActivity.VaultUtil.DownloadFile(tmpFile.Id, jpConfig.TempFolderPath)
            Job.VaultFolderID = tmpFile.FolderId
          Catch ex As Exception
            Job.IsError = True
            Job.ErrorList.Add(New lgErrors(ex.Message))
          End Try
        Else
          Job.IsError = True
          Job.ErrorList.Add(New lgErrors("File is Checked out."))
        End If
      End If
      VaultActivity.VaultUtil.LogOut()
    Else
      Job.IsError = True
      Job.ErrorList.Add(New lgErrors("Not able to login vault."))
    End If
    Return Job
  End Function

  Public Function DownloadComponentXL(ByVal Job As jobFile) As jobFile
    jp.msg("Login in Vault.")
    If VaultActivity.VaultUtil.Login(Job.VaultDB) Then
      Dim Found As Boolean = False
      jp.msg("Searching Componentt XL: " & IO.Path.GetFileNameWithoutExtension(Job.FileName))
      Dim tmpFile() As Autodesk.Connectivity.WebServices.File = VaultActivity.VaultUtil.SearchFileInFolder(IO.Path.GetFileNameWithoutExtension(Job.FileName), Job.VaultFolderID)
      'If tmpFile.Count <= 0 Then
      '  '====Wait to get Stable=====
      '  jp.msg("Waiting 2nd Attempt ComponentXL")
      '  Threading.Thread.Sleep(1000)
      '  '===========================
      '  tmpFile = VaultActivity.VaultUtil.SearchFileInFolder(IO.Path.GetFileNameWithoutExtension(Job.FileName), Job.VaultFolderID)
      'End If
      For Each tFile As Autodesk.Connectivity.WebServices.File In tmpFile
        Select Case IO.Path.GetExtension(tFile.Name.ToUpper)
          Case "XLS", "XLSX"
            If IO.Path.GetFileNameWithoutExtension(tFile.Name).ToUpper = IO.Path.GetFileNameWithoutExtension(Job.FileName) Then
              If Not tFile.CheckedOut Then
                Try
                  jp.msg("Downloading Component XL " & tFile.Name)
                  VaultActivity.VaultUtil.DownloadFile(tFile.Id, jpConfig.TempFolderPath)
                  jp.msg("Reading XML Data From " & tFile.Name)
                  Dim xData As List(Of DrawingData) = xlConverter.GetDrawingData(jpConfig.TempFolderPath & "\" & tFile.Name)
                  For Each tmpX As DrawingData In xData
                    If tmpX.DrawingFileName.ToUpper = Job.FileName.ToUpper Then
                      Job.ComponentXLData = tmpX
                      Job.IsComponentXLFound = True
                      Job.ComponentXLFileName = tFile.Name
                      Found = True
                      Exit For
                    End If
                  Next
                Catch ex As Exception
                  Job.ErrorList.Add(New lgErrors(ex.Message))
                End Try
                If Found Then
                  Exit For
                End If
              End If
            End If
        End Select
      Next
      If Not Found Then
        Job.IsError = True
      End If
      VaultActivity.VaultUtil.LogOut()
    Else
      Job.IsError = True
      Job.ErrorList.Add(New lgErrors("Not able to login vault."))
    End If
    Return Job
  End Function

  Public Function DownloadMCD(ByVal Job As jobFile) As jobFile
    jp.msg("Login in Vault.")
    If VaultActivity.VaultUtil.Login(Job.VaultDB) Then
      Dim Found As Boolean = False
      jp.msg("Searching All -MCD-")
      Dim tmpFile() As Autodesk.Connectivity.WebServices.File = VaultActivity.VaultUtil.SearchFileInFolder("-MCD-", Job.VaultFolderID)
      'If tmpFile.Count <= 0 Then
      '  '====Wait to get Stable=====
      '  jp.msg("Waiting 2nd attempt MCD")
      '  Threading.Thread.Sleep(1000)
      '  '===========================
      '  tmpFile = VaultActivity.VaultUtil.SearchFileInFolder("-MCD-", Job.VaultFolderID)
      'End If
      jp.msg(tmpFile.Count & " -MCD- files found. Checking All MCD.")
      For Each tFile As Autodesk.Connectivity.WebServices.File In tmpFile
        Select Case IO.Path.GetExtension(tFile.Name.ToUpper)
          Case "XLS", "XLSX"
            If Not tFile.CheckedOut Then
              Try
                jp.msg("Downloading " & tFile.Name)
                VaultActivity.VaultUtil.DownloadFile(tFile.Id, jpConfig.TempFolderPath)
                Try
                  jp.msg("Reading XML Data, and finding Job File.")
                  Dim xData As List(Of DrawingData) = xlConverter.GetDrawingData(jpConfig.TempFolderPath & "\" & tFile.Name)
                  For Each tmpX As DrawingData In xData
                    If tmpX.DrawingFileName.ToUpper = Job.FileName.ToUpper Then
                      Job.MCDxData = tmpX
                      Job.IsMCDFound = True
                      Job.MCDFileName = tFile.Name
                      Found = True
                      jp.msg("Job Details found in MCD: " & tFile.Name)
                      Exit For
                    End If
                  Next
                Catch ex As Exception
                End Try
                'Delete MCD
              Catch ex As Exception
                Job.ErrorList.Add(New lgErrors(ex.Message))
              End Try
              If Found Then
                Exit For
              End If
            End If
        End Select
      Next
      If Not Found Then
        Job.IsError = True
      End If
      VaultActivity.VaultUtil.LogOut()
    Else
      Job.IsError = True
      Job.ErrorList.Add(New lgErrors("Not able to login vault."))
    End If
    Return Job
  End Function


  Sub New(ByVal vltServer As String, ByVal vltUser As String, ByVal vltPass As String)
    VaultActivity.VaultUtil.serverName = vltServer
    VaultActivity.VaultUtil.userName = vltUser
    VaultActivity.VaultUtil.password = vltPass
  End Sub
  Sub New()
    'dummy
  End Sub

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
        jp = Nothing
        jpConfig = Nothing
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
