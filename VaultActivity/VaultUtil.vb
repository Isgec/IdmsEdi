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

Public Class VaultUtil
  Public Shared password As String = String.Empty
  Public Shared serverName As String = String.Empty
  Public Shared userName As String = String.Empty
  Public Shared vaultName As String = String.Empty
  Public Enum SrchOper
    Contains = 1
    Equal = 3
    NotEmpty = 5
  End Enum

  Private Shared Function BuildSearchCondition(ByVal entityClassId As String, ByVal propertyName As String, ByVal searchText As String, Optional ByVal searchOperation As SrchOper = SrchOper.Contains) As SrchCond
    Dim propertyService As PropertyService = ServiceManager.GetPropertyService()
    Dim source As PropDef() = Nothing
    Try
      source = propertyService.GetPropertyDefinitionsByEntityClassId(entityClassId)
    Catch exception As Exception
      Dim message As String = exception.Message
    End Try
    Dim def As PropDef = source.SingleOrDefault(Function(n) n.DispName = propertyName)

    Return New SrchCond() With {.PropDefId = def.Id, .PropTyp = PropertySearchType.SingleProperty, .SrchOper = searchOperation, .SrchRule = SearchRuleType.Must, .SrchTxt = searchText}
  End Function

  Public Shared Function DownloadFile(ByVal fileID As Long, ByVal localFolderPath As String) As Boolean
    Try
      Dim file As New FileIteration(ServiceManager.GetVaultConnection(), GetFileByID(fileID))
      Dim settings As New AcquireFilesSettings(ServiceManager.GetVaultConnection(), True) With {.LocalPath = New FolderPathAbsolute(localFolderPath)}
      settings.AddFileToAcquire(file, AcquireFilesSettings.AcquisitionOption.Download)
      ServiceManager.GetVaultConnection().FileManager.AcquireFiles(settings)
    Catch exception As Exception
      Throw exception
    End Try
    Return True
  End Function

  Public Shared Function DownloadFile(ByVal fileName As String, ByVal localFolderPath As String) As Boolean
    Dim flag As Boolean
    Try
      Dim file As Autodesk.Connectivity.WebServices.File = SearchFileByName(fileName, 1L)
      If Not (file Is Nothing) Then
        Return DownloadFile(file.Id, localFolderPath)
      End If
      flag = False
    Catch exception As Exception
      Throw exception
    End Try
    Return flag
  End Function

  Public Shared Function GetFileByID(ByVal fileID As Long) As Autodesk.Connectivity.WebServices.File
    Dim fileById As Autodesk.Connectivity.WebServices.File = Nothing
    Try
      fileById = ServiceManager.GetDocumentService().GetFileById(fileID)
    Catch exception As Exception
      Throw exception
    End Try
    Return fileById
  End Function

  Public Shared Function GetLatestFileByID(ByVal fileID As Long) As Autodesk.Connectivity.WebServices.File
    Dim fileById As Autodesk.Connectivity.WebServices.File = Nothing
    Try
      Dim documentService As DocumentService = ServiceManager.GetDocumentService()
      fileById = documentService.GetFileById(fileID)
      fileById = documentService.GetLatestFileByMasterId(fileById.MasterId)
    Catch exception As Exception
      Throw exception
    End Try
    Return fileById
  End Function
  Public Shared Function LogInStandard(ByVal serverName As String, ByVal vaultName As String, ByVal userName As String, ByVal password As String) As Boolean
    Try
      If ServiceManager.LogInStandard(serverName, vaultName, userName, password) IsNot Nothing Then
        Return True
      End If
    Catch exception As Exception
      Throw exception
    End Try
    Return False
  End Function
  Public Shared Function LogInReadOnly(ByVal serverName As String, ByVal vaultName As String, ByVal userName As String, ByVal password As String) As Boolean
    Try
      If ServiceManager.LogInReadOnly(serverName, vaultName, userName, password) IsNot Nothing Then
        Return True
      End If
    Catch exception As Exception
      Throw exception
    End Try
    Return False
  End Function
  Public Shared Function CheckOutFile(ByVal fileID As Long) As Boolean
    Try
      Dim file As New FileIteration(ServiceManager.GetVaultConnection(), VaultUtil.GetLatestFileByID(fileID))
      Dim settings As New AcquireFilesSettings(ServiceManager.GetVaultConnection(), True)
      settings.AddFileToAcquire(file, AcquireFilesSettings.AcquisitionOption.Checkout)
      ServiceManager.GetVaultConnection().FileManager.AcquireFiles(settings)
    Catch ex As Exception
      Throw ex
    End Try
    Return True
  End Function
  Public Shared Function CheckInFile(ByVal oFile As Autodesk.Connectivity.WebServices.File) As Autodesk.Connectivity.WebServices.File
    Dim file As Autodesk.Connectivity.WebServices.File
    Try
      file = ServiceManager.GetDocumentService().CheckinUploadedFile(oFile.MasterId, "Properties are updated by XmlProcessor", False, DateTime.Now, DirectCast(Nothing, FileAssocParam()), DirectCast(Nothing, BOM), _
       True, oFile.Name, oFile.FileClass, False, DirectCast(Nothing, ByteArray))
    Catch ex As Exception
      Throw ex
    End Try
    Return file
  End Function
  Public Shared Function UpdateFileLifecycle(ByVal fileMasterID As Long, ByVal toStateId As Long, ByVal comment As String) As Boolean
    Dim flag As Boolean
    Try
      ServiceManager.GetDocumentServiceExtensions().UpdateFileLifeCycleStates(New Long(0) {fileMasterID}, New Long(0) {toStateId}, comment)
      flag = True
    Catch ex As InvalidOperationException
      Throw New Exception("File not found in Vault")
    Catch ex As Exception
      Throw ex
    End Try
    Return flag
  End Function
  Public Shared Function GetPropDetailsByDispName(ByVal propDispName As String, ByVal entityClassId As String) As PropDef
    Dim propertyService As PropertyService = ServiceManager.GetPropertyService()
    Dim propDefArray As PropDef() = Nothing
    Try
      propDefArray = propertyService.GetPropertyDefinitionsByEntityClassId(entityClassId)
    Catch ex As Exception
      Dim message As String = ex.Message
    End Try
    Return Enumerable.[Single](Of PropDef)(DirectCast(propDefArray, IEnumerable(Of PropDef)), DirectCast(Function(n) n.DispName = propDispName, Func(Of PropDef, Boolean)))
  End Function

  Public Shared Function GetLifeCycleStateId(ByVal fileDefId As Long, ByVal fileStateName As String) As Long
    Dim lifeCycleService As LifeCycleService = ServiceManager.GetLifeCycleService()
    Dim num As Long = 0L
    Try
      Dim cycleDefinitions As LfCycDef() = lifeCycleService.GetAllLifeCycleDefinitions()
      Dim index As Integer = 0
      While index < Enumerable.Count(Of LfCycDef)(DirectCast(cycleDefinitions, IEnumerable(Of LfCycDef)))
        For Each lfCycState As LfCycState In cycleDefinitions(index).StateArray
          If lfCycState.LfCycDefId = fileDefId Then
            If (lfCycState.Name.ToLower = fileStateName.ToLower Or lfCycState.DispName.ToLower = fileStateName.ToLower) Then
              num = lfCycState.Id
              Exit For
            End If
          End If
        Next
        If num > 0L Then
          Exit While
        End If
        System.Threading.Interlocked.Increment(index)
      End While
    Catch ex As Exception
      Dim message As String = ex.Message
    End Try
    Return num
  End Function
  Public Shared Sub LogOut()
    ServiceManager.LogOut()
  End Sub
  Public Shared Function SearchFileByName(ByVal fileName As String, Optional ByVal srchOper As Long = 1L) As Autodesk.Connectivity.WebServices.File
    Dim file As Autodesk.Connectivity.WebServices.File = Nothing
    Try
      Dim documentService As DocumentService = ServiceManager.GetDocumentService()
      Dim cond As SrchCond = BuildSearchCondition("FILE", "File Name", fileName, srchOper)

      Dim bookmark As String = String.Empty
      Dim searchstatus As SrchStatus = Nothing
      Dim list As New List(Of Autodesk.Connectivity.WebServices.File)()
      Dim collection As Autodesk.Connectivity.WebServices.File() = Nothing
      While (searchstatus Is Nothing) OrElse (list.Count < searchstatus.TotalHits)
        collection = documentService.FindFilesBySearchConditions(New SrchCond() {cond}, Nothing, Nothing, False, True, bookmark, searchstatus)
        If collection Is Nothing Then
          Exit While
        End If
        list.AddRange(collection)
      End While
      If collection Is Nothing Then
        Throw New FileNotFoundException()
      End If
      file = collection(0)
    Catch exception As Exception
      Throw exception
    End Try
    Return file
  End Function
  Public Shared Function SearchFileInFolder(ByVal fileNameSearchText As String, ByVal folderID As Long, Optional ByVal srchOper As SrchOper = SrchOper.Contains) As Autodesk.Connectivity.WebServices.File()
    Dim fileArray() As Autodesk.Connectivity.WebServices.File = Nothing
    Try
      Dim documentService As DocumentService = ServiceManager.GetDocumentService()
      Dim cond As SrchCond = BuildSearchCondition("FILE", "File Name", fileNameSearchText, srchOper)
      Dim bookmark As String = String.Empty
      Dim searchstatus As SrchStatus = Nothing
      Dim list As New List(Of Autodesk.Connectivity.WebServices.File)
      Dim collection() As Autodesk.Connectivity.WebServices.File = Nothing
      While (searchstatus Is Nothing) OrElse (list.Count < searchstatus.TotalHits)
        collection = documentService.FindFilesBySearchConditions(New SrchCond() {cond}, Nothing, New Long() {folderID}, False, True, bookmark, searchstatus)
        If collection Is Nothing Then
          Exit While
        End If
        list.AddRange(collection)
      End While
      fileArray = list.ToArray
    Catch exception As Exception
      Throw exception
    End Try
    Return fileArray
  End Function
  Public Shared Function UndoCheckOut(ByVal fileMasterID As Long) As Autodesk.Connectivity.WebServices.File
    Dim file As Autodesk.Connectivity.WebServices.File = Nothing
    Try
      Dim array As New ByteArray
      file = ServiceManager.GetDocumentService().UndoCheckoutFile(fileMasterID, array)
    Catch ex As Exception
      Throw ex
    End Try
    Return file
  End Function

  Public Shared Function UpdateFileProperty(ByVal fileMasterID As Long, ByVal propDispName As String, ByVal propValue As Object) As Boolean
    Dim flag As Boolean = False
    Try
      Dim documentService As DocumentService = ServiceManager.GetDocumentService()
      Dim propDetailsByDispName As PropDef = GetPropDetailsByDispName(propDispName, "FILE")
      documentService.UpdateFileProperties(New Long() {fileMasterID}, New Long() {propDetailsByDispName.Id}, New Object() {propValue})
      flag = True
    Catch generatedExceptionName As InvalidOperationException
      Throw New Exception("Property " + propDispName + " not found in Vault")
    Catch ex As Exception
      Throw ex
    End Try
    Return flag
  End Function

  Public Shared Function CheckOutAndDownloadFile(ByVal fileID As Long, ByVal localFolderPath As String) As Boolean
    Try
      Dim file As New FileIteration(ServiceManager.GetVaultConnection(), GetLatestFileByID(fileID))
      Dim settings As New AcquireFilesSettings(ServiceManager.GetVaultConnection(), True) With {.LocalPath = New FolderPathAbsolute(localFolderPath)}
      settings.AddFileToAcquire(file, AcquireFilesSettings.AcquisitionOption.Checkout Or AcquireFilesSettings.AcquisitionOption.Download)
      ServiceManager.GetVaultConnection().FileManager.AcquireFiles(settings)
    Catch ex As Exception
      Throw ex
    End Try
    Return True
  End Function
  Public Shared Function Login(ByVal VaultDB As String) As Boolean
    VaultUtil.userName = "VAULTPROCESS"
    VaultUtil.password = "VAULTPROCESS"
    VaultUtil.serverName = "192.9.200.119"
    VaultUtil.vaultName = VaultDB
    Try
      VaultUtil.LogInStandard(VaultUtil.serverName, VaultUtil.vaultName, VaultUtil.userName, VaultUtil.password)
    Catch ex As Exception
      Return False
    End Try
    Return True
  End Function

End Class

'The PropTyp property basically determines if this is an advanced search or a basic search. 
'“SingleProperty” means advanced search because the value only relates to a single property. 
'“AllProperties” And “AllPropertiesAndContent” both refer to the basic search feature where a 
'Single value Is compared against all properties.

'SrchOper defines the comparison operator, such as “equals” or “less than”.  
'It's a number, and the SDK documentation has a table of all the legal values.

