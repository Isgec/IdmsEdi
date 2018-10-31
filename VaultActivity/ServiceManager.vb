Imports Autodesk.Connectivity.WebServices
Imports Autodesk.Connectivity.WebServicesTools
Imports Autodesk.DataManagement.Client.Framework.Vault
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections
Imports Autodesk.DataManagement.Client.Framework.Vault.Results
Imports System
Imports System.Web
Imports System.Web.Services.Protocols

Public Class ServiceManager
  Private mAuthService As AuthService
  Private mConnection As Connection
  Private mDocumentService As DocumentService
  Private mFileStoreService As FilestoreService
  Private mPropertyService As PropertyService
  Private mSecurityService As SecurityService
  Private mWSManager As WebServiceManager
  Private mLifeCycleService As LifeCycleService
  Private Shared oServiceManager As ServiceManager
  Private mDocumentServiceExtensions As DocumentServiceExtensions = DirectCast(Nothing, DocumentServiceExtensions)
  Public Shared Function GetServiceManager() As ServiceManager
    Try
      If oServiceManager Is Nothing Then
        oServiceManager = New ServiceManager()
      End If
    Catch exception As Exception
      Throw exception
    End Try
    Return oServiceManager
  End Function
  Public Shared Function GetDocumentService() As DocumentService
    Try
      GetServiceManager()
      If oServiceManager.mDocumentService Is Nothing Then
        oServiceManager.mDocumentService = oServiceManager.mWSManager.DocumentService
      End If
    Catch exception As Exception
      Throw exception
    End Try
    Return oServiceManager.mDocumentService
  End Function
  Public Shared Function GetFileStoreService() As FilestoreService
    Try
      GetServiceManager()
      If oServiceManager.mFileStoreService Is Nothing Then
        oServiceManager.mFileStoreService = oServiceManager.mWSManager.FilestoreService
      End If
    Catch exception As Exception
      Throw exception
    End Try
    Return oServiceManager.mFileStoreService
  End Function
  Public Shared Function GetPropertyService() As PropertyService
    Try
      GetServiceManager()
      If oServiceManager.mPropertyService Is Nothing Then
        oServiceManager.mPropertyService = oServiceManager.mWSManager.PropertyService
      End If
    Catch exception As Exception
      Throw exception
    End Try
    Return oServiceManager.mPropertyService
  End Function
  Public Shared Function GetDocumentServiceExtensions() As DocumentServiceExtensions
    Try
      GetServiceManager()
      If oServiceManager.mDocumentServiceExtensions Is Nothing Then
        oServiceManager.mDocumentServiceExtensions = oServiceManager.mWSManager.DocumentServiceExtensions
      End If
    Catch ex As Exception
      Throw ex
    End Try
    Return oServiceManager.mDocumentServiceExtensions
  End Function
  Public Shared Function GetLifeCycleService() As LifeCycleService
    Try
      GetServiceManager()
      If oServiceManager.mLifeCycleService Is Nothing Then
        oServiceManager.mLifeCycleService = oServiceManager.mWSManager.LifeCycleService
      End If
    Catch ex As Exception
      Throw ex
    End Try
    Return oServiceManager.mLifeCycleService
  End Function
  Public Shared Function GetVaultConnection() As Connection
    Return oServiceManager.mConnection
  End Function
  Public Shared Function LogInStandard(ByVal serverName As String, ByVal vaultName As String, ByVal userName As String, ByVal password As String) As SecurityService
    Try
      GetServiceManager()
      If oServiceManager.mAuthService Is Nothing Then
        Dim result As LogInResult = Library.ConnectionManager.LogIn(serverName, vaultName, userName, password, AuthenticationFlags.Standard, Nothing)
        If result.Connection Is Nothing Then
          Throw New Exception("Invalid Vault login credentials")
        End If
        oServiceManager.mConnection = result.Connection
        oServiceManager.mWSManager = oServiceManager.mConnection.WebServiceManager
        oServiceManager.mAuthService = oServiceManager.mConnection.WebServiceManager.AuthService
        oServiceManager.mSecurityService = oServiceManager.mWSManager.SecurityService
      End If
    Catch exception As Exception
      Throw exception
    End Try
    Return oServiceManager.mSecurityService
  End Function
  Public Shared Function LogInReadOnly(ByVal serverName As String, ByVal vaultName As String, ByVal userName As String, ByVal password As String) As SecurityService
    Try
      GetServiceManager()
      If oServiceManager.mAuthService Is Nothing Then
        Dim result As LogInResult = Library.ConnectionManager.LogIn(serverName, vaultName, userName, password, AuthenticationFlags.ReadOnly, Nothing)
        If result.Connection Is Nothing Then
          Throw New Exception("Invalid Vault login credentials")
        End If
        oServiceManager.mConnection = result.Connection
        oServiceManager.mWSManager = oServiceManager.mConnection.WebServiceManager
        oServiceManager.mAuthService = oServiceManager.mConnection.WebServiceManager.AuthService
        oServiceManager.mSecurityService = oServiceManager.mWSManager.SecurityService
      End If
    Catch exception As Exception
      Throw exception
    End Try
    Return oServiceManager.mSecurityService
  End Function
  Public Shared Sub LogOut()
    GetServiceManager()
    If Not (oServiceManager.mAuthService Is Nothing) Then
      oServiceManager.mAuthService.SignOut()
    End If
    oServiceManager = Nothing
  End Sub
  Private Function SetServiceUrl(ByVal service As SoapHttpClientProtocol) As String
    Dim builder As New UriBuilder(service.Url)
    builder.Host = VaultUtil.serverName
    service.Url = builder.Uri.ToString()
    Return service.Url
  End Function
End Class

