Imports System.Security
Imports System.Security.Principal
Imports System.IO


Public Class Impersonator
  Dim LOGON32_LOGON_INTERACTIVE As Integer = 2
  Dim LOGON32_PROVIDER_DEFAULT As Integer = 0
  '
  Private Declare Function LogonUserA Lib "advapi32.dll" (ByVal lpxzUsername As String,
                                            ByVal lpszDomain As String,
                                            ByVal lpszpassword As String,
                                            ByVal dwLogonType As Integer,
                                            ByVal dwLogonProvider As Integer,
                                            ByRef phToken As IntPtr) As Integer
  Private Declare Auto Function DuplicateToken Lib "advapi32.dll" (
                                            ByVal ExistingTokenHandle As IntPtr,
                                            ByVal ImpersonationLevel As Integer,
                                            ByRef DuplicateTokenHandle As IntPtr) As Integer
  Private Declare Auto Function RevertToSelf Lib "advapi32.dll" () As Long
  '
  Private Declare Auto Function CloseHandle Lib "Kernel32.dll" (ByVal handle As IntPtr) As Long
  '
  Dim impersonationContext As WindowsImpersonationContext

  Public Function ImpersonateValidUser(ByVal userName As String, ByVal domain As String, ByVal password As String) As Boolean
    Dim tempWindowsIdentity As WindowsIdentity
    Dim token As IntPtr = IntPtr.Zero
    Dim tokenDuplicate As IntPtr = IntPtr.Zero
    ImpersonateValidUser = False
    '
    Try
      If RevertToSelf() Then
        If LogonUserA(userName, domain, password, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, token) <> 0 Then
          If DuplicateToken(token, 2, tokenDuplicate) <> 0 Then
            tempWindowsIdentity = New WindowsIdentity(tokenDuplicate)
            impersonationContext = tempWindowsIdentity.Impersonate()
            If Not impersonationContext Is Nothing Then
              ImpersonateValidUser = True
            End If
          End If
        End If
      End If
      '
      If Not tokenDuplicate.Equals(IntPtr.Zero) Then
        CloseHandle(tokenDuplicate)
      End If
      '
      If Not token.Equals(IntPtr.Zero) Then
        CloseHandle(token)
      End If
      ''
    Catch ex As Exception
      Throw New ApplicationException("User not found in the server or incorrect password. User Name: " + userName + " Domain: " + domain, ex)
    End Try
  End Function
  '
  Public Sub UndoImpersonation()
    impersonationContext.Undo()
  End Sub
  Public Shared Function Impersonate(ByVal LoginID As String, ByVal Domain As String, ByVal Password As String) As Impersonator
    Dim tmp As New Impersonator
    tmp.ImpersonateValidUser(LoginID, Domain, Password)
    Return tmp
  End Function
End Class