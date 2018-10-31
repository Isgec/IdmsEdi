Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Text
Namespace SIS.EDI
  Public Class ediAlerts
    Public Shared Property Testing As Boolean = False
    Private Enum tmtlType
      Customer = 1
      Internal = 2
      Site = 3
      Vendor = 4
    End Enum
    Private Class emp
      Public Property empID As Integer = 0
      Public Property empName As String = ""
      Public Property empEMail As String = ""
      Public Property webUser As String = ""

      Public Shared Function GetEmp(ByVal empID As Integer) As emp
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " emp1.t_nama as empName,"
        mSql = mSql & " bpe1.t_mail as empEMail "
        mSql = mSql & " from ttccom001200 as emp1 "
        mSql = mSql & " left outer join tbpmdm001200 as bpe1 on emp1.t_emno=bpe1.t_emno "
        mSql = mSql & " where emp1.t_emno = '" & empID & "'"
        Dim tmp As emp = Nothing
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            tmp = New emp
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            If (Reader.Read()) Then
              With tmp
                .empID = empID
                If Not Convert.IsDBNull(Reader("empName")) Then .empName = Reader("empName")
                If Not Convert.IsDBNull(Reader("empEMail")) Then .empEMail = Reader("empEMail")
              End With
            End If
            Reader.Close()
          End Using
        End Using
        Return tmp
      End Function

      Public Shared Function GetReceiptCreator(ByVal tmtlID As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " t_user as [user] "
        mSql = mSql & " from tdmisg134200 "
        mSql = mSql & " where t_trno = '" & tmtlID & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            If (Reader.Read()) Then
              If Not Convert.IsDBNull(Reader("user")) Then tmp = Reader("user")
            End If
            Reader.Close()
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetPONoOfReceipt(ByVal tmtlID As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " t_orno  "
        mSql = mSql & " from tdmisg134200 "
        mSql = mSql & " where t_trno = '" & tmtlID & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
            If tmp Is Nothing Then tmp = ""
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetSiteEmailIDs(ByVal ProjectID As String, ByVal AddressID As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " t_mail  "
        mSql = mSql & " from tdmisg126200 "
        mSql = mSql & " where t_cprj = '" & ProjectID & "'"
        mSql = mSql & " and   t_cadr = '" & AddressID & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
            If tmp Is Nothing Then tmp = ""
          End Using
        End Using
        Return tmp
      End Function

      Public Shared Function GetTCPOIssuer(ByVal PONo As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " IssuedBy  "
        mSql = mSql & " from pak_po "
        mSql = mSql & " where pofor='TC' and ponumber = '" & PONo & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
            If tmp Is Nothing Then tmp = ""
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetPOBuyer(ByVal PONo As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " BuyerID  "
        mSql = mSql & " from pak_po "
        mSql = mSql & " where pofor='TC' and ponumber = '" & PONo & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
            If tmp Is Nothing Then tmp = ""
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetWebUser(ByVal webUser As String) As emp
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " UserFullName as empName,"
        mSql = mSql & " emailid as empEMail "
        mSql = mSql & " from aspnet_users "
        mSql = mSql & " where username = '" & webUser & "'"
        Dim tmp As emp = Nothing
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            tmp = New emp
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            If (Reader.Read()) Then
              With tmp
                .webUser = webUser
                If Not Convert.IsDBNull(Reader("empName")) Then .empName = Reader("empName")
                If Not Convert.IsDBNull(Reader("empEMail")) Then .empEMail = Reader("empEMail")
              End With
            End If
            Reader.Close()
          End Using
        End Using
        Return tmp
      End Function
    End Class
    Private Shared Function gma(ByVal tmp As emp, ByRef aErr As ArrayList, Optional ByVal State As String = "") As MailAddress
      Dim x As MailAddress = Nothing
      If tmp IsNot Nothing Then
        If tmp.empEMail <> "" Then
          Try
            x = New MailAddress(tmp.empEMail, tmp.empName)
          Catch ex As Exception
            aErr.Add(State & "=> " & ex.Message)
          End Try
        Else
          aErr.Add(State & "=> " & tmp.empID & " : " & tmp.empName)
        End If
      End If
      Return x
    End Function
    Public Shared Function TmtlAlert(ByVal TransmittalID As String) As Boolean
      Dim oTmtl As EDICommon.SIS.EDI.ediTmtlH = EDICommon.SIS.EDI.ediTmtlH.ediTmtlHGetByID(TransmittalID)

      Dim aErr As New ArrayList
      Dim mRet As String = ""
      Dim oClient As SmtpClient = New SmtpClient("192.9.200.214", 25)
      Dim oMsg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage()
      oClient.Credentials = New Net.NetworkCredential("adskvaultadmin", "isgec@123")

      Dim Issuer As emp = emp.GetEmp(oTmtl.t_isby)
      Dim Approver As emp = emp.GetEmp(oTmtl.t_apsu)
      Dim Creator As emp = emp.GetEmp(oTmtl.t_user)
      With oMsg
        .Subject = "Download Documents of Transmittal: " & TransmittalID
        Dim EMailIDError As String = ""
        Dim x As MailAddress = Nothing
        Try
          Select Case oTmtl.t_type
            Case tmtlType.Customer
              x = gma(Issuer, aErr, "CT-Issuer-From")
              If x IsNot Nothing Then
                .From = x
                .CC.Add(x)
              End If
              x = gma(Approver, aErr, "CT-Approver-To")
              If x IsNot Nothing Then .To.Add(x)
              x = gma(Creator, aErr, "CT-Creator-CC")
              If x IsNot Nothing Then .CC.Add(x)
            Case tmtlType.Internal
              Dim IssuedTo As emp = emp.GetEmp(oTmtl.t_logn)
              x = gma(Issuer, aErr, "IT-Issuer-From")
              If x IsNot Nothing Then
                .From = x
                .CC.Add(x)
              End If
              x = gma(IssuedTo, aErr, "IT-IssuedTo-To")
              If x IsNot Nothing Then .To.Add(x)
              x = gma(Approver, aErr, "IT-Approver-CC")
              If x IsNot Nothing Then .CC.Add(x)
              x = gma(Creator, aErr, "IT-Creator-CC")
              If x IsNot Nothing Then .CC.Add(x)
            Case tmtlType.Site
              x = gma(Issuer, aErr, "ST-Issuer-From")
              If x IsNot Nothing Then
                .From = x
                .CC.Add(x)
              End If
              x = gma(Approver, aErr, "ST-Approver-To")
              If x IsNot Nothing Then .To.Add(x)
              x = gma(Creator, aErr, "ST-Creator-CC")
              If x IsNot Nothing Then .CC.Add(x)
              'Transmittal Site Address
              Dim SiteIDs As String = emp.GetSiteEmailIDs(oTmtl.t_cprj, oTmtl.t_padr)
              If SiteIDs <> "" Then
                Dim aIDs() As String = SiteIDs.Split(",;".ToCharArray)
                For Each id As String In aIDs
                  Try
                    x = New MailAddress(id.Trim, id.Trim)
                    .CC.Add(x)
                  Catch ex As Exception
                  End Try
                Next
              End If
            Case tmtlType.Vendor
              x = gma(Issuer, aErr, "VT-Issuer-From")
              If x IsNot Nothing Then
                .From = x
                .CC.Add(x)
              End If
              Dim tmp As String = emp.GetReceiptCreator(TransmittalID)
              If tmp = "SUPPLIER" Or tmp = "" Then
                x = gma(Approver, aErr, "VT-Approver-To")
                If x IsNot Nothing Then .To.Add(x)
              Else
                Dim tmpR As emp = emp.GetEmp(tmp)
                x = gma(tmpR, aErr, "VT-ReceiptCreator-To")
                If x IsNot Nothing Then .To.Add(x)
                x = gma(Approver, aErr, "VT-Approver-CC")
                If x IsNot Nothing Then .CC.Add(x)
              End If
              x = gma(Creator, aErr, "VT-Creator-CC")
              If x IsNot Nothing Then .CC.Add(x)
              'GetPOIssuer From Joomla
              Dim tmpPONo As String = emp.GetPONoOfReceipt(TransmittalID)
              If tmpPONo <> "" Then
                Dim POIssuer As String = emp.GetTCPOIssuer(tmpPONo)
                Dim POBuyer As String = emp.GetPOBuyer(tmpPONo)
                x = gma(emp.GetWebUser(POIssuer), aErr, "PO-Issuer")
                If x IsNot Nothing Then .CC.Add(x)
                x = gma(emp.GetWebUser(POBuyer), aErr, "PO-Buyer")
                If x IsNot Nothing Then .CC.Add(x)
              End If
          End Select
        Catch ex As Exception
          EMailIDError = ex.Message
        End Try
        If .To.Count <= 0 Then
          x = New MailAddress("baansupport@isgec.co.in", "BaaN Support")
          If Not .To.Contains(x) Then .To.Add(x)
          x = New MailAddress("lalit@isgec.co.in", "Lalit Gupta")
          If Not .CC.Contains(x) Then .CC.Add(x)
          x = New MailAddress("harishkumar@isgec.co.in", "Harish Kaushik")
          If Not .CC.Contains(x) Then .CC.Add(x)
        End If
        If EMailIDError <> "" Then
          x = New MailAddress("lalit@isgec.co.in", "Lalit Gupta")
          If Not .To.Contains(x) Then .To.Add(x)
          x = New MailAddress("harishkumar@isgec.co.in", "Harish Kaushik")
          If Not .To.Contains(x) Then .To.Add(x)
        End If

        If .From Is Nothing Then
          x = New MailAddress("baansupport@isgec.co.in", "BaaN Support")
          .From = x
          If Not Testing Then
            If Not .CC.Contains(x) Then .CC.Add(x)
          End If
        End If
        '====================
        Dim TestIDs As New ArrayList
        If Testing Then
          For Each xx As MailAddress In .To
            TestIDs.Add("TO => " & xx.User & " : " & xx.Address)
          Next
          .To.Clear()
          x = New MailAddress("lalit@isgec.co.in", "Lalit Gupta")
          .To.Add(x)
          For Each xx As MailAddress In .CC
            TestIDs.Add("CC => " & xx.User & " : " & xx.Address)
          Next
          .CC.Clear()
        End If
        '====================
        .IsBodyHtml = True
        Dim tblStr As String = EDICommon.SIS.EDI.ediTmtlH.GetHTML(TransmittalID)
        Dim Header As String = ""
        Header &= "<html xmlns=""http://www.w3.org/1999/xhtml"">"
        Header &= "<head>"
        Header &= "<title></title>"
        Header &= "<style>"
        Header &= "table{"

        Header &= "border: solid 1pt black;"
        Header &= "border-collapse:collapse;"
        Header &= "font-family: Tahoma;}"

        Header &= "td{"
        Header &= "border: solid 1pt black;"
        Header &= "font-family: Tahoma;"
        Header &= "font-size: 12px;"
        Header &= "padding: 2px 2px 4px 4px;"
        Header &= "vertical-align:middle;}"

        Header &= "a{"
        Header &= "color: white;}"
        Header &= "a:hover{"
        Header &= "color: hotpink;}"

        Header &= "</style>"
        Header &= "</head>"
        Header &= "<body>"
        If aErr.Count > 0 Then
          Header &= "<br/>"
          Header &= "<br/>"
          Header &= "<table>"
          Header &= "<tr><td style=""color: red""><i><b>"
          Header &= "NOTE: Download Link could not be delivered to following recipient(s), Please update their E-Mail ID in ERP-LN."
          Header &= "</b></i></td></tr>"
          For Each Err As String In aErr
            Header &= "<tr><td color=""red""><i>"
            Header &= Err
            Header &= "</i></td></tr>"
          Next
          Header &= "</table>"
        End If
        If EMailIDError <> "" Then
          Header &= "<br/>"
          Header &= "<br/>"
          Header &= "<h3>" & EMailIDError & "</h3>"
        End If
        If Testing Then
          If TestIDs.Count > 0 Then
            Header &= "<br/>"
            Header &= "<br/>"
            Header &= "<table>"
            Header &= "<tr><td style=""color: red""><i><b>"
            Header &= "TESTING"
            Header &= "</b></i></td></tr>"
            For Each test As String In TestIDs
              Header &= "<tr><td color=""red""><i>"
              Header &= test
              Header &= "</i></td></tr>"
            Next
            Header &= "</table>"
          End If
        End If
        Header &= "<table style='margin-left:10px;width:1000px;'>"
        Header &= "<tr><td style='background-color:DodgerBlue;text-align:center;color:white;font-size:16px;height:30px;verticle-align:middle;'><b>"
        Header &= "<a href='http://192.9.200.146/WebEitl1/ediWTmtlH.aspx?t_tran=" & TransmittalID & "'>Click to Download Transmittal Documents</a>"
        Header &= "</b></td></tr>"
        Header &= "</table>"
        Header &= "<br/>"
        Header &= "<br/>"
        Header &= tblStr
        Header &= "</body></html>"
        .Body = Header
      End With
      Try
        oClient.Send(oMsg)
      Catch ex As Exception
      End Try
      Return True
    End Function
  End Class
End Namespace
