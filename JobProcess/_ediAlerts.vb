﻿Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Text
Imports EDICommon
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

      Public Shared Function GetEmp(ByVal empID As Integer, comp As String) As emp
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " emp1.t_nama as empName,"
        mSql = mSql & " bpe1.t_mail as empEMail "
        mSql = mSql & " from ttccom001" & comp & " as emp1 "
        mSql = mSql & " left outer join tbpmdm001" & comp & " as bpe1 on emp1.t_emno=bpe1.t_emno "
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

      Public Shared Function GetReceiptCreator(ByVal tmtlID As String, comp As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " t_user as [user] "
        mSql = mSql & " from tdmisg134" & comp & " "
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
      Public Shared Function GetPONoOfReceipt(ByVal tmtlID As String, comp As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " t_orno  "
        mSql = mSql & " from tdmisg134" & comp & " "
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
      Public Shared Function GetSiteEmailIDs(ByVal ProjectID As String, ByVal AddressID As String, Comp As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " t_mail  "
        mSql = mSql & " from tdmisg126" & Comp & " "
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
      Public Shared Function GetVendorEmailIDs(ByVal VendorID As String, ByVal AddressID As String, comp As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " t_mail  "
        mSql = mSql & " from tdmisg128" & comp & " "
        mSql = mSql & " where t_ofbp = '" & VendorID & "'"
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
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetToolsConnectionString())
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
      Public Shared Function GetPOSupplierEmailByReceiptNo(ReceiptNo As String, RevisionNo As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " isnull(aa.emailid,'') as emailid "
        mSql = mSql & " from vr_businessPartner as aa "
        mSql = mSql & " inner join pak_po as bb on aa.bpid=bb.supplierid "
        mSql = mSql & " inner join pak_polinerec as cc on bb.serialno=cc.serialno "
        mSql = mSql & " where cc.receiptno='" & ReceiptNo & "' and cc.revisionno='" & RevisionNo & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            If (Reader.Read()) Then
              If Not Convert.IsDBNull(Reader("emailid")) Then tmp = Reader("emailid")
            End If
            Reader.Close()
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetTransmittalCCUsers(ByVal tmtlID As String, comp As String) As List(Of String)
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " isnull(t_emno,'') as [user] "
        mSql = mSql & " from tdmisg031" & comp & " "
        mSql = mSql & " where t_tran = '" & tmtlID & "'"
        Dim tmp As New List(Of String)
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            While (Reader.Read())
              If Not Reader("user") = "" Then tmp.Add(Reader("user"))
            End While
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
    Public Shared Function TmtlAlert(ByVal TransmittalID As String, Comp As String) As String
      Dim documentBody As String = ""
      Dim oTmtl As EDICommon.SIS.EDI.ediTmtlH = EDICommon.SIS.EDI.ediTmtlH.GetTmtlH(TransmittalID, Comp)

      Dim aErr As New ArrayList
      Dim mRet As String = ""
      Dim oClient As SmtpClient = New SmtpClient("192.9.200.214", 25)
      Dim oMsg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage()
      oClient.Credentials = New Net.NetworkCredential("adskvaultadmin", "isgec@123")

      Dim Issuer As emp = Nothing
      Try
        Issuer = emp.GetEmp(oTmtl.t_isby, Comp)
      Catch ex As Exception
      End Try
      Dim Approver As emp = Nothing
      Try
        Approver = emp.GetEmp(oTmtl.t_apsu, Comp)
      Catch ex As Exception

      End Try
      Dim Creator As emp = Nothing
      Try
        Creator = emp.GetEmp(oTmtl.t_user, Comp)
      Catch ex As Exception

      End Try
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
              Dim IssuedTo As emp = emp.GetEmp(oTmtl.t_logn, Comp)
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
              Dim SiteIDs As String = emp.GetSiteEmailIDs(oTmtl.t_cprj, oTmtl.t_padr, Comp)
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
              Dim tmp As String = emp.GetReceiptCreator(TransmittalID, Comp)
              If tmp.ToLower = "supplier" Or tmp = "" Then
                x = gma(Approver, aErr, "VT-Approver-To")
                If x IsNot Nothing Then .To.Add(x)
              Else
                Dim tmpR As emp = emp.GetEmp(tmp, Comp)
                x = gma(tmpR, aErr, "VT-ReceiptCreator-To")
                If x IsNot Nothing Then .To.Add(x)
                x = gma(Approver, aErr, "VT-Approver-CC")
                If x IsNot Nothing Then .CC.Add(x)
              End If
              x = gma(Creator, aErr, "VT-Creator-CC")
              If x IsNot Nothing Then .CC.Add(x)
              'GetPOIssuer From Joomla
              Dim tmpPONo As String = emp.GetPONoOfReceipt(TransmittalID, Comp)
              If tmpPONo <> "" Then
                Dim POIssuer As String = emp.GetTCPOIssuer(tmpPONo)
                Dim POBuyer As String = emp.GetPOBuyer(tmpPONo)
                x = gma(emp.GetWebUser(POIssuer), aErr, "PO-Issuer")
                If x IsNot Nothing Then .CC.Add(x)
                x = gma(emp.GetWebUser(POBuyer), aErr, "PO-Buyer")
                If x IsNot Nothing Then .CC.Add(x)
              End If
              If oTmtl.t_ofbp = "SUPI00002" And oTmtl.t_issu = "007" Then
                'Transmittal Supplier Address
                Dim SiteIDs As String = emp.GetVendorEmailIDs(oTmtl.t_ofbp, oTmtl.t_vadr, Comp)
                If SiteIDs <> "" Then
                  Dim aIDs() As String = SiteIDs.Split(",;".ToCharArray)
                  For Each id As String In aIDs
                    Try
                      x = New MailAddress(id.Trim, id.Trim)
                      .To.Add(x)
                    Catch ex As Exception
                    End Try
                  Next
                End If
              End If
          End Select
          'In all Cases to Transmittal CC Users List| CR-586
          Dim ccUsers As List(Of String) = emp.GetTransmittalCCUsers(TransmittalID, Comp)
          For Each ccU As String In ccUsers
            Try
              Dim ccEmp As emp = emp.GetEmp(ccU.Trim, Comp)
              x = gma(ccEmp, aErr, "CC-Users")
              If x IsNot Nothing AndAlso Not .CC.Contains(x) Then .CC.Add(x)
            Catch ex As Exception
            End Try
          Next
          '==========================Tmtl CC Users=======================
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
        Dim tblStr As String = EDICommon.SIS.EDI.ediTmtlH.GetHTML(TransmittalID, Comp)
        Dim Header As String = ""
        Header &= "<html xmlns=""http://www.w3.org/1999/xhtml"">"
        Header &= "<head>"
        Header &= "<title></title>"
        Header &= "</head>"
        Header &= "<body>"
        documentBody = Header
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
        Header &= "<a href='http://192.9.200.146/WebEitl1/ediWTmtlH.aspx?t_tran=" & TransmittalID & "&comp=" & Comp & "'>Click to Download Transmittal Documents</a>"
        Header &= "</b></td></tr>"
        Header &= "</table>"
        Header &= "<br/>"
        Header &= "<br/>"
        Header &= tblStr
        Header &= "</body></html>"

        documentBody &= "<br/>"
        documentBody &= "<br/>"
        documentBody &= tblStr
        documentBody &= "</body></html>"

        .Body = Header
      End With
      Dim strRet As String = ""
      Try
        strRet = documentBody
        oClient.Send(oMsg)
      Catch ex As Exception
      End Try
      Return strRet
    End Function

    Public Shared Function CommentSubmittedAlertToVendor(ByVal TransmittalID As String, comp As String) As String
      Dim oTmtl As EDICommon.SIS.EDI.ediTmtlH = EDICommon.SIS.EDI.ediTmtlH.GetTmtlH(TransmittalID, comp)

      If oTmtl.t_type <> tmtlType.Vendor Then Return ""
      If oTmtl.t_ofbp = "SUPI00002" Then Return ""

      Dim oRec As EDICommon.SIS.EDI.dmIsg134 = EDICommon.SIS.EDI.dmIsg134.GetByTransmittalID(TransmittalID, comp)

      Dim aErr As New ArrayList

      Dim oClient As SmtpClient = New SmtpClient("192.9.200.214", 25)
      Dim oMsg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage()
      oClient.Credentials = New Net.NetworkCredential("adskvaultadmin", "isgec@123")

      Dim Issuer As emp = Nothing
      Try
        Issuer = emp.GetEmp(oTmtl.t_isby, comp)
      Catch ex As Exception
      End Try
      Dim Approver As emp = Nothing
      Try
        Approver = emp.GetEmp(oTmtl.t_apsu, comp)
      Catch ex As Exception

      End Try
      Dim Creator As emp = Nothing
      Try
        Creator = emp.GetEmp(oTmtl.t_user, comp)
      Catch ex As Exception

      End Try
      With oMsg
        .Subject = "Comment Submitted Receipt: " & oRec.t_rcno & "_" & oRec.t_revn & " PO: " & oRec.t_orno
        Dim EMailIDError As String = ""
        Dim x As MailAddress = Nothing
        Try
          Select Case oTmtl.t_type
            Case tmtlType.Vendor
              x = gma(Approver, aErr, "VT-Approver-From")
              If x IsNot Nothing Then
                .From = x
                .CC.Add(x)
              End If
              Dim tmp As String = emp.GetPOSupplierEmailByReceiptNo(oRec.t_rcno, oRec.t_revn)
              If tmp = "" Then
                Dim aTmp() As String = tmp.Split(",;".ToCharArray)
                For Each stmp As String In aTmp
                  stmp = stmp.Trim
                  If stmp <> "" Then
                    x = New MailAddress(stmp, stmp)
                    If Not .To.Contains(x) Then .To.Add(x)
                  End If
                Next
              End If
              x = gma(Creator, aErr, "VT-Creator-CC")
              If x IsNot Nothing Then .CC.Add(x)
              'GetPOIssuer From Joomla
              Dim tmpPONo As String = emp.GetPONoOfReceipt(TransmittalID, comp)
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
          x = New MailAddress("lalit@isgec.co.in", "Lalit Gupta")
          If Not .CC.Contains(x) Then .CC.Add(x)
          .Subject = "CommentSubmitted to supplier, TO address is empty"
        End If
        If .From Is Nothing Then
          x = New MailAddress("lalit@isgec.co.in", "Lalit Gupta")
          If Not .CC.Contains(x) Then .CC.Add(x)
          .Subject = "CommentSubmitted to supplier, FROM tmtl approver address is empty"
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
        Dim tblStr As String = EDICommon.SIS.EDI.ediTmtlH.GetHTML(TransmittalID, comp, False)
        Dim Header As String = ""
        Header &= "<html xmlns=""http://www.w3.org/1999/xhtml"">"
        Header &= "<head>"
        Header &= "<title></title>"
        Header &= "</head>"
        Header &= "<body>"
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
        Header &= "<span>Login to Vendor Portal for commented documents</span>"
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
      Return ""
    End Function

  End Class
End Namespace
