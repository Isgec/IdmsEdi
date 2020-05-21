Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.EDI
  Public Class dmIsg134
    Public Property t_rcno As String = ""
    Public Property t_revn As String = ""
    Public Property t_cprj As String = ""
    Public Property t_item As String = ""
    Public Property t_bpid As String = ""
    Public Property t_nama As String = ""
    Public Property t_stat As Integer = 0
    Public Property t_user As String = ""
    Public Property t_date As DateTime = ""
    Public Property t_sent_1 As Integer = 0
    Public Property t_sent_2 As Integer = 0
    Public Property t_sent_3 As Integer = 0
    Public Property t_sent_4 As Integer = 0
    Public Property t_sent_5 As Integer = 0
    Public Property t_sent_6 As Integer = 0
    Public Property t_sent_7 As Integer = 0
    Public Property t_rece_1 As Integer = 0
    Public Property t_rece_2 As Integer = 0
    Public Property t_rece_3 As Integer = 0
    Public Property t_rece_4 As Integer = 0
    Public Property t_rece_5 As Integer = 0
    Public Property t_rece_6 As Integer = 0
    Public Property t_rece_7 As Integer = 0
    Public Property t_suer As String = ""
    Public Property t_sdat As DateTime = ""
    Public Property t_appr As String = ""
    Public Property t_adat As DateTime = ""
    Public Property t_subm_1 As Integer = 0
    Public Property t_subm_2 As Integer = 0
    Public Property t_subm_3 As Integer = 0
    Public Property t_subm_4 As Integer = 0
    Public Property t_subm_5 As Integer = 0
    Public Property t_subm_6 As Integer = 0
    Public Property t_subm_7 As Integer = 0
    Public Property t_orno As String = ""
    Public Property t_pono As Integer = 0
    Public Property t_trno As String = ""
    Public Property t_Refcntd As Integer = 0
    Public Property t_Refcntu As Integer = 0
    Public Property t_docn As String = ""
    Public Property t_eunt As String = ""
    Public Property t_atch As String = ""
    Public Property t_rqln As Integer = 0
    Public Property t_rqno As String = ""
    Public Property t_pwfd As Integer = 0
    Public Property t_wfid As Integer = 0
    Public Property t_apid_1 As String = ""
    Public Property t_apid_2 As String = ""
    Public Property t_apid_3 As String = ""
    Public Property t_apid_4 As String = ""
    Public Property t_apid_5 As String = ""
    Public Property t_apid_6 As String = ""
    Public Property t_apid_7 As String = ""
    Public Shared Function GetByTransmittalID(TmtlID As String, Optional ERPCompany As String = "200") As SIS.EDI.dmIsg134
      Dim mRet As SIS.EDI.dmIsg134 = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "select * from tdmisg132" & ERPCompany & " where t_trno='" & TmtlID & "'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            mRet = New SIS.EDI.dmIsg134(Reader)
          End If
          Reader.Close()
        End Using
      End Using

      Return mRet
    End Function
    Sub New(rd As SqlDataReader)
      DBCommon.NewObj(Me, rd)
    End Sub
    Sub New()
      'Dummy
    End Sub

  End Class

End Namespace
