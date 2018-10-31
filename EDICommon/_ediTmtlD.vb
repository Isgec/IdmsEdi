Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.EDI
  <DataObject()> _
  Partial Public Class ediTmtlD
    Private Shared _RecordCount As Integer
    Private _t_tran As String = ""
    Private _t_docn As String = ""
    Private _t_revn As String = ""
    Private _t_stid As String = ""
    Private _t_pono As Int32 = 0
    Private _t_remk As String = ""
    Private _t_recv As String = ""
    Private _t_refr As String = ""
    Private _t_redt As String = ""
    Private _t_rekm As String = ""
    Private _t_lock As Int32 = 0
    Private _t_recc As Int32 = 0
    Private _t_revd As Int32 = 0
    Private _t_issu As Int32 = 0
    Public Property t_dttl As String = ""
    Private _t_Refcntd As Int32 = 0
    Private _t_Refcntu As Int32 = 0
    Public Property t_tran() As String
      Get
        Return _t_tran
      End Get
      Set(ByVal value As String)
        _t_tran = value
      End Set
    End Property
    Public Property t_docn() As String
      Get
        Return _t_docn
      End Get
      Set(ByVal value As String)
        _t_docn = value
      End Set
    End Property
    Public Property t_revn() As String
      Get
        Return _t_revn
      End Get
      Set(ByVal value As String)
        _t_revn = value
      End Set
    End Property
    Public Property t_stid() As String
      Get
        Return _t_stid
      End Get
      Set(ByVal value As String)
        _t_stid = value
      End Set
    End Property
    Public Property t_pono() As Int32
      Get
        Return _t_pono
      End Get
      Set(ByVal value As Int32)
        _t_pono = value
      End Set
    End Property
    Public Property t_remk() As String
      Get
        Return _t_remk
      End Get
      Set(ByVal value As String)
        _t_remk = value
      End Set
    End Property
    Public Property t_recv() As String
      Get
        Return _t_recv
      End Get
      Set(ByVal value As String)
        _t_recv = value
      End Set
    End Property
    Public Property t_refr() As String
      Get
        Return _t_refr
      End Get
      Set(ByVal value As String)
        _t_refr = value
      End Set
    End Property
    Public Property t_redt() As String
      Get
        If Not _t_redt = String.Empty Then
          Return Convert.ToDateTime(_t_redt).ToString("dd/MM/yyyy")
        End If
        Return _t_redt
      End Get
      Set(ByVal value As String)
         _t_redt = value
      End Set
    End Property
    Public Property t_rekm() As String
      Get
        Return _t_rekm
      End Get
      Set(ByVal value As String)
        _t_rekm = value
      End Set
    End Property
    Public Property t_lock() As Int32
      Get
        Return _t_lock
      End Get
      Set(ByVal value As Int32)
        _t_lock = value
      End Set
    End Property
    Public Property t_recc() As Int32
      Get
        Return _t_recc
      End Get
      Set(ByVal value As Int32)
        _t_recc = value
      End Set
    End Property
    Public Property t_revd() As Int32
      Get
        Return _t_revd
      End Get
      Set(ByVal value As Int32)
        _t_revd = value
      End Set
    End Property
    Public Property t_issu() As Int32
      Get
        Return _t_issu
      End Get
      Set(ByVal value As Int32)
        _t_issu = value
      End Set
    End Property
    Public Property t_Refcntd() As Int32
      Get
        Return _t_Refcntd
      End Get
      Set(ByVal value As Int32)
        _t_Refcntd = value
      End Set
    End Property
    Public Property t_Refcntu() As Int32
      Get
        Return _t_Refcntu
      End Get
      Set(ByVal value As Int32)
        _t_Refcntu = value
      End Set
    End Property
    Public Readonly Property DisplayField() As String
      Get
        Return ""
      End Get
    End Property
    Public Readonly Property PrimaryKey() As String
      Get
        Return _t_tran & "|" & _t_docn & "|" & _t_revn
      End Get
    End Property
    Public Shared Property RecordCount() As Integer
      Get
        Return _RecordCount
      End Get
      Set(ByVal value As Integer)
        _RecordCount = value
      End Set
    End Property
    Public Class PKediTmtlD
      Private _t_tran As String = ""
      Private _t_docn As String = ""
      Private _t_revn As String = ""
      Public Property t_tran() As String
        Get
          Return _t_tran
        End Get
        Set(ByVal value As String)
          _t_tran = value
        End Set
      End Property
      Public Property t_docn() As String
        Get
          Return _t_docn
        End Get
        Set(ByVal value As String)
          _t_docn = value
        End Set
      End Property
      Public Property t_revn() As String
        Get
          Return _t_revn
        End Get
        Set(ByVal value As String)
          _t_revn = value
        End Set
      End Property
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function ediTmtlDGetNewRecord() As SIS.EDI.ediTmtlD
      Return New SIS.EDI.ediTmtlD()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function ediTmtlDGetByID(ByVal t_tran As String, ByVal t_docn As String, ByVal t_revn As String) As SIS.EDI.ediTmtlD
      Dim Results As SIS.EDI.ediTmtlD = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediTmtlDSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_tran", SqlDbType.VarChar, t_tran.ToString.Length, t_tran)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_docn", SqlDbType.VarChar, t_docn.ToString.Length, t_docn)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_revn", SqlDbType.VarChar, t_revn.ToString.Length, t_revn)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "0340")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.EDI.ediTmtlD(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function ediTmtlDSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal t_tran As String, ByVal t_docn As String, ByVal t_revn As String) As List(Of SIS.EDI.ediTmtlD)
      Dim Results As List(Of SIS.EDI.ediTmtlD) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "spediTmtlDSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "spediTmtlDSelectListFilteres"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_t_tran", SqlDbType.VarChar, 9, IIf(t_tran Is Nothing, String.Empty, t_tran))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_t_docn", SqlDbType.VarChar, 32, IIf(t_docn Is Nothing, String.Empty, t_docn))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_t_revn", SqlDbType.VarChar, 20, IIf(t_revn Is Nothing, String.Empty, t_revn))
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "0340")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.EDI.ediTmtlD)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.EDI.ediTmtlD(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function ediTmtlDSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String, ByVal t_tran As String, ByVal t_docn As String, ByVal t_revn As String) As Integer
      Return _RecordCount
    End Function
      'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function ediTmtlDGetByID(ByVal t_tran As String, ByVal t_docn As String, ByVal t_revn As String, ByVal Filter_t_tran As String, ByVal Filter_t_docn As String, ByVal Filter_t_revn As String) As SIS.EDI.ediTmtlD
      Return ediTmtlDGetByID(t_tran, t_docn, t_revn)
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)>
    Public Shared Function ediTmtlDUpdate(ByVal Record As SIS.EDI.ediTmtlD) As SIS.EDI.ediTmtlD
      Dim _Rec As SIS.EDI.ediTmtlD = SIS.EDI.ediTmtlD.ediTmtlDGetByID(Record.t_tran, Record.t_docn, Record.t_revn)
      With _Rec
        .t_stid = Record.t_stid
        .t_pono = Record.t_pono
        .t_remk = Record.t_remk
        .t_recv = Record.t_recv
        .t_refr = Record.t_refr
        .t_redt = Record.t_redt
        .t_rekm = Record.t_rekm
        .t_lock = Record.t_lock
        .t_recc = Record.t_recc
        .t_revd = Record.t_revd
        .t_issu = Record.t_issu
        .t_Refcntd = Record.t_Refcntd
        .t_Refcntu = Record.t_Refcntu
      End With
      Return SIS.EDI.ediTmtlD.UpdateData(_Rec)
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.EDI.ediTmtlD) As SIS.EDI.ediTmtlD
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediTmtlDUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_t_tran", SqlDbType.VarChar, 10, Record.t_tran)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_t_docn", SqlDbType.VarChar, 33, Record.t_docn)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_t_revn", SqlDbType.VarChar, 21, Record.t_revn)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_tran", SqlDbType.VarChar, 10, Record.t_tran)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_docn", SqlDbType.VarChar, 33, Record.t_docn)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_revn", SqlDbType.VarChar, 21, Record.t_revn)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_stid", SqlDbType.VarChar, 4, Record.t_stid)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_pono", SqlDbType.Int, 11, Record.t_pono)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_remk", SqlDbType.VarChar, 101, Record.t_remk)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_recv", SqlDbType.VarChar, 4, Record.t_recv)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_refr", SqlDbType.VarChar, 31, Record.t_refr)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_redt", SqlDbType.DateTime, 21, Record.t_redt)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_rekm", SqlDbType.VarChar, 101, Record.t_rekm)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_lock", SqlDbType.Int, 11, Record.t_lock)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_recc", SqlDbType.Int, 11, Record.t_recc)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_revd", SqlDbType.Int, 11, Record.t_revd)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_issu", SqlDbType.Int, 11, Record.t_issu)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_Refcntd", SqlDbType.Int, 11, Record.t_Refcntd)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_Refcntu", SqlDbType.Int, 11, Record.t_Refcntu)
          Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Con.Open()
          Cmd.ExecuteNonQuery()
          _RecordCount = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Return Record
    End Function
    Public Sub New(ByVal Reader As SqlDataReader)
      Try
        For Each pi As System.Reflection.PropertyInfo In Me.GetType.GetProperties
          If pi.MemberType = Reflection.MemberTypes.Property Then
            Try
              Dim Found As Boolean = False
              For I As Integer = 0 To Reader.FieldCount - 1
                If Reader.GetName(I).ToLower = pi.Name.ToLower Then
                  Found = True
                  Exit For
                End If
              Next
              If Found Then
                If Convert.IsDBNull(Reader(pi.Name)) Then
                  Select Case Reader.GetDataTypeName(Reader.GetOrdinal(pi.Name))
                    Case "decimal"
                      CallByName(Me, pi.Name, CallType.Let, "0.00")
                    Case "bit"
                      CallByName(Me, pi.Name, CallType.Let, Boolean.FalseString)
                    Case Else
                      CallByName(Me, pi.Name, CallType.Let, String.Empty)
                  End Select
                Else
                  CallByName(Me, pi.Name, CallType.Let, Reader(pi.Name))
                End If
              End If
            Catch ex As Exception
            End Try
          End If
        Next
      Catch ex As Exception
      End Try
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace
