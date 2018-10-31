Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.EDI
  <DataObject()>
  Partial Public Class ediAFile
    Implements ICloneable
    Public Function Clone() As Object Implements ICloneable.Clone
      Dim tmp As ediAFile = MyBase.MemberwiseClone()
      Return tmp
    End Function

    Private Shared _RecordCount As Integer
    Private _t_drid As String = ""
    Private _t_dcid As String = ""
    Private _t_hndl As String = ""
    Private _t_indx As String = ""
    Private _t_prcd As String = ""
    Private _t_fnam As String = ""
    Private _t_lbcd As String = ""
    Private _t_atby As String = ""
    Private _t_aton As String = ""
    Private _t_Refcntd As Int32 = 0
    Private _t_Refcntu As Int32 = 0
    Public Property t_drid() As String
      Get
        Return _t_drid
      End Get
      Set(ByVal value As String)
        _t_drid = value
      End Set
    End Property
    Public Property t_dcid() As String
      Get
        Return _t_dcid
      End Get
      Set(ByVal value As String)
        _t_dcid = value
      End Set
    End Property
    Public Property t_hndl() As String
      Get
        Return _t_hndl
      End Get
      Set(ByVal value As String)
        _t_hndl = value
      End Set
    End Property
    Public Property t_indx() As String
      Get
        Return _t_indx
      End Get
      Set(ByVal value As String)
        _t_indx = value
      End Set
    End Property
    Public Property t_prcd() As String
      Get
        Return _t_prcd
      End Get
      Set(ByVal value As String)
        _t_prcd = value
      End Set
    End Property
    Public Property t_fnam() As String
      Get
        Return _t_fnam
      End Get
      Set(ByVal value As String)
        _t_fnam = value
      End Set
    End Property
    Public Property t_lbcd() As String
      Get
        Return _t_lbcd
      End Get
      Set(ByVal value As String)
        _t_lbcd = value
      End Set
    End Property
    Public Property t_atby() As String
      Get
        Return _t_atby
      End Get
      Set(ByVal value As String)
        _t_atby = value
      End Set
    End Property
    Public Property t_aton() As String
      Get
        If Not _t_aton = String.Empty Then
          Return Convert.ToDateTime(_t_aton).ToString("dd/MM/yyyy")
        End If
        Return _t_aton
      End Get
      Set(ByVal value As String)
        _t_aton = value
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
    Public ReadOnly Property DisplayField() As String
      Get
        Return ""
      End Get
    End Property
    Public ReadOnly Property PrimaryKey() As String
      Get
        Return _t_drid
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
    Public Class PKediAFile
      Private _t_drid As String = ""
      Public Property t_drid() As String
        Get
          Return _t_drid
        End Get
        Set(ByVal value As String)
          _t_drid = value
        End Set
      End Property
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediAFileGetNewRecord() As SIS.EDI.ediAFile
      Return New SIS.EDI.ediAFile()
    End Function
    Public Shared Function ediAFileGetByHandleIndex(ByVal t_hndl As String, ByVal t_indx As String) As SIS.EDI.ediAFile
      Dim Results As SIS.EDI.ediAFile = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select top 1 * from ttcisg132200 where t_hndl='" & t_hndl & "' and t_indx='" & t_indx & "'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.EDI.ediAFile(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediAFileGetByID(ByVal t_drid As String) As SIS.EDI.ediAFile
      Dim Results As SIS.EDI.ediAFile = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediAFileSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_drid", SqlDbType.VarChar, t_drid.ToString.Length, t_drid)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "0340")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.EDI.ediAFile(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediAFileSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal t_drid As String, ByVal t_hndl As String, ByVal t_indx As String) As List(Of SIS.EDI.ediAFile)
      Dim Results As List(Of SIS.EDI.ediAFile) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "spediAFileSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "spediAFileSelectListFilteres"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_t_drid", SqlDbType.VarChar, 50, IIf(t_drid Is Nothing, String.Empty, t_drid))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_t_hndl", SqlDbType.VarChar, 200, IIf(t_hndl Is Nothing, String.Empty, t_hndl))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_t_indx", SqlDbType.VarChar, 50, IIf(t_indx Is Nothing, String.Empty, t_indx))
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "0340")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.EDI.ediAFile)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.EDI.ediAFile(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function ediAFileSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String, ByVal t_drid As String, ByVal t_hndl As String, ByVal t_indx As String) As Integer
      Return _RecordCount
    End Function
    'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediAFileGetByID(ByVal t_drid As String, ByVal Filter_t_drid As String, ByVal Filter_t_hndl As String, ByVal Filter_t_indx As String) As SIS.EDI.ediAFile
      Return ediAFileGetByID(t_drid)
    End Function
    <DataObjectMethod(DataObjectMethodType.Insert, True)>
    Public Shared Function ediAFileInsert(ByVal Record As SIS.EDI.ediAFile) As SIS.EDI.ediAFile
      Dim _Rec As SIS.EDI.ediAFile = SIS.EDI.ediAFile.ediAFileGetNewRecord()
      With _Rec
        .t_drid = Record.t_drid
        .t_dcid = Record.t_dcid
        .t_hndl = Record.t_hndl
        .t_indx = Record.t_indx
        .t_prcd = Record.t_prcd
        .t_fnam = Record.t_fnam
        .t_lbcd = Record.t_lbcd
        .t_atby = Record.t_atby
        .t_aton = Record.t_aton
        .t_Refcntd = Record.t_Refcntd
        .t_Refcntu = Record.t_Refcntu
      End With
      Return SIS.EDI.ediAFile.InsertData(_Rec)
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.EDI.ediAFile) As SIS.EDI.ediAFile
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediAFileInsert"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_dcid", SqlDbType.VarChar, 201, Record.t_dcid)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_hndl", SqlDbType.VarChar, 201, Record.t_hndl)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_indx", SqlDbType.VarChar, 51, Record.t_indx)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_prcd", SqlDbType.VarChar, 51, Record.t_prcd)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_fnam", SqlDbType.VarChar, 251, Record.t_fnam)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_lbcd", SqlDbType.VarChar, 51, Record.t_lbcd)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_atby", SqlDbType.VarChar, 51, Record.t_atby)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_aton", SqlDbType.DateTime, 21, Record.t_aton)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_Refcntd", SqlDbType.Int, 11, Record.t_Refcntd)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_Refcntu", SqlDbType.Int, 11, Record.t_Refcntu)
          Cmd.Parameters.Add("@Return_t_drid", SqlDbType.VarChar, 51)
          Cmd.Parameters("@Return_t_drid").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.t_drid = Cmd.Parameters("@Return_t_drid").Value
        End Using
      End Using
      Return Record
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)>
    Public Shared Function ediAFileUpdate(ByVal Record As SIS.EDI.ediAFile) As SIS.EDI.ediAFile
      Dim _Rec As SIS.EDI.ediAFile = SIS.EDI.ediAFile.ediAFileGetByID(Record.t_drid)
      With _Rec
        .t_dcid = Record.t_dcid
        .t_hndl = Record.t_hndl
        .t_indx = Record.t_indx
        .t_prcd = Record.t_prcd
        .t_fnam = Record.t_fnam
        .t_lbcd = Record.t_lbcd
        .t_atby = Record.t_atby
        .t_aton = Record.t_aton
        .t_Refcntd = Record.t_Refcntd
        .t_Refcntu = Record.t_Refcntu
      End With
      Return SIS.EDI.ediAFile.UpdateData(_Rec)
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.EDI.ediAFile) As SIS.EDI.ediAFile
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediAFileUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_t_drid", SqlDbType.VarChar, 51, Record.t_drid)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_dcid", SqlDbType.VarChar, 201, Record.t_dcid)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_hndl", SqlDbType.VarChar, 201, Record.t_hndl)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_indx", SqlDbType.VarChar, 51, Record.t_indx)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_prcd", SqlDbType.VarChar, 51, Record.t_prcd)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_fnam", SqlDbType.VarChar, 251, Record.t_fnam)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_lbcd", SqlDbType.VarChar, 51, Record.t_lbcd)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_atby", SqlDbType.VarChar, 51, Record.t_atby)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_aton", SqlDbType.DateTime, 21, Record.t_aton)
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
    <DataObjectMethod(DataObjectMethodType.Delete, True)>
    Public Shared Function ediAFileDelete(ByVal Record As SIS.EDI.ediAFile) As Int32
      Dim _Result As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediAFileDelete"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_t_drid", SqlDbType.VarChar, Record.t_drid.ToString.Length, Record.t_drid)
          Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Con.Open()
          Cmd.ExecuteNonQuery()
          _RecordCount = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Return _RecordCount
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
