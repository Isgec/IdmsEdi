Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK
  <DataObject()> _
  Partial Public Class pakDocuments
    Private Shared _RecordCount As Integer
    Private _DocumentNo As Int32 = 0
    Private _DocumentID As String = ""
    Private _DocumentRevision As String = ""
    Private _Description As String = ""
    Private _ExternalDocument As Boolean = False
    Private _Filename As String = ""
    Private _DisskFile As String = ""
    Public Property DocumentNo() As Int32
      Get
        Return _DocumentNo
      End Get
      Set(ByVal value As Int32)
        _DocumentNo = value
      End Set
    End Property
    Public Property DocumentID() As String
      Get
        Return _DocumentID
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _DocumentID = ""
        Else
          _DocumentID = value
        End If
      End Set
    End Property
    Public Property DocumentRevision() As String
      Get
        Return _DocumentRevision
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _DocumentRevision = ""
        Else
          _DocumentRevision = value
        End If
      End Set
    End Property
    Public Property Description() As String
      Get
        Return _Description
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _Description = ""
        Else
          _Description = value
        End If
      End Set
    End Property
    Public Property ExternalDocument() As Boolean
      Get
        Return _ExternalDocument
      End Get
      Set(ByVal value As Boolean)
        _ExternalDocument = value
      End Set
    End Property
    Public Property Filename() As String
      Get
        Return _Filename
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _Filename = ""
        Else
          _Filename = value
        End If
      End Set
    End Property
    Public Property DisskFile() As String
      Get
        Return _DisskFile
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _DisskFile = ""
        Else
          _DisskFile = value
        End If
      End Set
    End Property
    Public ReadOnly Property DisplayField() As String
      Get
        Return "" & _Description.ToString.PadRight(100, " ")
      End Get
    End Property
    Public ReadOnly Property PrimaryKey() As String
      Get
        Return _DocumentNo
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
    Public Class PKpakDocuments
      Private _DocumentNo As Int32 = 0
      Public Property DocumentNo() As Int32
        Get
          Return _DocumentNo
        End Get
        Set(ByVal value As Int32)
          _DocumentNo = value
        End Set
      End Property
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakDocumentsSelectList(ByVal OrderBy As String) As List(Of SIS.PAK.pakDocuments)
      Dim Results As List(Of SIS.PAK.pakDocuments) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakDocumentsSelectList"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakDocuments)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakDocuments(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakDocumentsGetNewRecord() As SIS.PAK.pakDocuments
      Return New SIS.PAK.pakDocuments()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakDocumentsGetByID(ByVal DocumentNo As Int32) As SIS.PAK.pakDocuments
      Dim Results As SIS.PAK.pakDocuments = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakDocumentsSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentNo", SqlDbType.Int, DocumentNo.ToString.Length, DocumentNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakDocuments(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakDocumentsSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String) As List(Of SIS.PAK.pakDocuments)
      Dim Results As List(Of SIS.PAK.pakDocuments) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "sppakDocumentsSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "sppakDocumentsSelectListFilteres"
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakDocuments)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakDocuments(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function pakDocumentsSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String) As Integer
      Return _RecordCount
    End Function
    'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Insert, True)>
    Public Shared Function pakDocumentsInsert(ByVal Record As SIS.PAK.pakDocuments) As SIS.PAK.pakDocuments
      Dim _Rec As SIS.PAK.pakDocuments = SIS.PAK.pakDocuments.pakDocumentsGetNewRecord()
      With _Rec
        .DocumentID = Record.DocumentID
        .DocumentRevision = Record.DocumentRevision
        .Description = Record.Description
        .ExternalDocument = Record.ExternalDocument
        .Filename = Record.Filename
        .DisskFile = Record.DisskFile
      End With
      Return SIS.PAK.pakDocuments.InsertData(_Rec)
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.PAK.pakDocuments) As SIS.PAK.pakDocuments
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakDocumentsInsert"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentID", SqlDbType.NVarChar, 51, IIf(Record.DocumentID = "", Convert.DBNull, Record.DocumentID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentRevision", SqlDbType.NVarChar, 11, IIf(Record.DocumentRevision = "", Convert.DBNull, Record.DocumentRevision))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 101, IIf(Record.Description = "", Convert.DBNull, Record.Description))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ExternalDocument", SqlDbType.Bit, 3, Record.ExternalDocument)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Filename", SqlDbType.NVarChar, 101, IIf(Record.Filename = "", Convert.DBNull, Record.Filename))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DisskFile", SqlDbType.NVarChar, 51, IIf(Record.DisskFile = "", Convert.DBNull, Record.DisskFile))
          Cmd.Parameters.Add("@Return_DocumentNo", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_DocumentNo").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.DocumentNo = Cmd.Parameters("@Return_DocumentNo").Value
        End Using
      End Using
      Return Record
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)>
    Public Shared Function pakDocumentsUpdate(ByVal Record As SIS.PAK.pakDocuments) As SIS.PAK.pakDocuments
      Dim _Rec As SIS.PAK.pakDocuments = SIS.PAK.pakDocuments.pakDocumentsGetByID(Record.DocumentNo)
      With _Rec
        .DocumentID = Record.DocumentID
        .DocumentRevision = Record.DocumentRevision
        .Description = Record.Description
        .ExternalDocument = Record.ExternalDocument
        .Filename = Record.Filename
        .DisskFile = Record.DisskFile
      End With
      Return SIS.PAK.pakDocuments.UpdateData(_Rec)
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.PAK.pakDocuments) As SIS.PAK.pakDocuments
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakDocumentsUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_DocumentNo", SqlDbType.Int, 11, Record.DocumentNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentID", SqlDbType.NVarChar, 51, IIf(Record.DocumentID = "", Convert.DBNull, Record.DocumentID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentRevision", SqlDbType.NVarChar, 11, IIf(Record.DocumentRevision = "", Convert.DBNull, Record.DocumentRevision))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 101, IIf(Record.Description = "", Convert.DBNull, Record.Description))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ExternalDocument", SqlDbType.Bit, 3, Record.ExternalDocument)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Filename", SqlDbType.NVarChar, 101, IIf(Record.Filename = "", Convert.DBNull, Record.Filename))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DisskFile", SqlDbType.NVarChar, 51, IIf(Record.DisskFile = "", Convert.DBNull, Record.DisskFile))
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
    Public Shared Function pakDocumentsDelete(ByVal Record As SIS.PAK.pakDocuments) As Int32
      Dim _Result As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakDocumentsDelete"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_DocumentNo", SqlDbType.Int, Record.DocumentNo.ToString.Length, Record.DocumentNo)
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
    '    Autocomplete Method
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
