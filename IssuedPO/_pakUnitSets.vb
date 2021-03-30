Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK
  <DataObject()> _
  Partial Public Class pakUnitSets
    Private Shared _RecordCount As Integer
    Private _UnitSetID As Int32 = 0
    Private _Description As String = ""
    Private _BaseUnitID As String = ""
    Private _PAK_Units1_Description As String = ""
    Private _FK_PAK_UnitSets_BaseUnitID As SIS.PAK.pakUnits = Nothing
    Public Property UnitSetID() As Int32
      Get
        Return _UnitSetID
      End Get
      Set(ByVal value As Int32)
        _UnitSetID = value
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
    Public Property BaseUnitID() As String
      Get
        Return _BaseUnitID
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _BaseUnitID = ""
        Else
          _BaseUnitID = value
        End If
      End Set
    End Property
    Public Property PAK_Units1_Description() As String
      Get
        Return _PAK_Units1_Description
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _PAK_Units1_Description = ""
        Else
          _PAK_Units1_Description = value
        End If
      End Set
    End Property
    Public ReadOnly Property DisplayField() As String
      Get
        Return "" & _Description.ToString.PadRight(50, " ")
      End Get
    End Property
    Public ReadOnly Property PrimaryKey() As String
      Get
        Return _UnitSetID
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
    Public Class PKpakUnitSets
      Private _UnitSetID As Int32 = 0
      Public Property UnitSetID() As Int32
        Get
          Return _UnitSetID
        End Get
        Set(ByVal value As Int32)
          _UnitSetID = value
        End Set
      End Property
    End Class
    Public ReadOnly Property FK_PAK_UnitSets_BaseUnitID() As SIS.PAK.pakUnits
      Get
        If _FK_PAK_UnitSets_BaseUnitID Is Nothing Then
          _FK_PAK_UnitSets_BaseUnitID = SIS.PAK.pakUnits.pakUnitsGetByID(_BaseUnitID)
        End If
        Return _FK_PAK_UnitSets_BaseUnitID
      End Get
    End Property
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakUnitSetsSelectList(ByVal OrderBy As String) As List(Of SIS.PAK.pakUnitSets)
      Dim Results As List(Of SIS.PAK.pakUnitSets) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitSetsSelectList"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakUnitSets)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakUnitSets(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakUnitSetsGetNewRecord() As SIS.PAK.pakUnitSets
      Return New SIS.PAK.pakUnitSets()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakUnitSetsGetByID(ByVal UnitSetID As Int32) As SIS.PAK.pakUnitSets
      Dim Results As SIS.PAK.pakUnitSets = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitSetsSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UnitSetID", SqlDbType.Int, UnitSetID.ToString.Length, UnitSetID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakUnitSets(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakUnitSetsSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String) As List(Of SIS.PAK.pakUnitSets)
      Dim Results As List(Of SIS.PAK.pakUnitSets) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "sppakUnitSetsSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "sppakUnitSetsSelectListFilteres"
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakUnitSets)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakUnitSets(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function pakUnitSetsSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String) As Integer
      Return _RecordCount
    End Function
    'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Insert, True)>
    Public Shared Function pakUnitSetsInsert(ByVal Record As SIS.PAK.pakUnitSets) As SIS.PAK.pakUnitSets
      Dim _Rec As SIS.PAK.pakUnitSets = SIS.PAK.pakUnitSets.pakUnitSetsGetNewRecord()
      With _Rec
        .Description = Record.Description
        .BaseUnitID = Record.BaseUnitID
      End With
      Return SIS.PAK.pakUnitSets.InsertData(_Rec)
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.PAK.pakUnitSets) As SIS.PAK.pakUnitSets
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitSetsInsert"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 51, IIf(Record.Description = "", Convert.DBNull, Record.Description))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@BaseUnitID", SqlDbType.Int, 11, IIf(Record.BaseUnitID = "", Convert.DBNull, Record.BaseUnitID))
          Cmd.Parameters.Add("@Return_UnitSetID", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_UnitSetID").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.UnitSetID = Cmd.Parameters("@Return_UnitSetID").Value
        End Using
      End Using
      Return Record
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)>
    Public Shared Function pakUnitSetsUpdate(ByVal Record As SIS.PAK.pakUnitSets) As SIS.PAK.pakUnitSets
      Dim _Rec As SIS.PAK.pakUnitSets = SIS.PAK.pakUnitSets.pakUnitSetsGetByID(Record.UnitSetID)
      With _Rec
        .Description = Record.Description
        .BaseUnitID = Record.BaseUnitID
      End With
      Return SIS.PAK.pakUnitSets.UpdateData(_Rec)
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.PAK.pakUnitSets) As SIS.PAK.pakUnitSets
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitSetsUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_UnitSetID", SqlDbType.Int, 11, Record.UnitSetID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 51, IIf(Record.Description = "", Convert.DBNull, Record.Description))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@BaseUnitID", SqlDbType.Int, 11, IIf(Record.BaseUnitID = "", Convert.DBNull, Record.BaseUnitID))
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
    Public Shared Function pakUnitSetsDelete(ByVal Record As SIS.PAK.pakUnitSets) As Int32
      Dim _Result As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitSetsDelete"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_UnitSetID", SqlDbType.Int, Record.UnitSetID.ToString.Length, Record.UnitSetID)
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
