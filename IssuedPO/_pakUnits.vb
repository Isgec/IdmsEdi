Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK
  Public Class UnitMultiplicationFactor
    Public Property Unit As SIS.PAK.pakUnits = Nothing
    Public Property BaseUnit As SIS.PAK.pakUnits = Nothing
    Public Property MF As Decimal = 0
    Public Shared Function GetMultiplicationFactorToBaseUnit(ByVal mcUnit As UnitMultiplicationFactor) As UnitMultiplicationFactor
      Dim tmpUS As SIS.PAK.pakUnitSets = SIS.PAK.pakUnitSets.pakUnitSetsGetByID(mcUnit.Unit.UnitSetID)
      mcUnit.BaseUnit = SIS.PAK.pakUnits.pakUnitsGetByID(tmpUS.BaseUnitID)
      mcUnit.MF = mcUnit.Unit.ConversionFactor / mcUnit.BaseUnit.ConversionFactor
      Return mcUnit
    End Function
  End Class
  <DataObject()>
  Partial Public Class pakUnits
    Private Shared _RecordCount As Integer
    Private _UnitID As Int32 = 0
    Private _Description As String = ""
    Private _UnitSetID As String = ""
    Private _ConversionFactor As Decimal = 0
    Private _PAK_UnitSets1_Description As String = ""
    Private _FK_PAK_Units_UnitSetID As SIS.PAK.pakUnitSets = Nothing

    Public Shared Function pakUnitsGetByDescription(ByVal Description As String) As SIS.PAK.pakUnits
      Dim Results As SIS.PAK.pakUnits = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitsSelectByDescription"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 51, Description)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakUnits(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function

    Public Property UnitID() As Int32
      Get
        Return _UnitID
      End Get
      Set(ByVal value As Int32)
        _UnitID = value
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
    Public Property UnitSetID() As String
      Get
        Return _UnitSetID
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _UnitSetID = ""
        Else
          _UnitSetID = value
        End If
      End Set
    End Property
    Public Property ConversionFactor() As Decimal
      Get
        Return _ConversionFactor
      End Get
      Set(ByVal value As Decimal)
        _ConversionFactor = value
      End Set
    End Property
    Public Property PAK_UnitSets1_Description() As String
      Get
        Return _PAK_UnitSets1_Description
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _PAK_UnitSets1_Description = ""
        Else
          _PAK_UnitSets1_Description = value
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
        Return _UnitID
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
    Public Class PKpakUnits
      Private _UnitID As Int32 = 0
      Public Property UnitID() As Int32
        Get
          Return _UnitID
        End Get
        Set(ByVal value As Int32)
          _UnitID = value
        End Set
      End Property
    End Class
    Public ReadOnly Property FK_PAK_Units_UnitSetID() As SIS.PAK.pakUnitSets
      Get
        If _FK_PAK_Units_UnitSetID Is Nothing Then
          _FK_PAK_Units_UnitSetID = SIS.PAK.pakUnitSets.pakUnitSetsGetByID(_UnitSetID)
        End If
        Return _FK_PAK_Units_UnitSetID
      End Get
    End Property
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakUnitsSelectList(ByVal OrderBy As String) As List(Of SIS.PAK.pakUnits)
      Dim Results As List(Of SIS.PAK.pakUnits) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitsSelectList"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakUnits)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakUnits(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakUnitsGetNewRecord() As SIS.PAK.pakUnits
      Return New SIS.PAK.pakUnits()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakUnitsGetByID(ByVal UnitID As Int32) As SIS.PAK.pakUnits
      Dim Results As SIS.PAK.pakUnits = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitsSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UnitID", SqlDbType.Int, UnitID.ToString.Length, UnitID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakUnits(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakUnitsSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal UnitSetID As Int32) As List(Of SIS.PAK.pakUnits)
      Dim Results As List(Of SIS.PAK.pakUnits) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "sppakUnitsSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "sppakUnitsSelectListFilteres"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_UnitSetID", SqlDbType.Int, 10, IIf(UnitSetID = Nothing, 0, UnitSetID))
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakUnits)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakUnits(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function pakUnitsSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String, ByVal UnitSetID As Int32) As Integer
      Return _RecordCount
    End Function
    'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakUnitsGetByID(ByVal UnitID As Int32, ByVal Filter_UnitSetID As Int32) As SIS.PAK.pakUnits
      Return pakUnitsGetByID(UnitID)
    End Function
    <DataObjectMethod(DataObjectMethodType.Insert, True)>
    Public Shared Function pakUnitsInsert(ByVal Record As SIS.PAK.pakUnits) As SIS.PAK.pakUnits
      Dim _Rec As SIS.PAK.pakUnits = SIS.PAK.pakUnits.pakUnitsGetNewRecord()
      With _Rec
        .Description = Record.Description
        .UnitSetID = Record.UnitSetID
        .ConversionFactor = Record.ConversionFactor
      End With
      Return SIS.PAK.pakUnits.InsertData(_Rec)
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.PAK.pakUnits) As SIS.PAK.pakUnits
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitsInsert"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 51, IIf(Record.Description = "", Convert.DBNull, Record.Description))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UnitSetID", SqlDbType.Int, 11, IIf(Record.UnitSetID = "", Convert.DBNull, Record.UnitSetID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ConversionFactor", SqlDbType.Decimal, 27, Record.ConversionFactor)
          Cmd.Parameters.Add("@Return_UnitID", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_UnitID").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.UnitID = Cmd.Parameters("@Return_UnitID").Value
        End Using
      End Using
      Return Record
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)>
    Public Shared Function pakUnitsUpdate(ByVal Record As SIS.PAK.pakUnits) As SIS.PAK.pakUnits
      Dim _Rec As SIS.PAK.pakUnits = SIS.PAK.pakUnits.pakUnitsGetByID(Record.UnitID)
      With _Rec
        .Description = Record.Description
        .UnitSetID = Record.UnitSetID
        .ConversionFactor = Record.ConversionFactor
      End With
      Return SIS.PAK.pakUnits.UpdateData(_Rec)
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.PAK.pakUnits) As SIS.PAK.pakUnits
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitsUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_UnitID", SqlDbType.Int, 11, Record.UnitID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 51, IIf(Record.Description = "", Convert.DBNull, Record.Description))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UnitSetID", SqlDbType.Int, 11, IIf(Record.UnitSetID = "", Convert.DBNull, Record.UnitSetID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ConversionFactor", SqlDbType.Decimal, 27, Record.ConversionFactor)
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
    Public Shared Function pakUnitsDelete(ByVal Record As SIS.PAK.pakUnits) As Int32
      Dim _Result As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakUnitsDelete"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_UnitID", SqlDbType.Int, Record.UnitID.ToString.Length, Record.UnitID)
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
