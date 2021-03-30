Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK
  <DataObject()> _
  Partial Public Class pakElements
    Private Shared _RecordCount As Integer
    Private _ElementID As String = ""
    Private _Description As String = ""
    Private _ResponsibleAgencyID As String = ""
    Private _ParentElementID As String = ""
    Private _ElementLevel As String = ""
    Private _Prefix As String = ""
    Private _PAK_Elements1_Description As String = ""
    Private _PAK_ResponsibleAgencies2_Description As String = ""
    Private _FK_PAK_Elements_ParentElementID As SIS.PAK.pakElements = Nothing
    Public Property ElementID() As String
      Get
        Return _ElementID
      End Get
      Set(ByVal value As String)
        _ElementID = value
      End Set
    End Property
    Public Property Description() As String
      Get
        Return _Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _Description = ""
         Else
           _Description = value
         End If
      End Set
    End Property
    Public Property ResponsibleAgencyID() As String
      Get
        Return _ResponsibleAgencyID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ResponsibleAgencyID = ""
         Else
           _ResponsibleAgencyID = value
         End If
      End Set
    End Property
    Public Property ParentElementID() As String
      Get
        Return _ParentElementID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ParentElementID = ""
         Else
           _ParentElementID = value
         End If
      End Set
    End Property
    Public Property ElementLevel() As String
      Get
        Return _ElementLevel
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ElementLevel = ""
         Else
           _ElementLevel = value
         End If
      End Set
    End Property
    Public Property Prefix() As String
      Get
        Return _Prefix
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _Prefix = ""
         Else
           _Prefix = value
         End If
      End Set
    End Property
    Public Property PAK_Elements1_Description() As String
      Get
        Return _PAK_Elements1_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_Elements1_Description = ""
         Else
           _PAK_Elements1_Description = value
         End If
      End Set
    End Property
    Public Property PAK_ResponsibleAgencies2_Description() As String
      Get
        Return _PAK_ResponsibleAgencies2_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_ResponsibleAgencies2_Description = ""
         Else
           _PAK_ResponsibleAgencies2_Description = value
         End If
      End Set
    End Property
    Public Readonly Property DisplayField() As String
      Get
        Return "" & _Description.ToString.PadRight(100, " ")
      End Get
    End Property
    Public Readonly Property PrimaryKey() As String
      Get
        Return _ElementID
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
    Public Class PKpakElements
      Private _ElementID As String = ""
      Public Property ElementID() As String
        Get
          Return _ElementID
        End Get
        Set(ByVal value As String)
          _ElementID = value
        End Set
      End Property
    End Class
    Public Shared Function pakElementsGetByID(ByVal ElementID As String) As SIS.PAK.pakElements
      Dim Results As SIS.PAK.pakElements = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakElementsSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ElementID", SqlDbType.NVarChar, ElementID.ToString.Length, ElementID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakElements(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function GetByParentElementID(ByVal ParentElementID As String, ByVal OrderBy As String) As List(Of SIS.PAK.pakElements)
      Dim Results As List(Of SIS.PAK.pakElements) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakElementsSelectByParentElementID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentElementID", SqlDbType.NVarChar, ParentElementID.ToString.Length, ParentElementID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakElements)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakElements(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakElementsGetByID(ByVal ElementID As String, ByVal Filter_ElementID As String, ByVal Filter_ResponsibleAgencyID As Int32) As SIS.PAK.pakElements
      Return pakElementsGetByID(ElementID)
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.PAK.pakElements) As SIS.PAK.pakElements
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakElementsInsert"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ElementID", SqlDbType.NVarChar, 9, Record.ElementID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 101, IIf(Record.Description = "", Convert.DBNull, Record.Description))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ResponsibleAgencyID", SqlDbType.Int, 11, IIf(Record.ResponsibleAgencyID = "", Convert.DBNull, Record.ResponsibleAgencyID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentElementID", SqlDbType.NVarChar, 9, IIf(Record.ParentElementID = "", Convert.DBNull, Record.ParentElementID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ElementLevel", SqlDbType.Int, 11, IIf(Record.ElementLevel = "", Convert.DBNull, Record.ElementLevel))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Prefix", SqlDbType.NVarChar, 1001, IIf(Record.Prefix = "", Convert.DBNull, Record.Prefix))
          Cmd.Parameters.Add("@Return_ElementID", SqlDbType.NVarChar, 9)
          Cmd.Parameters("@Return_ElementID").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.ElementID = Cmd.Parameters("@Return_ElementID").Value
        End Using
      End Using
      Return Record
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.PAK.pakElements) As SIS.PAK.pakElements
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakElementsUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_ElementID", SqlDbType.NVarChar, 9, Record.ElementID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ElementID", SqlDbType.NVarChar, 9, Record.ElementID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 101, IIf(Record.Description = "", Convert.DBNull, Record.Description))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ResponsibleAgencyID", SqlDbType.Int, 11, IIf(Record.ResponsibleAgencyID = "", Convert.DBNull, Record.ResponsibleAgencyID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentElementID", SqlDbType.NVarChar, 9, IIf(Record.ParentElementID = "", Convert.DBNull, Record.ParentElementID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ElementLevel", SqlDbType.Int, 11, IIf(Record.ElementLevel = "", Convert.DBNull, Record.ElementLevel))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Prefix", SqlDbType.NVarChar, 1001, IIf(Record.Prefix = "", Convert.DBNull, Record.Prefix))
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
      EDICommon.DBCommon.NewObj(Me, Reader)
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace
