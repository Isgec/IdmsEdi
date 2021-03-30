Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK
  <DataObject()> _
  Partial Public Class pakERPPOLine
    Private Shared _RecordCount As Integer
    Private _t_orno As String = ""
    Private _t_pono As Int32 = 0
    Private _t_item As String = ""
    Private _t_qoor As Double = 0
    Private _t_cuqp As String = ""
    Private _t_pric As Double = 0
    Private _t_oamt As Double = 0
    Private _t_cprj As String = ""
    Private _t_cspa As String = ""
    Public Property ItemDescription As String = ""
    Public Shared Function GetPOBOM(ByVal ERPPOLine As SIS.PAK.pakERPPOLine, ByVal PO As SIS.PAK.pakPO) As SIS.PAK.pakPOBOM
      Dim tmp As New SIS.PAK.pakPOBOM
      With tmp
        .SerialNo = PO.SerialNo
        .BOMNo = 0
        .ItemNo = ERPPOLine.t_pono
        .ItemCode = ERPPOLine.t_item.Trim
        .ItemDescription = ERPPOLine.ItemDescription
        .SupplierItemCode = ""
        .DivisionID = PO.DivisionID
        .ElementID = ERPPOLine.t_cspa
        If ERPPOLine.t_cuqp <> "" Then
          .UOMQuantity = SIS.PAK.pakUnits.pakUnitsGetByDescription(ERPPOLine.t_cuqp).UnitID
        End If
        .Quantity = ERPPOLine.t_qoor
        .UOMWeight = ""
        .WeightPerUnit = 0
        .DocumentNo = ""
        .ParentItemNo = ""
        .StatusID = pakItemStates.FreezedbyISGEC
        .ISGECRemarks = ""
        .SupplierRemarks = ""
        .ISGECQuantity = 0
        .ISGECWeightPerUnit = 0
        .SupplierQuantity = 0
        .SupplierWeightPerUnit = 0
        .Accepted = True
        .AcceptedBy = ""
        .AcceptedOn = ""
        .Freezed = True
        .FreezedBy = ""
        .FreezedOn = ""
        .Changed = False
        .CreatedBySupplier = False
        .ChangedBySupplier = False
        .AcceptedBySupplier = True
        .FreezedBySupplier = False
        .Active = True
        .Root = True
        .Middle = False
        .Bottom = False
        .Free = False
        .ItemLevel = 0
        .Prefix = ""
      End With
      Return tmp
    End Function

    Public Property t_orno() As String
      Get
        Return _t_orno
      End Get
      Set(ByVal value As String)
        _t_orno = value
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
    Public Property t_item() As String
      Get
        Return _t_item
      End Get
      Set(ByVal value As String)
        _t_item = value
      End Set
    End Property
    Public Property t_qoor() As Double
      Get
        Return _t_qoor
      End Get
      Set(ByVal value As Double)
        _t_qoor = value
      End Set
    End Property
    Public Property t_cuqp() As String
      Get
        Return _t_cuqp
      End Get
      Set(ByVal value As String)
        _t_cuqp = value
      End Set
    End Property
    Public Property t_pric() As Double
      Get
        Return _t_pric
      End Get
      Set(ByVal value As Double)
        _t_pric = value
      End Set
    End Property
    Public Property t_oamt() As Double
      Get
        Return _t_oamt
      End Get
      Set(ByVal value As Double)
        _t_oamt = value
      End Set
    End Property
    Public Property t_cprj() As String
      Get
        Return _t_cprj
      End Get
      Set(ByVal value As String)
        _t_cprj = value
      End Set
    End Property
    Public Property t_cspa() As String
      Get
        Return _t_cspa
      End Get
      Set(ByVal value As String)
        _t_cspa = value
      End Set
    End Property
    Public Readonly Property DisplayField() As String
      Get
        Return ""
      End Get
    End Property
    Public Readonly Property PrimaryKey() As String
      Get
        Return _t_orno & "|" & _t_pono
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
    Public Class PKpakERPPOLine
      Private _t_orno As String = ""
      Private _t_pono As Int32 = 0
      Public Property t_orno() As String
        Get
          Return _t_orno
        End Get
        Set(ByVal value As String)
          _t_orno = value
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
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function pakERPPOLineGetNewRecord() As SIS.PAK.pakERPPOLine
      Return New SIS.PAK.pakERPPOLine()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function pakERPPOLineGetByID(ByVal t_orno As String, ByVal t_pono As Int32) As SIS.PAK.pakERPPOLine
      Dim Results As SIS.PAK.pakERPPOLine = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakERPPOLineSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_orno", SqlDbType.VarChar, t_orno.ToString.Length, t_orno)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_pono", SqlDbType.Int, t_pono.ToString.Length, t_pono)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakERPPOLine(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakERPPOLineSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String) As List(Of SIS.PAK.pakERPPOLine)
      Dim Results As List(Of SIS.PAK.pakERPPOLine) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "sppakERPPOLineSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "sppakERPPOLineSelectListFilteres"
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakERPPOLine)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakERPPOLine(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function pakERPPOLineSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String) As Integer
      Return _RecordCount
    End Function
      'Select By ID One Record Filtered Overloaded GetByID
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
