Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK
  <DataObject()> _
  Partial Public Class pakPOBOM
    Private Shared _RecordCount As Integer
    Private _BOMNo As Int32 = 0
    Private _ItemNo As Int32 = 0
    Private _ItemDescription As String = ""
    Private _ElementID As String = ""
    Private _StatusID As String = ""
    Private _ISGECRemarks As String = ""
    Private _SupplierRemarks As String = ""
    Private _AcceptedBySupplier As Boolean = False
    Private _FreezedBySupplier As Boolean = False
    Private _Active As Boolean = False
    Private _Changed As Boolean = False
    Private _CreatedBySupplier As Boolean = False
    Private _ChangedBySupplier As Boolean = False
    Private _Root As Boolean = False
    Private _ItemLevel As Int32 = 0
    Private _Prefix As String = ""
    Private _SerialNo As Int32 = 0
    Private _Middle As Boolean = False
    Private _Bottom As Boolean = False
    Private _Free As Boolean = False
    Private _FreezedOn As String = ""
    Private _WeightPerUnit As Decimal = 0
    Private _UOMWeight As String = ""
    Private _ParentItemNo As String = ""
    Private _DocumentNo As String = ""
    Private _Quantity As Decimal = 0
    Private _SupplierItemCode As String = ""
    Private _ItemCode As String = ""
    Private _UOMQuantity As String = ""
    Private _DivisionID As String = ""
    Private _AcceptedOn As String = ""
    Private _AcceptedBy As String = ""
    Private _FreezedBy As String = ""
    Private _Freezed As Boolean = False
    Private _Accepted As Boolean = False
    Private _ISGECWeightPerUnit As Decimal = 0
    Private _ISGECQuantity As Decimal = 0
    Private _SupplierWeightPerUnit As Decimal = 0
    Private _SupplierQuantity As Decimal = 0
    Private _aspnet_Users1_UserFullName As String = ""
    Private _aspnet_Users2_UserFullName As String = ""
    Private _PAK_Divisions3_Description As String = ""
    Private _PAK_Documents4_cmba As String = ""
    Private _PAK_Elements5_Description As String = ""
    Private _PAK_PO6_PODescription As String = ""
    Private _PAK_POBOMStatus7_Description As String = ""
    Private _PAK_Units8_Description As String = ""
    Private _PAK_Units9_Description As String = ""
    Public Property QualityClearedQty As Decimal = 0
    Public Property BOMNo() As Int32
      Get
        Return _BOMNo
      End Get
      Set(ByVal value As Int32)
        _BOMNo = value
      End Set
    End Property
    Public Property ItemNo() As Int32
      Get
        Return _ItemNo
      End Get
      Set(ByVal value As Int32)
        _ItemNo = value
      End Set
    End Property
    Public Property ItemDescription() As String
      Get
        Return _ItemDescription
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ItemDescription = ""
         Else
           _ItemDescription = value
         End If
      End Set
    End Property
    Public Property ElementID() As String
      Get
        Return _ElementID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ElementID = ""
         Else
           _ElementID = value
         End If
      End Set
    End Property
    Public Property StatusID() As String
      Get
        Return _StatusID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _StatusID = ""
         Else
           _StatusID = value
         End If
      End Set
    End Property
    Public Property ISGECRemarks() As String
      Get
        Return _ISGECRemarks
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ISGECRemarks = ""
         Else
           _ISGECRemarks = value
         End If
      End Set
    End Property
    Public Property SupplierRemarks() As String
      Get
        Return _SupplierRemarks
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _SupplierRemarks = ""
         Else
           _SupplierRemarks = value
         End If
      End Set
    End Property
    Public Property AcceptedBySupplier() As Boolean
      Get
        Return _AcceptedBySupplier
      End Get
      Set(ByVal value As Boolean)
        _AcceptedBySupplier = value
      End Set
    End Property
    Public Property FreezedBySupplier() As Boolean
      Get
        Return _FreezedBySupplier
      End Get
      Set(ByVal value As Boolean)
        _FreezedBySupplier = value
      End Set
    End Property
    Public Property Active() As Boolean
      Get
        Return _Active
      End Get
      Set(ByVal value As Boolean)
        _Active = value
      End Set
    End Property
    Public Property Changed() As Boolean
      Get
        Return _Changed
      End Get
      Set(ByVal value As Boolean)
        _Changed = value
      End Set
    End Property
    Public Property CreatedBySupplier() As Boolean
      Get
        Return _CreatedBySupplier
      End Get
      Set(ByVal value As Boolean)
        _CreatedBySupplier = value
      End Set
    End Property
    Public Property ChangedBySupplier() As Boolean
      Get
        Return _ChangedBySupplier
      End Get
      Set(ByVal value As Boolean)
        _ChangedBySupplier = value
      End Set
    End Property
    Public Property Root() As Boolean
      Get
        Return _Root
      End Get
      Set(ByVal value As Boolean)
        _Root = value
      End Set
    End Property
    Public Property ItemLevel() As Int32
      Get
        Return _ItemLevel
      End Get
      Set(ByVal value As Int32)
        _ItemLevel = value
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
    Public Property SerialNo() As Int32
      Get
        Return _SerialNo
      End Get
      Set(ByVal value As Int32)
        _SerialNo = value
      End Set
    End Property
    Public Property Middle() As Boolean
      Get
        Return _Middle
      End Get
      Set(ByVal value As Boolean)
        _Middle = value
      End Set
    End Property
    Public Property Bottom() As Boolean
      Get
        Return _Bottom
      End Get
      Set(ByVal value As Boolean)
        _Bottom = value
      End Set
    End Property
    Public Property Free() As Boolean
      Get
        Return _Free
      End Get
      Set(ByVal value As Boolean)
        _Free = value
      End Set
    End Property
    Public Property FreezedOn() As String
      Get
        If Not _FreezedOn = String.Empty Then
          Return Convert.ToDateTime(_FreezedOn).ToString("dd/MM/yyyy")
        End If
        Return _FreezedOn
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _FreezedOn = ""
         Else
           _FreezedOn = value
         End If
      End Set
    End Property
    Public Property WeightPerUnit() As Decimal
      Get
        If _WeightPerUnit <= 0 Then Return 0
        Return _WeightPerUnit
      End Get
      Set(ByVal value As Decimal)
        _WeightPerUnit = value
      End Set
    End Property
    Public Property UOMWeight() As String
      Get
        Return _UOMWeight
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _UOMWeight = ""
         Else
           _UOMWeight = value
         End If
      End Set
    End Property
    Public Property ParentItemNo() As String
      Get
        Return _ParentItemNo
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ParentItemNo = ""
         Else
           _ParentItemNo = value
         End If
      End Set
    End Property
    Public Property DocumentNo() As String
      Get
        Return _DocumentNo
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _DocumentNo = ""
         Else
           _DocumentNo = value
         End If
      End Set
    End Property
    Public Property Quantity() As Decimal
      Get
        Return _Quantity
      End Get
      Set(ByVal value As Decimal)
        _Quantity = value
      End Set
    End Property
    Public Property SupplierItemCode() As String
      Get
        Return _SupplierItemCode
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _SupplierItemCode = ""
         Else
           _SupplierItemCode = value
         End If
      End Set
    End Property
    Public Property ItemCode() As String
      Get
        Return _ItemCode
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ItemCode = ""
         Else
           _ItemCode = value
         End If
      End Set
    End Property
    Public Property UOMQuantity() As String
      Get
        Return _UOMQuantity
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _UOMQuantity = ""
         Else
           _UOMQuantity = value
         End If
      End Set
    End Property
    Public Property DivisionID() As String
      Get
        Return _DivisionID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _DivisionID = ""
         Else
           _DivisionID = value
         End If
      End Set
    End Property
    Public Property AcceptedOn() As String
      Get
        If Not _AcceptedOn = String.Empty Then
          Return Convert.ToDateTime(_AcceptedOn).ToString("dd/MM/yyyy")
        End If
        Return _AcceptedOn
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _AcceptedOn = ""
         Else
           _AcceptedOn = value
         End If
      End Set
    End Property
    Public Property AcceptedBy() As String
      Get
        Return _AcceptedBy
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _AcceptedBy = ""
         Else
           _AcceptedBy = value
         End If
      End Set
    End Property
    Public Property FreezedBy() As String
      Get
        Return _FreezedBy
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _FreezedBy = ""
         Else
           _FreezedBy = value
         End If
      End Set
    End Property
    Public Property Freezed() As Boolean
      Get
        Return _Freezed
      End Get
      Set(ByVal value As Boolean)
        _Freezed = value
      End Set
    End Property
    Public Property Accepted() As Boolean
      Get
        Return _Accepted
      End Get
      Set(ByVal value As Boolean)
        _Accepted = value
      End Set
    End Property
    Public Property ISGECWeightPerUnit() As Decimal
      Get
        Return _ISGECWeightPerUnit
      End Get
      Set(ByVal value As Decimal)
        _ISGECWeightPerUnit = value
      End Set
    End Property
    Public Property ISGECQuantity() As Decimal
      Get
        Return _ISGECQuantity
      End Get
      Set(ByVal value As Decimal)
        _ISGECQuantity = value
      End Set
    End Property
    Public Property SupplierWeightPerUnit() As Decimal
      Get
        Return _SupplierWeightPerUnit
      End Get
      Set(ByVal value As Decimal)
        _SupplierWeightPerUnit = value
      End Set
    End Property
    Public Property SupplierQuantity() As Decimal
      Get
        Return _SupplierQuantity
      End Get
      Set(ByVal value As Decimal)
        _SupplierQuantity = value
      End Set
    End Property
    Public Property aspnet_Users1_UserFullName() As String
      Get
        Return _aspnet_Users1_UserFullName
      End Get
      Set(ByVal value As String)
        _aspnet_Users1_UserFullName = value
      End Set
    End Property
    Public Property aspnet_Users2_UserFullName() As String
      Get
        Return _aspnet_Users2_UserFullName
      End Get
      Set(ByVal value As String)
        _aspnet_Users2_UserFullName = value
      End Set
    End Property
    Public Property PAK_Divisions3_Description() As String
      Get
        Return _PAK_Divisions3_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_Divisions3_Description = ""
         Else
           _PAK_Divisions3_Description = value
         End If
      End Set
    End Property
    Public Property PAK_Documents4_cmba() As String
      Get
        Return _PAK_Documents4_cmba
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_Documents4_cmba = ""
         Else
           _PAK_Documents4_cmba = value
         End If
      End Set
    End Property
    Public Property PAK_Elements5_Description() As String
      Get
        Return _PAK_Elements5_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_Elements5_Description = ""
         Else
           _PAK_Elements5_Description = value
         End If
      End Set
    End Property
    Public Property PAK_PO6_PODescription() As String
      Get
        Return _PAK_PO6_PODescription
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_PO6_PODescription = ""
         Else
           _PAK_PO6_PODescription = value
         End If
      End Set
    End Property
    Public Property PAK_POBOMStatus7_Description() As String
      Get
        Return _PAK_POBOMStatus7_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_POBOMStatus7_Description = ""
         Else
           _PAK_POBOMStatus7_Description = value
         End If
      End Set
    End Property
    Public Property PAK_Units8_Description() As String
      Get
        Return _PAK_Units8_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_Units8_Description = ""
         Else
           _PAK_Units8_Description = value
         End If
      End Set
    End Property
    Public Property PAK_Units9_Description() As String
      Get
        Return _PAK_Units9_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_Units9_Description = ""
         Else
           _PAK_Units9_Description = value
         End If
      End Set
    End Property
    Public Readonly Property DisplayField() As String
      Get
        Return "" & _ItemDescription.ToString.PadRight(100, " ")
      End Get
    End Property
    Public Readonly Property PrimaryKey() As String
      Get
        Return _SerialNo & "|" & _BOMNo
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
    Public Class PKpakPOBOM
      Private _SerialNo As Int32 = 0
      Private _BOMNo As Int32 = 0
      Public Property SerialNo() As Int32
        Get
          Return _SerialNo
        End Get
        Set(ByVal value As Int32)
          _SerialNo = value
        End Set
      End Property
      Public Property BOMNo() As Int32
        Get
          Return _BOMNo
        End Get
        Set(ByVal value As Int32)
          _BOMNo = value
        End Set
      End Property
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakPOBOMGetByID(ByVal SerialNo As Int32, ByVal BOMNo As Int32) As SIS.PAK.pakPOBOM
      Dim Results As SIS.PAK.pakPOBOM = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOBOMSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, SerialNo.ToString.Length, SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@BOMNo", SqlDbType.Int, BOMNo.ToString.Length, BOMNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakPOBOM(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakPOBOMSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal SerialNo As Int32) As List(Of SIS.PAK.pakPOBOM)
      Dim Results As List(Of SIS.PAK.pakPOBOM) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "sppakPOBOMSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "sppakPOBOMSelectListFilteres"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_SerialNo", SqlDbType.Int, 10, IIf(SerialNo = Nothing, 0, SerialNo))
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakPOBOM)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakPOBOM(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function pakPOBOMSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String, ByVal SerialNo As Int32) As Integer
      Return _RecordCount
    End Function
    'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakPOBOMGetByID(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal Filter_SerialNo As Int32) As SIS.PAK.pakPOBOM
      Return pakPOBOMGetByID(SerialNo, BOMNo)
    End Function
    <DataObjectMethod(DataObjectMethodType.Insert, True)>
    Public Shared Function pakPOBOMInsert(ByVal Record As SIS.PAK.pakPOBOM) As SIS.PAK.pakPOBOM
      Dim _Rec As SIS.PAK.pakPOBOM = New SIS.PAK.pakPOBOM
      With _Rec
        .ItemNo = Record.ItemNo
        .ItemDescription = Record.ItemDescription
        .ElementID = Record.ElementID
        .StatusID = Record.StatusID
        .ISGECRemarks = Record.ISGECRemarks
        .SupplierRemarks = Record.SupplierRemarks
        .AcceptedBySupplier = Record.AcceptedBySupplier
        .FreezedBySupplier = Record.FreezedBySupplier
        .Active = Record.Active
        .Changed = Record.Changed
        .CreatedBySupplier = Record.CreatedBySupplier
        .ChangedBySupplier = Record.ChangedBySupplier
        .Root = Record.Root
        .ItemLevel = Record.ItemLevel
        .Prefix = Record.Prefix
        .SerialNo = Record.SerialNo
        .Middle = Record.Middle
        .Bottom = Record.Bottom
        .Free = Record.Free
        .FreezedOn = Record.FreezedOn
        .WeightPerUnit = Record.WeightPerUnit
        .UOMWeight = Record.UOMWeight
        .ParentItemNo = Record.ParentItemNo
        .DocumentNo = Record.DocumentNo
        .Quantity = Record.Quantity
        .SupplierItemCode = Record.SupplierItemCode
        .ItemCode = Record.ItemCode
        .UOMQuantity = Record.UOMQuantity
        .DivisionID = Record.DivisionID
        .AcceptedOn = Record.AcceptedOn
        .AcceptedBy = Record.AcceptedBy
        .FreezedBy = Record.FreezedBy
        .Freezed = Record.Freezed
        .Accepted = Record.Accepted
        .ISGECWeightPerUnit = Record.ISGECWeightPerUnit
        .ISGECQuantity = Record.ISGECQuantity
        .SupplierWeightPerUnit = Record.SupplierWeightPerUnit
        .SupplierQuantity = Record.SupplierQuantity
        .QualityClearedQty = Record.QualityClearedQty
      End With
      Return SIS.PAK.pakPOBOM.InsertData(_Rec)
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.PAK.pakPOBOM) As SIS.PAK.pakPOBOM
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOBOMInsert"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemNo", SqlDbType.Int, 11, Record.ItemNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemDescription", SqlDbType.NVarChar, 101, IIf(Record.ItemDescription = "", Convert.DBNull, Record.ItemDescription))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ElementID", SqlDbType.NVarChar, 9, IIf(Record.ElementID = "", Convert.DBNull, Record.ElementID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StatusID", SqlDbType.Int, 11, IIf(Record.StatusID = "", Convert.DBNull, Record.StatusID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECRemarks", SqlDbType.NVarChar, 501, IIf(Record.ISGECRemarks = "", Convert.DBNull, Record.ISGECRemarks))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierRemarks", SqlDbType.NVarChar, 501, IIf(Record.SupplierRemarks = "", Convert.DBNull, Record.SupplierRemarks))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedBySupplier", SqlDbType.Bit, 3, Record.AcceptedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedBySupplier", SqlDbType.Bit, 3, Record.FreezedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Active", SqlDbType.Bit, 3, Record.Active)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Changed", SqlDbType.Bit, 3, Record.Changed)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedBySupplier", SqlDbType.Bit, 3, Record.CreatedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ChangedBySupplier", SqlDbType.Bit, 3, Record.ChangedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Root", SqlDbType.Bit, 3, Record.Root)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemLevel", SqlDbType.Int, 11, Record.ItemLevel)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Prefix", SqlDbType.NVarChar, 1001, IIf(Record.Prefix = "", Convert.DBNull, Record.Prefix))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, 11, Record.SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Middle", SqlDbType.Bit, 3, Record.Middle)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Bottom", SqlDbType.Bit, 3, Record.Bottom)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Free", SqlDbType.Bit, 3, Record.Free)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedOn", SqlDbType.DateTime, 21, IIf(Record.FreezedOn = "", Convert.DBNull, Record.FreezedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@WeightPerUnit", SqlDbType.Decimal, 21, Record.WeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UOMWeight", SqlDbType.Int, 11, IIf(Record.UOMWeight = "", Convert.DBNull, Record.UOMWeight))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentItemNo", SqlDbType.Int, 11, IIf(Record.ParentItemNo = "", Convert.DBNull, Record.ParentItemNo))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentNo", SqlDbType.Int, 11, IIf(Record.DocumentNo = "", Convert.DBNull, Record.DocumentNo))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Quantity", SqlDbType.Decimal, 21, Record.Quantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierItemCode", SqlDbType.NVarChar, 51, IIf(Record.SupplierItemCode = "", Convert.DBNull, Record.SupplierItemCode))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemCode", SqlDbType.NVarChar, 51, IIf(Record.ItemCode = "", Convert.DBNull, Record.ItemCode))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UOMQuantity", SqlDbType.Int, 11, IIf(Record.UOMQuantity = "", Convert.DBNull, Record.UOMQuantity))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DivisionID", SqlDbType.Int, 11, IIf(Record.DivisionID = "", Convert.DBNull, Record.DivisionID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedOn", SqlDbType.DateTime, 21, IIf(Record.AcceptedOn = "", Convert.DBNull, Record.AcceptedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedBy", SqlDbType.NVarChar, 9, IIf(Record.AcceptedBy = "", Convert.DBNull, Record.AcceptedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedBy", SqlDbType.NVarChar, 9, IIf(Record.FreezedBy = "", Convert.DBNull, Record.FreezedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Freezed", SqlDbType.Bit, 3, Record.Freezed)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Accepted", SqlDbType.Bit, 3, Record.Accepted)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECWeightPerUnit", SqlDbType.Decimal, 21, Record.ISGECWeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECQuantity", SqlDbType.Decimal, 21, Record.ISGECQuantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierWeightPerUnit", SqlDbType.Decimal, 21, Record.SupplierWeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierQuantity", SqlDbType.Decimal, 21, Record.SupplierQuantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QualityClearedQty", SqlDbType.Decimal, 21, Record.QualityClearedQty)
          Cmd.Parameters.Add("@Return_SerialNo", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_SerialNo").Direction = ParameterDirection.Output
          Cmd.Parameters.Add("@Return_BOMNo", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_BOMNo").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.SerialNo = Cmd.Parameters("@Return_SerialNo").Value
          Record.BOMNo = Cmd.Parameters("@Return_BOMNo").Value
        End Using
      End Using
      Return Record
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)>
    Public Shared Function pakPOBOMUpdate(ByVal Record As SIS.PAK.pakPOBOM) As SIS.PAK.pakPOBOM
      Dim _Rec As SIS.PAK.pakPOBOM = SIS.PAK.pakPOBOM.pakPOBOMGetByID(Record.SerialNo, Record.BOMNo)
      With _Rec
        .ItemNo = Record.ItemNo
        .ItemDescription = Record.ItemDescription
        .ElementID = Record.ElementID
        .StatusID = Record.StatusID
        .ISGECRemarks = Record.ISGECRemarks
        .SupplierRemarks = Record.SupplierRemarks
        .AcceptedBySupplier = Record.AcceptedBySupplier
        .FreezedBySupplier = Record.FreezedBySupplier
        .Active = Record.Active
        .Changed = Record.Changed
        .CreatedBySupplier = Record.CreatedBySupplier
        .ChangedBySupplier = Record.ChangedBySupplier
        .Root = Record.Root
        .ItemLevel = Record.ItemLevel
        .Prefix = Record.Prefix
        .Middle = Record.Middle
        .Bottom = Record.Bottom
        .Free = Record.Free
        .FreezedOn = Record.FreezedOn
        .WeightPerUnit = Record.WeightPerUnit
        .UOMWeight = Record.UOMWeight
        .ParentItemNo = Record.ParentItemNo
        .DocumentNo = Record.DocumentNo
        .Quantity = Record.Quantity
        .SupplierItemCode = Record.SupplierItemCode
        .ItemCode = Record.ItemCode
        .UOMQuantity = Record.UOMQuantity
        .DivisionID = Record.DivisionID
        .AcceptedOn = Record.AcceptedOn
        .AcceptedBy = Record.AcceptedBy
        .FreezedBy = Record.FreezedBy
        .Freezed = Record.Freezed
        .Accepted = Record.Accepted
        .ISGECWeightPerUnit = Record.ISGECWeightPerUnit
        .ISGECQuantity = Record.ISGECQuantity
        .SupplierWeightPerUnit = Record.SupplierWeightPerUnit
        .SupplierQuantity = Record.SupplierQuantity
        .QualityClearedQty = Record.QualityClearedQty
      End With
      Return SIS.PAK.pakPOBOM.UpdateData(_Rec)
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.PAK.pakPOBOM) As SIS.PAK.pakPOBOM
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOBOMUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_BOMNo", SqlDbType.Int, 11, Record.BOMNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_SerialNo", SqlDbType.Int, 11, Record.SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemNo", SqlDbType.Int, 11, Record.ItemNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemDescription", SqlDbType.NVarChar, 101, IIf(Record.ItemDescription = "", Convert.DBNull, Record.ItemDescription))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ElementID", SqlDbType.NVarChar, 9, IIf(Record.ElementID = "", Convert.DBNull, Record.ElementID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StatusID", SqlDbType.Int, 11, IIf(Record.StatusID = "", Convert.DBNull, Record.StatusID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECRemarks", SqlDbType.NVarChar, 501, IIf(Record.ISGECRemarks = "", Convert.DBNull, Record.ISGECRemarks))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierRemarks", SqlDbType.NVarChar, 501, IIf(Record.SupplierRemarks = "", Convert.DBNull, Record.SupplierRemarks))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedBySupplier", SqlDbType.Bit, 3, Record.AcceptedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedBySupplier", SqlDbType.Bit, 3, Record.FreezedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Active", SqlDbType.Bit, 3, Record.Active)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Changed", SqlDbType.Bit, 3, Record.Changed)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedBySupplier", SqlDbType.Bit, 3, Record.CreatedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ChangedBySupplier", SqlDbType.Bit, 3, Record.ChangedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Root", SqlDbType.Bit, 3, Record.Root)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemLevel", SqlDbType.Int, 11, Record.ItemLevel)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Prefix", SqlDbType.NVarChar, 1001, IIf(Record.Prefix = "", Convert.DBNull, Record.Prefix))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, 11, Record.SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Middle", SqlDbType.Bit, 3, Record.Middle)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Bottom", SqlDbType.Bit, 3, Record.Bottom)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Free", SqlDbType.Bit, 3, Record.Free)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedOn", SqlDbType.DateTime, 21, IIf(Record.FreezedOn = "", Convert.DBNull, Record.FreezedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@WeightPerUnit", SqlDbType.Decimal, 21, Record.WeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UOMWeight", SqlDbType.Int, 11, IIf(Record.UOMWeight = "", Convert.DBNull, Record.UOMWeight))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentItemNo", SqlDbType.Int, 11, IIf(Record.ParentItemNo = "", Convert.DBNull, Record.ParentItemNo))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentNo", SqlDbType.Int, 11, IIf(Record.DocumentNo = "", Convert.DBNull, Record.DocumentNo))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Quantity", SqlDbType.Decimal, 21, Record.Quantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierItemCode", SqlDbType.NVarChar, 51, IIf(Record.SupplierItemCode = "", Convert.DBNull, Record.SupplierItemCode))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemCode", SqlDbType.NVarChar, 51, IIf(Record.ItemCode = "", Convert.DBNull, Record.ItemCode))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UOMQuantity", SqlDbType.Int, 11, IIf(Record.UOMQuantity = "", Convert.DBNull, Record.UOMQuantity))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DivisionID", SqlDbType.Int, 11, IIf(Record.DivisionID = "", Convert.DBNull, Record.DivisionID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedOn", SqlDbType.DateTime, 21, IIf(Record.AcceptedOn = "", Convert.DBNull, Record.AcceptedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedBy", SqlDbType.NVarChar, 9, IIf(Record.AcceptedBy = "", Convert.DBNull, Record.AcceptedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedBy", SqlDbType.NVarChar, 9, IIf(Record.FreezedBy = "", Convert.DBNull, Record.FreezedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Freezed", SqlDbType.Bit, 3, Record.Freezed)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Accepted", SqlDbType.Bit, 3, Record.Accepted)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECWeightPerUnit", SqlDbType.Decimal, 21, Record.ISGECWeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECQuantity", SqlDbType.Decimal, 21, Record.ISGECQuantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierWeightPerUnit", SqlDbType.Decimal, 21, Record.SupplierWeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierQuantity", SqlDbType.Decimal, 21, Record.SupplierQuantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QualityClearedQty", SqlDbType.Decimal, 21, Record.QualityClearedQty)
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
    Public Shared Function pakPOBOMDelete(ByVal Record As SIS.PAK.pakPOBOM) As Int32
      Dim _Result As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOBOMDelete"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_SerialNo", SqlDbType.Int, Record.SerialNo.ToString.Length, Record.SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_BOMNo", SqlDbType.Int, Record.BOMNo.ToString.Length, Record.BOMNo)
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
    Public Shared Function UZ_pakPOBOMSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal SerialNo As Int32) As List(Of SIS.PAK.pakPOBOM)
      Dim Results As List(Of SIS.PAK.pakPOBOM) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "sppak_LG_POBOMSelectListSearch"
            Cmd.CommandText = "sppakPOBOMSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "sppak_LG_POBOMSelectListFilteres"
            Cmd.CommandText = "sppakPOBOMSelectListFilteres"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_SerialNo", SqlDbType.Int, 10, IIf(SerialNo = Nothing, 0, SerialNo))
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakPOBOM)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakPOBOM(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function GetPOBItem(ByVal POBom As SIS.PAK.pakPOBOM) As SIS.PAK.pakPOBItems
      Dim tmpPOBItem As New SIS.PAK.pakPOBItems
      With tmpPOBItem
        .SerialNo = POBom.SerialNo
        .BOMNo = POBom.BOMNo
        .ItemNo = POBom.ItemNo
        .ItemCode = POBom.ItemCode
        .ItemDescription = POBom.ItemDescription
        .SupplierItemCode = POBom.SupplierItemCode
        .DivisionID = POBom.DivisionID
        .ElementID = POBom.ElementID
        .UOMQuantity = POBom.UOMQuantity
        .Quantity = POBom.Quantity
        .UOMWeight = POBom.UOMWeight
        .WeightPerUnit = POBom.WeightPerUnit
        .DocumentNo = POBom.DocumentNo
        .ParentItemNo = POBom.ParentItemNo
        .StatusID = POBom.StatusID
        .ISGECRemarks = POBom.ISGECRemarks
        .SupplierRemarks = POBom.SupplierRemarks
        .ISGECQuantity = POBom.ISGECQuantity
        .ISGECWeightPerUnit = POBom.ISGECWeightPerUnit
        .SupplierQuantity = POBom.SupplierQuantity
        .SupplierWeightPerUnit = POBom.SupplierWeightPerUnit
        .QualityClearedQty = POBom.QualityClearedQty
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
      Return tmpPOBItem
    End Function

  End Class
End Namespace
