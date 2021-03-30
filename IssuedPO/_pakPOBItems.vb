Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK
  <DataObject()> _
  Partial Public Class pakPOBItems
    Private Shared _RecordCount As Integer
    Private _ItemNo As Int32 = 0
    Private _ItemCode As String = ""
    Private _ItemDescription As String = ""
    Private _ElementID As String = ""
    Private _Quantity As Decimal = 0
    Private _WeightPerUnit As Decimal = 0
    Private _StatusID As String = ""
    Private _Bottom As Boolean = False
    Private _Free As Boolean = False
    Private _Middle As Boolean = False
    Private _Root As Boolean = False
    Private _ChangedBySupplier As Boolean = False
    Private _CreatedBySupplier As Boolean = False
    Private _Changed As Boolean = False
    Private _Active As Boolean = False
    Private _FreezedBySupplier As Boolean = False
    Private _AcceptedBySupplier As Boolean = False
    Private _QuantityDespatched As Decimal = 0
    Private _TotalWeightToDespatch As Decimal = 0
    Private _TotalWeightDespatched As Decimal = 0
    Private _TotalWeightReceived As Decimal = 0
    Private _QuantityReceived As Decimal = 0
    Private _Prefix As String = ""
    Private _ItemLevel As Int32 = 0
    Private _QuantityToPack As Decimal = 0
    Private _QuantityToDespatch As Decimal = 0
    Private _TotalWeightToPack As Decimal = 0
    Private _DocumentNo As String = ""
    Private _UOMWeight As String = ""
    Private _ParentItemNo As String = ""
    Private _SupplierRemarks As String = ""
    Private _ISGECRemarks As String = ""
    Private _BOMNo As Int32 = 0
    Private _SerialNo As Int32 = 0
    Private _SupplierItemCode As String = ""
    Private _UOMQuantity As String = ""
    Private _DivisionID As String = ""
    Private _AcceptedOn As String = ""
    Private _AcceptedBy As String = ""
    Private _Freezed As Boolean = False
    Private _FreezedOn As String = ""
    Private _FreezedBy As String = ""
    Private _ISGECWeightPerUnit As Decimal = 0
    Private _ISGECQuantity As Decimal = 0
    Private _SupplierQuantity As Decimal = 0
    Private _Accepted As Boolean = False
    Private _SupplierWeightPerUnit As Decimal = 0
    Public Property ItemReference As String = ""
    Public Property SubItem As String = ""
    Public Property SubItem2 As String = ""
    Public Property SubItem3 As String = ""
    Public Property SubItem4 As String = ""
    Public Property QuantityDespatchedToPort As Decimal = 0
    Public Property TotalWeightDespatchedToPort As Decimal = 0
    Public Property QuantityReceivedAtPort As Decimal = 0
    Public Property TotalWeightReceivedAtPort As Decimal = 0
    Public Property QuantityDespatchedfromPort As Decimal = 0
    Public Property TotalWeightDespatchedFromPort As Decimal = 0
    Public Property TotalWeight As Decimal = 0
    Public Property QualityClearedQty As Decimal = 0
    Public Property QualityClearedQtyStage As Decimal = 0
    Public Property DeletedInERP As Boolean = False
    Private _aspnet_Users1_UserFullName As String = ""
    Private _aspnet_Users2_UserFullName As String = ""
    Private _PAK_Divisions3_Description As String = ""
    Private _PAK_Documents4_cmba As String = ""
    Private _PAK_Elements5_Description As String = ""
    Private _PAK_PO6_PODescription As String = ""
    Private _PAK_POBItems7_ItemDescription As String = ""
    Private _PAK_POBOM8_ItemDescription As String = ""
    Private _PAK_POBOMStatus9_Description As String = ""
    Private _PAK_Units10_Description As String = ""
    Private _PAK_Units11_Description As String = ""
    Public Property ItemNo() As Int32
      Get
        Return _ItemNo
      End Get
      Set(ByVal value As Int32)
        _ItemNo = value
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
    Public Property Quantity() As Decimal
      Get
        Return _Quantity
      End Get
      Set(ByVal value As Decimal)
        _Quantity = value
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
    Public Property Middle() As Boolean
      Get
        Return _Middle
      End Get
      Set(ByVal value As Boolean)
        _Middle = value
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
    Public Property ChangedBySupplier() As Boolean
      Get
        Return _ChangedBySupplier
      End Get
      Set(ByVal value As Boolean)
        _ChangedBySupplier = value
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
    Public Property Changed() As Boolean
      Get
        Return _Changed
      End Get
      Set(ByVal value As Boolean)
        _Changed = value
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
    Public Property FreezedBySupplier() As Boolean
      Get
        Return _FreezedBySupplier
      End Get
      Set(ByVal value As Boolean)
        _FreezedBySupplier = value
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
    Public Property QuantityDespatched() As Decimal
      Get
        Return _QuantityDespatched
      End Get
      Set(ByVal value As Decimal)
        _QuantityDespatched = value
      End Set
    End Property
    Public Property TotalWeightToDespatch() As Decimal
      Get
        Return _TotalWeightToDespatch
      End Get
      Set(ByVal value As Decimal)
        _TotalWeightToDespatch = value
      End Set
    End Property
    Public Property TotalWeightDespatched() As Decimal
      Get
        Return _TotalWeightDespatched
      End Get
      Set(ByVal value As Decimal)
        _TotalWeightDespatched = value
      End Set
    End Property
    Public Property TotalWeightReceived() As Decimal
      Get
        Return _TotalWeightReceived
      End Get
      Set(ByVal value As Decimal)
        _TotalWeightReceived = value
      End Set
    End Property
    Public Property QuantityReceived() As Decimal
      Get
        Return _QuantityReceived
      End Get
      Set(ByVal value As Decimal)
        _QuantityReceived = value
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
    Public Property ItemLevel() As Int32
      Get
        Return _ItemLevel
      End Get
      Set(ByVal value As Int32)
        _ItemLevel = value
      End Set
    End Property
    Public Property QuantityToPack() As Decimal
      Get
        Return _QuantityToPack
      End Get
      Set(ByVal value As Decimal)
        _QuantityToPack = value
      End Set
    End Property
    Public Property QuantityToDespatch() As Decimal
      Get
        Return _QuantityToDespatch
      End Get
      Set(ByVal value As Decimal)
        _QuantityToDespatch = value
      End Set
    End Property
    Public Property TotalWeightToPack() As Decimal
      Get
        Return _TotalWeightToPack
      End Get
      Set(ByVal value As Decimal)
        _TotalWeightToPack = value
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
    Public Property BOMNo() As Int32
      Get
        Return _BOMNo
      End Get
      Set(ByVal value As Int32)
        _BOMNo = value
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
    Public Property Freezed() As Boolean
      Get
        Return _Freezed
      End Get
      Set(ByVal value As Boolean)
        _Freezed = value
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
    Public Property SupplierQuantity() As Decimal
      Get
        Return _SupplierQuantity
      End Get
      Set(ByVal value As Decimal)
        _SupplierQuantity = value
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
    Public Property SupplierWeightPerUnit() As Decimal
      Get
        Return _SupplierWeightPerUnit
      End Get
      Set(ByVal value As Decimal)
        _SupplierWeightPerUnit = value
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
    Public Property PAK_POBItems7_ItemDescription() As String
      Get
        Return _PAK_POBItems7_ItemDescription
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_POBItems7_ItemDescription = ""
         Else
           _PAK_POBItems7_ItemDescription = value
         End If
      End Set
    End Property
    Public Property PAK_POBOM8_ItemDescription() As String
      Get
        Return _PAK_POBOM8_ItemDescription
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_POBOM8_ItemDescription = ""
         Else
           _PAK_POBOM8_ItemDescription = value
         End If
      End Set
    End Property
    Public Property PAK_POBOMStatus9_Description() As String
      Get
        Return _PAK_POBOMStatus9_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_POBOMStatus9_Description = ""
         Else
           _PAK_POBOMStatus9_Description = value
         End If
      End Set
    End Property
    Public Property PAK_Units10_Description() As String
      Get
        Return _PAK_Units10_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_Units10_Description = ""
         Else
           _PAK_Units10_Description = value
         End If
      End Set
    End Property
    Public Property PAK_Units11_Description() As String
      Get
        Return _PAK_Units11_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PAK_Units11_Description = ""
         Else
           _PAK_Units11_Description = value
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
        Return _SerialNo & "|" & _BOMNo & "|" & _ItemNo
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
    Public Class PKpakPOBItems
      Private _SerialNo As Int32 = 0
      Private _BOMNo As Int32 = 0
      Private _ItemNo As Int32 = 0
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
      Public Property ItemNo() As Int32
        Get
          Return _ItemNo
        End Get
        Set(ByVal value As Int32)
          _ItemNo = value
        End Set
      End Property
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakPOBItemsGetNewRecord() As SIS.PAK.pakPOBItems
      Return New SIS.PAK.pakPOBItems()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakPOBItemsGetByID(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemNo As Int32) As SIS.PAK.pakPOBItems
      Dim Results As SIS.PAK.pakPOBItems = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOBItemsSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, SerialNo.ToString.Length, SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@BOMNo", SqlDbType.Int, BOMNo.ToString.Length, BOMNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemNo", SqlDbType.Int, ItemNo.ToString.Length, ItemNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakPOBItems(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakPOBItemsSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal BOMNo As Int32, ByVal SerialNo As Int32) As List(Of SIS.PAK.pakPOBItems)
      Dim Results As List(Of SIS.PAK.pakPOBItems) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "sppakPOBItemsSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "sppakPOBItemsSelectListFilteres"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_BOMNo", SqlDbType.Int, 10, IIf(BOMNo = Nothing, 0, BOMNo))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_SerialNo", SqlDbType.Int, 10, IIf(SerialNo = Nothing, 0, SerialNo))
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakPOBItems)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakPOBItems(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function pakPOBItemsSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String, ByVal BOMNo As Int32, ByVal SerialNo As Int32) As Integer
      Return _RecordCount
    End Function
    'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function pakPOBItemsGetByID(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemNo As Int32, ByVal Filter_BOMNo As Int32, ByVal Filter_SerialNo As Int32) As SIS.PAK.pakPOBItems
      Return pakPOBItemsGetByID(SerialNo, BOMNo, ItemNo)
    End Function
    <DataObjectMethod(DataObjectMethodType.Insert, True)>
    Public Shared Function pakPOBItemsInsert(ByVal Record As SIS.PAK.pakPOBItems) As SIS.PAK.pakPOBItems
      Dim _Rec As SIS.PAK.pakPOBItems = SIS.PAK.pakPOBItems.pakPOBItemsGetNewRecord()
      With _Rec
        .ItemNo = Record.ItemNo
        .ItemCode = Record.ItemCode
        .ItemDescription = Record.ItemDescription
        .ElementID = Record.ElementID
        .Quantity = Record.Quantity
        .WeightPerUnit = Record.WeightPerUnit
        .StatusID = Record.StatusID
        .Bottom = Record.Bottom
        .Free = Record.Free
        .Middle = Record.Middle
        .Root = Record.Root
        .ChangedBySupplier = Record.ChangedBySupplier
        .CreatedBySupplier = Record.CreatedBySupplier
        .Changed = Record.Changed
        .Active = Record.Active
        .FreezedBySupplier = Record.FreezedBySupplier
        .AcceptedBySupplier = Record.AcceptedBySupplier
        .QuantityDespatched = Record.QuantityDespatched
        .TotalWeightToDespatch = Record.TotalWeightToDespatch
        .TotalWeightDespatched = Record.TotalWeightDespatched
        .TotalWeightReceived = Record.TotalWeightReceived
        .QuantityReceived = Record.QuantityReceived
        .Prefix = Record.Prefix
        .ItemLevel = Record.ItemLevel
        .QuantityToPack = Record.QuantityToPack
        .QuantityToDespatch = Record.QuantityToDespatch
        .TotalWeightToPack = Record.TotalWeightToPack
        .DocumentNo = Record.DocumentNo
        .UOMWeight = Record.UOMWeight
        .ParentItemNo = Record.ParentItemNo
        .SupplierRemarks = Record.SupplierRemarks
        .ISGECRemarks = Record.ISGECRemarks
        .BOMNo = Record.BOMNo
        .SerialNo = Record.SerialNo
        .SupplierItemCode = Record.SupplierItemCode
        .UOMQuantity = Record.UOMQuantity
        .DivisionID = Record.DivisionID
        .AcceptedOn = Record.AcceptedOn
        .AcceptedBy = Record.AcceptedBy
        .Freezed = Record.Freezed
        .FreezedOn = Record.FreezedOn
        .FreezedBy = Record.FreezedBy
        .ISGECWeightPerUnit = Record.ISGECWeightPerUnit
        .ISGECQuantity = Record.ISGECQuantity
        .SupplierQuantity = Record.SupplierQuantity
        .Accepted = Record.Accepted
        .SupplierWeightPerUnit = Record.SupplierWeightPerUnit
        .QualityClearedQty = Record.QualityClearedQty
        .ItemReference = Record.ItemReference
        .SubItem = Record.SubItem
        .SubItem2 = Record.SubItem2
        .SubItem3 = Record.SubItem3
        .SubItem4 = Record.SubItem4
        .QualityClearedQtyStage = Record.QualityClearedQtyStage
      End With
      Return SIS.PAK.pakPOBItems.InsertData(_Rec)
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.PAK.pakPOBItems) As SIS.PAK.pakPOBItems
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOBItemsInsert"
          With Record
            .TotalWeight = SIS.PAK.pakPO.GetTotalWeight(.Quantity, .WeightPerUnit, .UOMQuantity, .UOMWeight)
          End With
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemNo", SqlDbType.Int, 11, Record.ItemNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemCode", SqlDbType.NVarChar, 51, IIf(Record.ItemCode = "", Convert.DBNull, Record.ItemCode))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemDescription", SqlDbType.NVarChar, 101, IIf(Record.ItemDescription = "", Convert.DBNull, Record.ItemDescription))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ElementID", SqlDbType.NVarChar, 9, IIf(Record.ElementID = "", Convert.DBNull, Record.ElementID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Quantity", SqlDbType.Decimal, 21, Record.Quantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@WeightPerUnit", SqlDbType.Decimal, 21, Record.WeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StatusID", SqlDbType.Int, 11, IIf(Record.StatusID = "", Convert.DBNull, Record.StatusID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Bottom", SqlDbType.Bit, 3, Record.Bottom)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Free", SqlDbType.Bit, 3, Record.Free)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Middle", SqlDbType.Bit, 3, Record.Middle)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Root", SqlDbType.Bit, 3, Record.Root)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ChangedBySupplier", SqlDbType.Bit, 3, Record.ChangedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedBySupplier", SqlDbType.Bit, 3, Record.CreatedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Changed", SqlDbType.Bit, 3, Record.Changed)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Active", SqlDbType.Bit, 3, Record.Active)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedBySupplier", SqlDbType.Bit, 3, Record.FreezedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedBySupplier", SqlDbType.Bit, 3, Record.AcceptedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityDespatched", SqlDbType.Decimal, 21, Record.QuantityDespatched)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightToDespatch", SqlDbType.Decimal, 21, Record.TotalWeightToDespatch)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightDespatched", SqlDbType.Decimal, 21, Record.TotalWeightDespatched)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightReceived", SqlDbType.Decimal, 21, Record.TotalWeightReceived)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityReceived", SqlDbType.Decimal, 21, Record.QuantityReceived)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Prefix", SqlDbType.NVarChar, 1001, IIf(Record.Prefix = "", Convert.DBNull, Record.Prefix))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemLevel", SqlDbType.Int, 11, Record.ItemLevel)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityToPack", SqlDbType.Decimal, 21, Record.QuantityToPack)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityToDespatch", SqlDbType.Decimal, 21, Record.QuantityToDespatch)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightToPack", SqlDbType.Decimal, 21, Record.TotalWeightToPack)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentNo", SqlDbType.Int, 11, IIf(Record.DocumentNo = "", Convert.DBNull, Record.DocumentNo))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UOMWeight", SqlDbType.Int, 11, IIf(Record.UOMWeight = "", Convert.DBNull, Record.UOMWeight))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentItemNo", SqlDbType.Int, 11, IIf(Record.ParentItemNo = "", Convert.DBNull, Record.ParentItemNo))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierRemarks", SqlDbType.NVarChar, 501, IIf(Record.SupplierRemarks = "", Convert.DBNull, Record.SupplierRemarks))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECRemarks", SqlDbType.NVarChar, 501, IIf(Record.ISGECRemarks = "", Convert.DBNull, Record.ISGECRemarks))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@BOMNo", SqlDbType.Int, 11, Record.BOMNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, 11, Record.SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierItemCode", SqlDbType.NVarChar, 51, IIf(Record.SupplierItemCode = "", Convert.DBNull, Record.SupplierItemCode))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UOMQuantity", SqlDbType.Int, 11, IIf(Record.UOMQuantity = "", Convert.DBNull, Record.UOMQuantity))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DivisionID", SqlDbType.Int, 11, IIf(Record.DivisionID = "", Convert.DBNull, Record.DivisionID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedOn", SqlDbType.DateTime, 21, IIf(Record.AcceptedOn = "", Convert.DBNull, Record.AcceptedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedBy", SqlDbType.NVarChar, 9, IIf(Record.AcceptedBy = "", Convert.DBNull, Record.AcceptedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Freezed", SqlDbType.Bit, 3, Record.Freezed)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedOn", SqlDbType.DateTime, 21, IIf(Record.FreezedOn = "", Convert.DBNull, Record.FreezedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedBy", SqlDbType.NVarChar, 9, IIf(Record.FreezedBy = "", Convert.DBNull, Record.FreezedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECWeightPerUnit", SqlDbType.Decimal, 21, Record.ISGECWeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECQuantity", SqlDbType.Decimal, 21, Record.ISGECQuantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierQuantity", SqlDbType.Decimal, 21, Record.SupplierQuantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Accepted", SqlDbType.Bit, 3, Record.Accepted)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierWeightPerUnit", SqlDbType.Decimal, 21, Record.SupplierWeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QualityClearedQty", SqlDbType.Decimal, 21, Record.QualityClearedQty)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemReference", SqlDbType.NVarChar, 201, IIf(Record.ItemReference = "", Convert.DBNull, Record.ItemReference))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubItem", SqlDbType.NVarChar, 10, IIf(Record.SubItem = "", Convert.DBNull, Record.SubItem))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubItem2", SqlDbType.NVarChar, 151, IIf(Record.SubItem2 = "", Convert.DBNull, Record.SubItem2))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubItem3", SqlDbType.NVarChar, 151, IIf(Record.SubItem3 = "", Convert.DBNull, Record.SubItem3))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubItem4", SqlDbType.NVarChar, 151, IIf(Record.SubItem4 = "", Convert.DBNull, Record.SubItem4))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityReceivedAtPort", SqlDbType.Decimal, 21, Record.QuantityReceivedAtPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightReceivedAtPort", SqlDbType.Decimal, 21, Record.TotalWeightReceivedAtPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityDespatchedfromPort", SqlDbType.Decimal, 21, Record.QuantityDespatchedfromPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightDespatchedFromPort", SqlDbType.Decimal, 21, Record.TotalWeightDespatchedFromPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityDespatchedToPort", SqlDbType.Decimal, 21, Record.QuantityDespatchedToPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightDespatchedToPort", SqlDbType.Decimal, 21, Record.TotalWeightDespatchedToPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QualityClearedQtyStage", SqlDbType.Decimal, 21, Record.QualityClearedQtyStage)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeight", SqlDbType.Decimal, 21, Record.TotalWeight)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DeletedInERP", SqlDbType.Bit, 3, Record.DeletedInERP)
          Cmd.Parameters.Add("@Return_SerialNo", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_SerialNo").Direction = ParameterDirection.Output
          Cmd.Parameters.Add("@Return_BOMNo", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_BOMNo").Direction = ParameterDirection.Output
          Cmd.Parameters.Add("@Return_ItemNo", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_ItemNo").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.SerialNo = Cmd.Parameters("@Return_SerialNo").Value
          Record.BOMNo = Cmd.Parameters("@Return_BOMNo").Value
          Record.ItemNo = Cmd.Parameters("@Return_ItemNo").Value
        End Using
      End Using
      Return Record
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)>
    Public Shared Function pakPOBItemsUpdate(ByVal Record As SIS.PAK.pakPOBItems) As SIS.PAK.pakPOBItems
      Dim _Rec As SIS.PAK.pakPOBItems = SIS.PAK.pakPOBItems.pakPOBItemsGetByID(Record.SerialNo, Record.BOMNo, Record.ItemNo)
      With _Rec
        .ItemCode = Record.ItemCode
        .ItemDescription = Record.ItemDescription
        .ElementID = Record.ElementID
        .Quantity = Record.Quantity
        .WeightPerUnit = Record.WeightPerUnit
        .StatusID = Record.StatusID
        .Bottom = Record.Bottom
        .Free = Record.Free
        .Middle = Record.Middle
        .Root = Record.Root
        .ChangedBySupplier = Record.ChangedBySupplier
        .CreatedBySupplier = Record.CreatedBySupplier
        .Changed = Record.Changed
        .Active = Record.Active
        .FreezedBySupplier = Record.FreezedBySupplier
        .AcceptedBySupplier = Record.AcceptedBySupplier
        .QuantityDespatched = Record.QuantityDespatched
        .TotalWeightToDespatch = Record.TotalWeightToDespatch
        .TotalWeightDespatched = Record.TotalWeightDespatched
        .TotalWeightReceived = Record.TotalWeightReceived
        .QuantityReceived = Record.QuantityReceived
        .Prefix = Record.Prefix
        .ItemLevel = Record.ItemLevel
        .QuantityToPack = Record.QuantityToPack
        .QuantityToDespatch = Record.QuantityToDespatch
        .TotalWeightToPack = Record.TotalWeightToPack
        .DocumentNo = Record.DocumentNo
        .UOMWeight = Record.UOMWeight
        .ParentItemNo = Record.ParentItemNo
        .SupplierRemarks = Record.SupplierRemarks
        .ISGECRemarks = Record.ISGECRemarks
        .SupplierItemCode = Record.SupplierItemCode
        .UOMQuantity = Record.UOMQuantity
        .DivisionID = Record.DivisionID
        .AcceptedOn = Record.AcceptedOn
        .AcceptedBy = Record.AcceptedBy
        .Freezed = Record.Freezed
        .FreezedOn = Record.FreezedOn
        .FreezedBy = Record.FreezedBy
        .ISGECWeightPerUnit = Record.ISGECWeightPerUnit
        .ISGECQuantity = Record.ISGECQuantity
        .SupplierQuantity = Record.SupplierQuantity
        .Accepted = Record.Accepted
        .SupplierWeightPerUnit = Record.SupplierWeightPerUnit
        .QualityClearedQty = Record.QualityClearedQty
        .ItemReference = Record.ItemReference
        .SubItem = Record.SubItem
        .SubItem2 = Record.SubItem2
        .SubItem3 = Record.SubItem3
        .SubItem4 = Record.SubItem4
      End With
      Return SIS.PAK.pakPOBItems.UpdateData(_Rec)
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.PAK.pakPOBItems) As SIS.PAK.pakPOBItems
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOBItemsUpdate"
          With Record
            .TotalWeight = SIS.PAK.pakPO.GetTotalWeight(.Quantity, .WeightPerUnit, .UOMQuantity, .UOMWeight)
          End With
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_ItemNo", SqlDbType.Int, 11, Record.ItemNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_BOMNo", SqlDbType.Int, 11, Record.BOMNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_SerialNo", SqlDbType.Int, 11, Record.SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemNo", SqlDbType.Int, 11, Record.ItemNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemCode", SqlDbType.NVarChar, 51, IIf(Record.ItemCode = "", Convert.DBNull, Record.ItemCode))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemDescription", SqlDbType.NVarChar, 101, IIf(Record.ItemDescription = "", Convert.DBNull, Record.ItemDescription))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ElementID", SqlDbType.NVarChar, 9, IIf(Record.ElementID = "", Convert.DBNull, Record.ElementID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Quantity", SqlDbType.Decimal, 21, Record.Quantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@WeightPerUnit", SqlDbType.Decimal, 21, Record.WeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StatusID", SqlDbType.Int, 11, IIf(Record.StatusID = "", Convert.DBNull, Record.StatusID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Bottom", SqlDbType.Bit, 3, Record.Bottom)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Free", SqlDbType.Bit, 3, Record.Free)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Middle", SqlDbType.Bit, 3, Record.Middle)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Root", SqlDbType.Bit, 3, Record.Root)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ChangedBySupplier", SqlDbType.Bit, 3, Record.ChangedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedBySupplier", SqlDbType.Bit, 3, Record.CreatedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Changed", SqlDbType.Bit, 3, Record.Changed)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Active", SqlDbType.Bit, 3, Record.Active)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedBySupplier", SqlDbType.Bit, 3, Record.FreezedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedBySupplier", SqlDbType.Bit, 3, Record.AcceptedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityDespatched", SqlDbType.Decimal, 21, Record.QuantityDespatched)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightToDespatch", SqlDbType.Decimal, 21, Record.TotalWeightToDespatch)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightDespatched", SqlDbType.Decimal, 21, Record.TotalWeightDespatched)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightReceived", SqlDbType.Decimal, 21, Record.TotalWeightReceived)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityReceived", SqlDbType.Decimal, 21, Record.QuantityReceived)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Prefix", SqlDbType.NVarChar, 1001, IIf(Record.Prefix = "", Convert.DBNull, Record.Prefix))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemLevel", SqlDbType.Int, 11, Record.ItemLevel)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityToPack", SqlDbType.Decimal, 21, Record.QuantityToPack)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityToDespatch", SqlDbType.Decimal, 21, Record.QuantityToDespatch)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightToPack", SqlDbType.Decimal, 21, Record.TotalWeightToPack)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentNo", SqlDbType.Int, 11, IIf(Record.DocumentNo = "", Convert.DBNull, Record.DocumentNo))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UOMWeight", SqlDbType.Int, 11, IIf(Record.UOMWeight = "", Convert.DBNull, Record.UOMWeight))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentItemNo", SqlDbType.Int, 11, IIf(Record.ParentItemNo = "", Convert.DBNull, Record.ParentItemNo))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierRemarks", SqlDbType.NVarChar, 501, IIf(Record.SupplierRemarks = "", Convert.DBNull, Record.SupplierRemarks))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECRemarks", SqlDbType.NVarChar, 501, IIf(Record.ISGECRemarks = "", Convert.DBNull, Record.ISGECRemarks))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@BOMNo", SqlDbType.Int, 11, Record.BOMNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, 11, Record.SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierItemCode", SqlDbType.NVarChar, 51, IIf(Record.SupplierItemCode = "", Convert.DBNull, Record.SupplierItemCode))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@UOMQuantity", SqlDbType.Int, 11, IIf(Record.UOMQuantity = "", Convert.DBNull, Record.UOMQuantity))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DivisionID", SqlDbType.Int, 11, IIf(Record.DivisionID = "", Convert.DBNull, Record.DivisionID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedOn", SqlDbType.DateTime, 21, IIf(Record.AcceptedOn = "", Convert.DBNull, Record.AcceptedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedBy", SqlDbType.NVarChar, 9, IIf(Record.AcceptedBy = "", Convert.DBNull, Record.AcceptedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Freezed", SqlDbType.Bit, 3, Record.Freezed)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedOn", SqlDbType.DateTime, 21, IIf(Record.FreezedOn = "", Convert.DBNull, Record.FreezedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@FreezedBy", SqlDbType.NVarChar, 9, IIf(Record.FreezedBy = "", Convert.DBNull, Record.FreezedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECWeightPerUnit", SqlDbType.Decimal, 21, Record.ISGECWeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECQuantity", SqlDbType.Decimal, 21, Record.ISGECQuantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierQuantity", SqlDbType.Decimal, 21, Record.SupplierQuantity)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Accepted", SqlDbType.Bit, 3, Record.Accepted)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierWeightPerUnit", SqlDbType.Decimal, 21, Record.SupplierWeightPerUnit)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QualityClearedQty", SqlDbType.Decimal, 21, Record.QualityClearedQty)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemReference", SqlDbType.NVarChar, 201, IIf(Record.ItemReference = "", Convert.DBNull, Record.ItemReference))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubItem", SqlDbType.NVarChar, 10, IIf(Record.SubItem = "", Convert.DBNull, Record.SubItem))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubItem2", SqlDbType.NVarChar, 151, IIf(Record.SubItem2 = "", Convert.DBNull, Record.SubItem2))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubItem3", SqlDbType.NVarChar, 151, IIf(Record.SubItem3 = "", Convert.DBNull, Record.SubItem3))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubItem4", SqlDbType.NVarChar, 151, IIf(Record.SubItem4 = "", Convert.DBNull, Record.SubItem4))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityReceivedAtPort", SqlDbType.Decimal, 21, Record.QuantityReceivedAtPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightReceivedAtPort", SqlDbType.Decimal, 21, Record.TotalWeightReceivedAtPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityDespatchedfromPort", SqlDbType.Decimal, 21, Record.QuantityDespatchedfromPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightDespatchedFromPort", SqlDbType.Decimal, 21, Record.TotalWeightDespatchedFromPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QuantityDespatchedToPort", SqlDbType.Decimal, 21, Record.QuantityDespatchedToPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeightDespatchedToPort", SqlDbType.Decimal, 21, Record.TotalWeightDespatchedToPort)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QualityClearedQtyStage", SqlDbType.Decimal, 21, Record.QualityClearedQtyStage)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TotalWeight", SqlDbType.Decimal, 21, Record.TotalWeight)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DeletedInERP", SqlDbType.Bit, 3, Record.DeletedInERP)
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
    Public Shared Function pakPOBItemsDelete(ByVal Record As SIS.PAK.pakPOBItems) As Int32
      '===================================
      'Check QC or Despatch, If Not found then only can initiate or delete
      If Record.QualityClearedQty > 0 Or Record.QualityClearedQtyStage > 0 Or Record.QuantityDespatched > 0 Or Record.QuantityDespatchedToPort > 0 Then
        Throw New Exception("There are QC or Despatch in Item [tree from here], CAN Not perform Delete. You may try to delete individual Item.")
      End If
      Dim docs As List(Of SIS.PAK.pakPOBIDocuments) = SIS.PAK.pakPOBIDocuments.UZ_pakPOBIDocumentsSelectList(0, 9999, "", False, "", Record.SerialNo, Record.BOMNo, Record.ItemNo)
      For Each doc As SIS.PAK.pakPOBIDocuments In docs
        SIS.PAK.pakPOBIDocuments.UZ_pakPOBIDocumentsDelete(doc)
      Next
      '===================================
      Dim _Result As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOBItemsDelete"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_SerialNo", SqlDbType.Int, Record.SerialNo.ToString.Length, Record.SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_BOMNo", SqlDbType.Int, Record.BOMNo.ToString.Length, Record.BOMNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_ItemNo", SqlDbType.Int, Record.ItemNo.ToString.Length, Record.ItemNo)
          Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Con.Open()
          Cmd.ExecuteNonQuery()
          _RecordCount = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Dim tmpP As SIS.PAK.pakPOBItems = SIS.PAK.pakPOBItems.pakPOBItemsGetByID(Record.SerialNo, Record.BOMNo, Record.ParentItemNo)
      If tmpP IsNot Nothing Then
        If Not tmpP.Root Then
          Dim tmpPs As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.GetByParentPOBItemNo(Record.SerialNo, Record.BOMNo, Record.ParentItemNo, "")
          If tmpPs.Count <= 0 Then
            tmpP.Middle = False
            tmpP.Bottom = True
            tmpP = UpdateData(tmpP)
          End If
        End If
      End If
      Return _RecordCount
    End Function
    Public Shared Function pakPOBItemsDeleteRecursive(Results As SIS.PAK.pakPOBItems) As SIS.PAK.pakPOBItems
      RecursiveDelete(Results)
      Return Results
    End Function
    Private Shared Sub RecursiveDelete(ByVal pItm As SIS.PAK.pakPOBItems)
      Dim Items As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.GetByParentPOBItemNo(pItm.SerialNo, pItm.BOMNo, pItm.ItemNo, "")
      If Items.Count > 0 Then
        For Each Itm As SIS.PAK.pakPOBItems In Items
          RecursiveDelete(Itm)
        Next
      End If
      SIS.PAK.pakPOBItems.pakPOBItemsDelete(pItm)
    End Sub
    Public Shared Function CanBeDeleted(ByVal pItm As SIS.PAK.pakPOBItems) As Boolean
      If pItm.QualityClearedQty > 0 Or pItm.QualityClearedQtyStage > 0 Or pItm.QuantityDespatched > 0 Or pItm.QuantityDespatchedToPort > 0 Then
        Return False
      End If
      Dim Items As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.GetByParentPOBItemNo(pItm.SerialNo, pItm.BOMNo, pItm.ItemNo, "")
      If Items.Count > 0 Then
        For Each Itm As SIS.PAK.pakPOBItems In Items
          If Not CanBeDeleted(Itm) Then
            Return False
          End If
        Next
      End If
      Return True
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
