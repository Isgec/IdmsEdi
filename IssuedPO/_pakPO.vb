Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel

Namespace SIS.PAK
  <DataObject()> _
  Partial Public Class pakPO
    Private Shared _RecordCount As Integer
    Private _SerialNo As Int32 = 0
    Private _POConsignee As String = ""
    Private _POOtherDetails As String = ""
    Private _IssueReasonID As String = ""
    Private _PONumber As String = ""
    Private _PORevision As String = ""
    Private _PODate As String = ""
    Private _PODescription As String = ""
    Private _POTypeID As String = ""
    Private _SupplierID As String = ""
    Private _ProjectID As String = ""
    Private _BuyerID As String = ""
    Private _POStatusID As String = ""
    Private _ISGECRemarks As String = ""
    Private _SupplierRemarks As String = ""
    Private _IssuedBy As String = ""
    Private _IssuedOn As String = ""
    Private _ClosedBy As String = ""
    Private _ClosedOn As String = ""
    Private _DivisionID As String = ""
    Private _aspnet_Users1_UserFullName As String = ""
    Private _aspnet_Users2_UserFullName As String = ""
    Private _aspnet_Users3_UserFullName As String = ""
    Private _IDM_Projects4_Description As String = ""
    Private _PAK_Divisions5_Description As String = ""
    Private _PAK_POStatus6_Description As String = ""
    Private _PAK_POTypes7_Description As String = ""
    Private _PAK_Reasons8_Description As String = ""
    Private _VR_BusinessPartner9_BPName As String = ""
    Public Property POFOR As String = "PK"
    Public Property QCRequired As Boolean = False
    Public Property AcceptedBySupplier As Boolean = False
    Public Property AcceptedBySupplierOn As String = ""
    Public Property POWeight As Decimal = 0
    Public Property PortRequired As Boolean = False
    Public Property PortID As String = ""
    Public Property erpPOStatus As String = ""
    Public Property SerialNo() As Int32
      Get
        Return _SerialNo
      End Get
      Set(ByVal value As Int32)
        _SerialNo = value
      End Set
    End Property
    Public Property POConsignee() As String
      Get
        Return _POConsignee
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _POConsignee = ""
        Else
          _POConsignee = value
        End If
      End Set
    End Property
    Public Property POOtherDetails() As String
      Get
        Return _POOtherDetails
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _POOtherDetails = ""
        Else
          _POOtherDetails = value
        End If
      End Set
    End Property
    Public Property IssueReasonID() As String
      Get
        Return _IssueReasonID
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _IssueReasonID = ""
        Else
          _IssueReasonID = value
        End If
      End Set
    End Property
    Public Property PONumber() As String
      Get
        Return _PONumber
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _PONumber = ""
        Else
          _PONumber = value
        End If
      End Set
    End Property
    Public Property PORevision() As String
      Get
        Return _PORevision
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _PORevision = ""
        Else
          _PORevision = value
        End If
      End Set
    End Property
    Public Property PODate() As String
      Get
        If Not _PODate = String.Empty Then
          Return Convert.ToDateTime(_PODate).ToString("dd/MM/yyyy")
        End If
        Return _PODate
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _PODate = ""
        Else
          _PODate = value
        End If
      End Set
    End Property
    Public Property PODescription() As String
      Get
        Return _PODescription
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _PODescription = ""
        Else
          _PODescription = value
        End If
      End Set
    End Property
    Public Property POTypeID() As String
      Get
        Return _POTypeID
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _POTypeID = ""
        Else
          _POTypeID = value
        End If
      End Set
    End Property
    Public Property SupplierID() As String
      Get
        Return _SupplierID
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _SupplierID = ""
        Else
          _SupplierID = value
        End If
      End Set
    End Property
    Public Property ProjectID() As String
      Get
        Return _ProjectID
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _ProjectID = ""
        Else
          _ProjectID = value
        End If
      End Set
    End Property
    Public Property BuyerID() As String
      Get
        Return _BuyerID
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _BuyerID = ""
        Else
          _BuyerID = value
        End If
      End Set
    End Property
    Public Property POStatusID() As String
      Get
        Return _POStatusID
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _POStatusID = ""
        Else
          _POStatusID = value
        End If
      End Set
    End Property
    Public Property ISGECRemarks() As String
      Get
        Return _ISGECRemarks
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
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
        If Convert.IsDBNull(value) Then
          _SupplierRemarks = ""
        Else
          _SupplierRemarks = value
        End If
      End Set
    End Property
    Public Property IssuedBy() As String
      Get
        Return _IssuedBy
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _IssuedBy = ""
        Else
          _IssuedBy = value
        End If
      End Set
    End Property
    Public Property IssuedOn() As String
      Get
        If Not _IssuedOn = String.Empty Then
          Return Convert.ToDateTime(_IssuedOn).ToString("dd/MM/yyyy")
        End If
        Return _IssuedOn
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _IssuedOn = ""
        Else
          _IssuedOn = value
        End If
      End Set
    End Property
    Public Property ClosedBy() As String
      Get
        Return _ClosedBy
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _ClosedBy = ""
        Else
          _ClosedBy = value
        End If
      End Set
    End Property
    Public Property ClosedOn() As String
      Get
        If Not _ClosedOn = String.Empty Then
          Return Convert.ToDateTime(_ClosedOn).ToString("dd/MM/yyyy")
        End If
        Return _ClosedOn
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _ClosedOn = ""
        Else
          _ClosedOn = value
        End If
      End Set
    End Property
    Public Property DivisionID() As String
      Get
        Return _DivisionID
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _DivisionID = ""
        Else
          _DivisionID = value
        End If
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
    Public Property aspnet_Users3_UserFullName() As String
      Get
        Return _aspnet_Users3_UserFullName
      End Get
      Set(ByVal value As String)
        _aspnet_Users3_UserFullName = value
      End Set
    End Property
    Public Property IDM_Projects4_Description() As String
      Get
        Return _IDM_Projects4_Description
      End Get
      Set(ByVal value As String)
        _IDM_Projects4_Description = value
      End Set
    End Property
    Public Property PAK_Divisions5_Description() As String
      Get
        Return _PAK_Divisions5_Description
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _PAK_Divisions5_Description = ""
        Else
          _PAK_Divisions5_Description = value
        End If
      End Set
    End Property
    Public Property PAK_POStatus6_Description() As String
      Get
        Return _PAK_POStatus6_Description
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _PAK_POStatus6_Description = ""
        Else
          _PAK_POStatus6_Description = value
        End If
      End Set
    End Property
    Public Property PAK_POTypes7_Description() As String
      Get
        Return _PAK_POTypes7_Description
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _PAK_POTypes7_Description = ""
        Else
          _PAK_POTypes7_Description = value
        End If
      End Set
    End Property
    Public Property PAK_Reasons8_Description() As String
      Get
        Return _PAK_Reasons8_Description
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _PAK_Reasons8_Description = ""
        Else
          _PAK_Reasons8_Description = value
        End If
      End Set
    End Property
    Public Property VR_BusinessPartner9_BPName() As String
      Get
        Return _VR_BusinessPartner9_BPName
      End Get
      Set(ByVal value As String)
        _VR_BusinessPartner9_BPName = value
      End Set
    End Property
    Public ReadOnly Property DisplayField() As String
      Get
        Return "" & _PODescription.ToString.PadRight(100, " ")
      End Get
    End Property
    Public ReadOnly Property PrimaryKey() As String
      Get
        Return _SerialNo
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
    Public Class PKpakPO
      Private _SerialNo As Int32 = 0
      Public Property SerialNo() As Int32
        Get
          Return _SerialNo
        End Get
        Set(ByVal value As Int32)
          _SerialNo = value
        End Set
      End Property
    End Class
    Public Shared Function pakPOGetByID(ByVal SerialNo As Int32) As SIS.PAK.pakPO
      Dim Results As SIS.PAK.pakPO = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, SerialNo.ToString.Length, SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakPO(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function

    Public Shared Function pakPOGetByPONumber(ByVal PONumber As String, ByVal forTC As Boolean) As SIS.PAK.pakPO
      Dim Results As SIS.PAK.pakPO = Nothing
      Dim TC As String = IIf(forTC, "TC", "PK")
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppak_LG_POSelectByPONumber"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@forTC", SqlDbType.NVarChar, 2, TC)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@PONumber", SqlDbType.NVarChar, PONumber.Length, PONumber)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakPO(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.PAK.pakPO) As SIS.PAK.pakPO
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_SerialNo", SqlDbType.Int, 11, Record.SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@POConsignee", SqlDbType.NVarChar, 101, IIf(Record.POConsignee = "", Convert.DBNull, Record.POConsignee))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@POOtherDetails", SqlDbType.NVarChar, 501, IIf(Record.POOtherDetails = "", Convert.DBNull, Record.POOtherDetails))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@IssueReasonID", SqlDbType.Int, 11, IIf(Record.IssueReasonID = "", Convert.DBNull, Record.IssueReasonID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@PONumber", SqlDbType.NVarChar, 11, IIf(Record.PONumber = "", Convert.DBNull, Record.PONumber))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@PORevision", SqlDbType.NVarChar, 11, IIf(Record.PORevision = "", Convert.DBNull, Record.PORevision))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@PODate", SqlDbType.DateTime, 21, IIf(Record.PODate = "", Convert.DBNull, Record.PODate))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@PODescription", SqlDbType.NVarChar, 101, IIf(Record.PODescription = "", Convert.DBNull, Record.PODescription))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@POTypeID", SqlDbType.Int, 11, IIf(Record.POTypeID = "", Convert.DBNull, Record.POTypeID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierID", SqlDbType.NVarChar, 10, IIf(Record.SupplierID = "", Convert.DBNull, Record.SupplierID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ProjectID", SqlDbType.NVarChar, 7, IIf(Record.ProjectID = "", Convert.DBNull, Record.ProjectID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@BuyerID", SqlDbType.NVarChar, 9, IIf(Record.BuyerID = "", Convert.DBNull, Record.BuyerID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@POStatusID", SqlDbType.Int, 11, IIf(Record.POStatusID = "", Convert.DBNull, Record.POStatusID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ISGECRemarks", SqlDbType.NVarChar, 501, IIf(Record.ISGECRemarks = "", Convert.DBNull, Record.ISGECRemarks))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SupplierRemarks", SqlDbType.NVarChar, 501, IIf(Record.SupplierRemarks = "", Convert.DBNull, Record.SupplierRemarks))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@IssuedBy", SqlDbType.NVarChar, 9, IIf(Record.IssuedBy = "", Convert.DBNull, Record.IssuedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@IssuedOn", SqlDbType.DateTime, 21, IIf(Record.IssuedOn = "", Convert.DBNull, Record.IssuedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ClosedBy", SqlDbType.NVarChar, 9, IIf(Record.ClosedBy = "", Convert.DBNull, Record.ClosedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ClosedOn", SqlDbType.DateTime, 21, IIf(Record.ClosedOn = "", Convert.DBNull, Record.ClosedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DivisionID", SqlDbType.Int, 11, IIf(Record.DivisionID = "", Convert.DBNull, Record.DivisionID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@POFOR", SqlDbType.NVarChar, 2, Record.POFOR)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@QCRequired", SqlDbType.Bit, 3, Record.QCRequired)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedBySupplier", SqlDbType.Bit, 3, Record.AcceptedBySupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcceptedBySupplierOn", SqlDbType.DateTime, 21, IIf(Record.AcceptedBySupplierOn = "", Convert.DBNull, Record.AcceptedBySupplierOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@POWeight", SqlDbType.Decimal, 25, Record.POWeight)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@PortRequired", SqlDbType.Bit, 3, Record.PortRequired)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@PortID", SqlDbType.Int, 11, IIf(Record.PortID = "", Convert.DBNull, Record.PortID))
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
    Public Shared Function GetTotalWeight(ByVal Qty As String, ByVal UnitWt As String, ByVal UOMqty As String, ByVal UOMWt As String) As Decimal
      Dim mRet As Decimal = 0
      Dim mc As SIS.PAK.UnitMultiplicationFactor = Nothing
      Try
        Dim tmpUnit As SIS.PAK.pakUnits = SIS.PAK.pakUnits.pakUnitsGetByID(UOMqty)
        If tmpUnit.UnitSetID = 3 Then '3=>Weight Unit Set
          Try
            mc = New SIS.PAK.UnitMultiplicationFactor
            mc.Unit = tmpUnit
            mc = SIS.PAK.UnitMultiplicationFactor.GetMultiplicationFactorToBaseUnit(mc)
            mRet = Qty * mc.MF
          Catch ex As Exception
          End Try
        Else
          Try
            mc = New SIS.PAK.UnitMultiplicationFactor
            mc.Unit = SIS.PAK.pakUnits.pakUnitsGetByID(UOMWt)
            mc = SIS.PAK.UnitMultiplicationFactor.GetMultiplicationFactorToBaseUnit(mc)
            mRet = Qty * UnitWt * mc.MF
          Catch ex As Exception
          End Try
        End If
      Catch ex As Exception
        Try
          mc = New SIS.PAK.UnitMultiplicationFactor
          mc.Unit = SIS.PAK.pakUnits.pakUnitsGetByID(UOMWt)
          mc = SIS.PAK.UnitMultiplicationFactor.GetMultiplicationFactorToBaseUnit(mc)
          mRet = Qty * UnitWt * mc.MF
        Catch ex1 As Exception
        End Try
      End Try
      Return mRet
    End Function

  End Class
End Namespace
