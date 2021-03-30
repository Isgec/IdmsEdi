Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports EDICommon
Namespace SIS.PAK
  Partial Public Class pakPOBItems
    Public ReadOnly Property pFree As String
      Get
        If Free Then
          Return "*"
        End If
        Return ""
      End Get
    End Property
    Public ReadOnly Property PWeightPerUnit As String
      Get
        If WeightPerUnit <= 0 Then
          Return ""
        Else
          Return WeightPerUnit
        End If
      End Get
    End Property
    Public ReadOnly Property PQuantity As String
      Get
        If Quantity <= 0 Then
          Return ""
        Else
          Return Quantity
        End If
      End Get
    End Property
    Public ReadOnly Property PItemDescription As String
      Get
        Return Prefix & " " & ItemDescription
      End Get
    End Property
    Public ReadOnly Property FontBold As Boolean
      Get
        Dim mRet As Boolean = False
        If Not Bottom Then
          mRet = True
        End If
        Return mRet
      End Get
    End Property
    Public Function GetVisible() As Boolean
      Dim mRet As Boolean = True
      Return mRet
    End Function
    Public Function GetEnable() As Boolean
      Dim mRet As Boolean = True
      Return mRet
    End Function
    Public Function GetEditable() As Boolean
      Dim mRet As Boolean = True
      Return mRet
    End Function
    Public ReadOnly Property Deleteable() As Boolean
      Get
        Dim mRet As Boolean = False
        Try
        Catch ex As Exception
        End Try
        Return mRet
      End Get
    End Property
    Public ReadOnly Property DeleteWFEnable() As Boolean
      Get
        Dim mRet As Boolean = True
        Try
          mRet = GetEnable()
        Catch ex As Exception
        End Try
        Return mRet
      End Get
    End Property
    Private Shared Sub rDeleteWF(ByVal pItm As SIS.PAK.pakPOBItems)
      Dim Items As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.GetByParentPOBItemNo(pItm.SerialNo, pItm.BOMNo, pItm.ItemNo, "")
      If Items.Count > 0 Then
        For Each Itm As SIS.PAK.pakPOBItems In Items
          rDeleteWF(Itm)
        Next
        SIS.PAK.pakPOBItems.pakPOBItemsDelete(pItm)
      Else
        SIS.PAK.pakPOBItems.pakPOBItemsDelete(pItm)
      End If
    End Sub
    Public Shared Function DeleteWF(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemNo As Int32) As SIS.PAK.pakPOBItems
      Dim Results As SIS.PAK.pakPOBItems = pakPOBItemsGetByID(SerialNo, BOMNo, ItemNo)
      rDeleteWF(Results)
      Dim tmpP As SIS.PAK.pakPOBItems = SIS.PAK.pakPOBItems.pakPOBItemsGetByID(Results.SerialNo, Results.BOMNo, Results.ParentItemNo)
      If tmpP IsNot Nothing Then
        If Not tmpP.Root Then
          Dim tmpPs As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.GetByParentPOBItemNo(Results.SerialNo, Results.BOMNo, Results.ParentItemNo, "")
          If tmpPs.Count <= 0 Then
            tmpP.Middle = False
            tmpP.Bottom = True
            tmpP = UpdateData(tmpP)
          End If
        End If
      End If
      Return Results
    End Function
    Public Shared Function GetMaxItemNo(ByVal SerialNo As Int32, ByVal BOMNo As Int32) As Integer
      Dim mRet As Integer = 1
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "select MAX(ISNULL(ItemNo,1)) as cnt from pak_POBItems where serialno=" & SerialNo & " and bomno=" & BOMNo
          Con.Open()
          mRet = Cmd.ExecuteScalar()
          mRet = mRet + 1
        End Using
      End Using
      Return mRet
    End Function

    Private Shared Sub rCopyCWF(ByVal pItm As SIS.PAK.pakPOBItems, ByVal ParentItemNo As Integer)
      Dim Items As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.GetByParentPOBItemNo(pItm.SerialNo, pItm.BOMNo, ParentItemNo, "")
      If Items.Count > 0 Then
        For Each Itm As SIS.PAK.pakPOBItems In Items
          Dim ItemNo As Integer = Itm.ItemNo
          Itm.ItemNo = GetMaxItemNo(Itm.SerialNo, Itm.BOMNo)
          Itm.ParentItemNo = pItm.ItemNo
          Itm.Free = True
          Itm.StatusID = pakItemStates.CreatedByISGEC
          Itm = pakPOBItems.pakPOBItemsInsert(Itm)
          rCopyCWF(Itm, ItemNo)
        Next
      End If
    End Sub

    Public Shared Function CopyCWF(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemNo As Int32) As SIS.PAK.pakPOBItems
      Dim Results As SIS.PAK.pakPOBItems = pakPOBItemsGetByID(SerialNo, BOMNo, ItemNo)
      With Results
        .ItemNo = GetMaxItemNo(SerialNo, BOMNo)
        .Free = True
        .StatusID = pakItemStates.CreatedByISGEC
      End With
      Results = pakPOBItems.pakPOBItemsInsert(Results)
      rCopyCWF(Results, ItemNo)
      Return Results
    End Function
    Public ReadOnly Property InitiateWFVisible() As Boolean
      Get
        Dim mRet As Boolean = True
        Try
          mRet = GetVisible()
        Catch ex As Exception
        End Try
        Return mRet
      End Get
    End Property
    Public ReadOnly Property InitiateWFEnable() As Boolean
      Get
        Dim mRet As Boolean = True
        Try
          mRet = GetEnable()
        Catch ex As Exception
        End Try
        Return mRet
      End Get
    End Property
    Public ReadOnly Property ApproveWFEnable() As Boolean
      Get
        Dim mRet As Boolean = True
        Try
          mRet = GetEnable()
        Catch ex As Exception
        End Try
        Return mRet
      End Get
    End Property
    Public ReadOnly Property RejectWFEnable() As Boolean
      Get
        Dim mRet As Boolean = True
        Try
          mRet = GetEnable()
        Catch ex As Exception
        End Try
        Return mRet
      End Get
    End Property
    Public ReadOnly Property CompleteWFVisible() As Boolean
      Get
        Dim mRet As Boolean = True
        Try
          mRet = GetVisible()
        Catch ex As Exception
        End Try
        Return mRet
      End Get
    End Property
    Public ReadOnly Property CompleteWFEnable() As Boolean
      Get
        Dim mRet As Boolean = True
        Try
          mRet = GetEnable()
        Catch ex As Exception
        End Try
        Return mRet
      End Get
    End Property
    Public Shared Function InitiateWF(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemNo As Int32) As SIS.PAK.pakPOBItems
      Dim Results As SIS.PAK.pakPOBItems = pakPOBItemsGetByID(SerialNo, BOMNo, ItemNo)
      Return Results
    End Function
    Private Shared Sub rApproveWF(ByVal pItm As SIS.PAK.pakPOBItems, ByVal ParentItemNo As Integer)
      Dim Items As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.GetByParentPOBItemNo(pItm.SerialNo, pItm.BOMNo, ParentItemNo, "")
      If Items.Count > 0 Then
        For Each Itm As SIS.PAK.pakPOBItems In Items
          With Itm
            .StatusID = pakItemStates.FreezedbyISGEC
            .FreezedBy = ""
            .FreezedOn = Now
            .Freezed = True
            .Changed = False
          End With
          Itm = pakPOBItems.UpdateData(Itm)
          rApproveWF(Itm, Itm.ItemNo)
        Next
      End If
    End Sub
    Public Shared Function ApproveWF(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemNo As Int32) As SIS.PAK.pakPOBItems
      Dim Results As SIS.PAK.pakPOBItems = pakPOBItemsGetByID(SerialNo, BOMNo, ItemNo)
      With Results
        .StatusID = pakItemStates.FreezedbyISGEC
        .FreezedBy = ""
        .FreezedOn = Now
        .Freezed = True
        .Changed = False
      End With
      Results = pakPOBItems.UpdateData(Results)
      rApproveWF(Results, ItemNo)
      Return Results
    End Function
    Private Shared Sub rRejectWF(ByVal pItm As SIS.PAK.pakPOBItems, ByVal ParentItemNo As Integer)
      Dim Items As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.GetByParentPOBItemNo(pItm.SerialNo, pItm.BOMNo, ParentItemNo, "")
      If Items.Count > 0 Then
        For Each Itm As SIS.PAK.pakPOBItems In Items
          With Itm
            .StatusID = pakItemStates.ChangeRequiredByISGEC
            .FreezedBy = ""
            .FreezedOn = Now
            .Freezed = False
            .Changed = True
          End With
          Itm = pakPOBItems.UpdateData(Itm)
          rRejectWF(Itm, Itm.ItemNo)
        Next
      End If
    End Sub
    Public Shared Function RejectWF(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemNo As Int32) As SIS.PAK.pakPOBItems
      Dim Results As SIS.PAK.pakPOBItems = pakPOBItemsGetByID(SerialNo, BOMNo, ItemNo)
      With Results
        .StatusID = pakItemStates.ChangeRequiredByISGEC
        .FreezedBy = ""
        .FreezedOn = Now
        .Freezed = False
        .Changed = True
      End With
      Results = pakPOBItems.UpdateData(Results)
      rRejectWF(Results, ItemNo)
      Return Results
    End Function
    Public Shared Function CompleteWF(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemNo As Int32) As SIS.PAK.pakPOBItems
      Dim Results As SIS.PAK.pakPOBItems = pakPOBItemsGetByID(SerialNo, BOMNo, ItemNo)
      Return Results
    End Function
    Public Shared Function GetpakPOBItemsCreatedByISGECCount(ByVal SerialNo As Int32, ByVal BOMNo As Int32) As Integer
      Dim Results As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select Count(ItemNo) as cnt from pak_pobItems where serialno=" & SerialNo & " and bomno=" & BOMNo & " and bottom=1 and StatusID=" & pakItemStates.CreatedByISGEC
          Con.Open()
          Results = Cmd.ExecuteScalar()
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function GetpakPOBItemsUnFreezedCount(ByVal SerialNo As Int32, ByVal BOMNo As Int32) As Integer
      Dim Results As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select Count(ItemNo) as cnt from pak_pobItems where serialno=" & SerialNo & " and bomno=" & BOMNo & " and bottom=1 and Freezed=0"
          Con.Open()
          Results = Cmd.ExecuteScalar()
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function GetpakPOBItemsUnCheckedISGECCount(ByVal SerialNo As Int32, ByVal BOMNo As Int32) As Integer
      Dim Results As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select Count(ItemNo) as cnt from pak_pobItems where serialno=" & SerialNo & " and bomno=" & BOMNo & " and bottom=1 and StatusID IN (" & pakItemStates.ChangeRequiredBySupplier & "," & pakItemStates.VerifiedbySupplier & " )"
          Con.Open()
          Results = Cmd.ExecuteScalar()
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function pakPOBItemsDeleteAll(ByVal SerialNo As Int32, ByVal BOMNo As Int32) As Integer
      Dim Results As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Delete from pak_pobItems where serialno=" & SerialNo & " and bomno=" & BOMNo
          Con.Open()
          Results = Cmd.ExecuteNonQuery()
        End Using
      End Using
      Return Results
    End Function

    Public Shared Function GetByParentPOBItemNo(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ParentItemNo As Int32, ByVal OrderBy As String) As List(Of SIS.PAK.pakPOBItems)
      Dim Results As List(Of SIS.PAK.pakPOBItems) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppakPOBItemsSelectByParentItemNo"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@BOMNo", SqlDbType.Int, 10, IIf(BOMNo = Nothing, 0, BOMNo))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, 10, IIf(SerialNo = Nothing, 0, SerialNo))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentItemNo", SqlDbType.Int, ParentItemNo.ToString.Length, ParentItemNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          RecordCount = -1
          Results = New List(Of SIS.PAK.pakPOBItems)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakPOBItems(Reader))
          End While
          Reader.Close()
          RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Private Shared Sub getpakHBItems(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal pItem As SIS.PAK.pakPOBItems, ByRef cList As List(Of SIS.PAK.pakPOBItems))
      cList.Add(pItem)
      Dim Results As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.GetByParentPOBItemNo(SerialNo, BOMNo, pItem.ItemNo, "")
      For Each tmp As SIS.PAK.pakPOBItems In Results
        getpakHBItems(SerialNo, BOMNo, tmp, cList)
      Next
    End Sub
    Public Shared Function UZ_pakPOBItemsSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal BOMNo As Int32, ByVal SerialNo As Int32) As List(Of SIS.PAK.pakPOBItems)
      Dim tmp As SIS.PAK.pakPOBOM = SIS.PAK.pakPOBOM.pakPOBOMGetByID(SerialNo, BOMNo)
      Dim Results As New List(Of SIS.PAK.pakPOBItems)
      Dim rItm As SIS.PAK.pakPOBItems = SIS.PAK.pakPOBItems.pakPOBItemsGetByID(SerialNo, BOMNo, tmp.ItemNo)
      getpakHBItems(SerialNo, BOMNo, rItm, Results)
      RecordCount = Results.Count
      Return Results
    End Function
    Public Shared Function UZ_pakPOBItemsDelete(ByVal Record As SIS.PAK.pakPOBItems) As Integer
      Dim _Result As Integer = pakPOBItemsDelete(Record)
      Return _Result
    End Function
    Public Shared Function pakPOBItemsGetByItemCode(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemCode As String) As SIS.PAK.pakPOBItems
      Dim Results As SIS.PAK.pakPOBItems = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppak_LG_POBItemsSelectByItemCode"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, SerialNo.ToString.Length, SerialNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@BOMNo", SqlDbType.Int, BOMNo.ToString.Length, BOMNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemCode", SqlDbType.NVarChar, ItemCode.Length, ItemCode)
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
    Public Shared Function GetByDocRev(serialNo As Integer, ItemNo As Integer, doc As String, rev As String) As List(Of SIS.PAK.pakPOBItems)
      Dim mRet As New List(Of SIS.PAK.pakPOBItems)
      Dim Sql As String = ""
      Sql &= " select a.* "
      Sql &= " from PAK_POBItems as a"
      Sql &= " inner join PAK_Documents as b "
      Sql &= " on a.DocumentNo = b.DocumentNo "
      Sql &= " where a.SerialNo=" & serialNo
      Sql &= " and a.ParentItemNo=" & ItemNo
      Sql &= " and b.DocumentID='" & doc & "'"
      Sql &= " and b.DocumentRevision =(select isnull(max(c.DocumentRevision),'') from PAK_Documents as c where c.DocumentID=b.DocumentID and c.DocumentRevision < '" & rev & "')"
      Sql &= ""
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = Sql
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            mRet.Add(New SIS.PAK.pakPOBItems(Reader))
          End While
          Reader.Close()
        End Using
      End Using
      Return mRet
    End Function
    Public Shared Function CreateNewRevDocumentAndLink(serialNo As Integer, ItemNo As Integer, doc As String, n_rev As String) As List(Of SIS.PAK.pakPOBItems)
      Dim mRet As New List(Of SIS.PAK.pakPOBItems)
      Dim Sql As String = ""
      Dim NewDocNo As String = ""
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Con.Open()
        '1. Create and select New Revision`s documntno
        Sql = ""
        Sql &= " Insert Into PAK_Documents "
        Sql &= " (DocumentID, DocumentRevision, Description, ExternalDocument, Filename, DisskFile)"
        Sql &= " SELECT TOP 1 DocumentID, '" & n_rev & "' as DR, Description, 0 AS ed, NULL AS fn, NULL AS df "
        Sql &= " from pak_documents where "
        Sql &= " DocumentID='" & doc & "'"
        Sql &= " and DocumentRevision =(select isnull(max(c.DocumentRevision),'') from PAK_Documents as c where c.DocumentID='" & doc & "' and c.DocumentRevision < '" & n_rev & "')"
        Sql &= " Select * From PAK_Documents "
        Sql &= " where "
        Sql &= " and DocumentID='" & doc & "'"
        Sql &= " and DocumentRevision ='" & n_rev & "'"
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = Sql
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            NewDocNo = Reader("DocumentNo")
          End While
          Reader.Close()
        End Using

        Sql = ""
        Sql &= " update "
        Sql &= " PAK_POBItems "
        Sql &= " set DocumentNo=" & NewDocNo
        Sql &= " where SerialNo=" & serialNo
        Sql &= " and ParentItemNo=" & ItemNo
        Sql &= " and DocumentNo=("
        Sql &= " SELECT TOP 1 DocumentNo from pak_documents "
        Sql &= "   where "
        Sql &= "   DocumentID='" & doc & "'"
        Sql &= "   and DocumentRevision =(select isnull(max(c.DocumentRevision),'') from PAK_Documents as c where c.DocumentID='" & doc & "' and c.DocumentRevision < '" & n_rev & "')"
        Sql &= "   )"
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = Sql
          Cmd.ExecuteNonQuery()
        End Using
      End Using
      Return mRet
    End Function
  End Class
End Namespace
