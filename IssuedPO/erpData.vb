Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK

#Region " ERP Components Classes "
  Public Class erpData
    Public Class erpPO
      Public Delegate Sub showMsg(ByVal str As String)
      Public Shared Function pakPOBOMGetByERPPOLine(ByVal SerialNo As Int32, ByVal ItemNo As Int32) As SIS.PAK.pakPOBOM
        Dim Results As SIS.PAK.pakPOBOM = Nothing
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.StoredProcedure
            Cmd.CommandText = "sppak_LG_POBOMSelectByERPPOLine"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, SerialNo.ToString.Length, SerialNo)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemNo", SqlDbType.Int, ItemNo.ToString.Length, ItemNo)
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
      Public Shared Sub UpdateIssuedPO(TmtlID As String, Comp As String, msg As showMsg)
        Dim POs As List(Of EffectedPOByDrawings) = EffectedPOByDrawings.GetPOList(Comp, TmtlID)
        If POs.Count > 0 Then
          Dim MoreThanOnePO As Boolean = IIf(POs.Count > 1, True, False)
          For Each PO As EffectedPOByDrawings In POs
            '===========================Update From PO==============
            '======UpdateIssuedPOFromERPPO(PO.PONumber, Comp, msg)
            '===========================Needs Code ReVisit==========
            UpdateIssuedPOFromDrawing(PO, Comp, msg, MoreThanOnePO)
          Next
        End If
      End Sub

      Public Shared Sub UpdateIssuedPOFromDrawing(ePO As EffectedPOByDrawings, Comp As String, msg As showMsg, MoreThanOnePO As Boolean)
        Dim PO As SIS.PAK.pakPO = SIS.PAK.erpData.erpPO.pakPOGetByPONumber(ePO.PONumber, "PK")
        If PO IsNot Nothing Then
          Try
            Select Case PO.POTypeID
              Case pakErpPOTypes.ISGECEngineered, pakErpPOTypes.Boughtout
                For Each dwg As TmtlDrawings In ePO.Drawings
                  Dim dmItems As List(Of dmisg002) = dmisg002.GetItems(dwg.DocumentID, dwg.RevisionNo, Comp)
                  Dim bItems As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.GetByDocRev(PO.SerialNo, ePO.POLine, dwg.DocumentID, dwg.RevisionNo)
                  Dim Found As Boolean = False
                  For Each dmItem As dmisg002 In dmItems
                    Found = False
                    Dim bItem As SIS.PAK.pakPOBItems = Nothing
                    For Each bItem In bItems
                      If bItem.ItemCode.Trim = dmItem.t_item.Trim Then
                        Found = True
                        Exit For
                      End If
                    Next
                    If Not Found Then
                      If MoreThanOnePO Then
                        Continue For 'Do Not Insert New Item
                      End If
                      bItem = bItems(0)
                      Dim mNextItemNo As Integer = GetNextItemNo(bItem.SerialNo, bItem.BOMNo)
                      With bItem
                        .ItemNo = mNextItemNo
                        .ItemCode = dmItem.t_item.Trim
                        .ItemDescription = dmItem.t_dsca
                        .QuantityToPack = 0
                        .TotalWeightToPack = 0
                        .QuantityToDespatch = 0
                        .TotalWeightToDespatch = 0
                        .QuantityDespatched = 0
                        .TotalWeightDespatched = 0
                        .QuantityReceived = 0
                        .TotalWeightReceived = 0
                      End With
                    End If
                    With bItem
                      Try
                        .ItemDescription = dmItem.t_dsca
                        If dmItem.t_cuni <> "" Then
                          .UOMQuantity = SIS.PAK.pakUnits.pakUnitsGetByDescription(dmItem.t_cuni).UnitID
                        End If
                        If dmItem.t_qnty < 0.0001 Then dmItem.t_qnty = 0
                        If PO.QCRequired Then
                          If dmItem.t_qnty < bItem.QualityClearedQty Then
                            dmItem.t_qnty = bItem.QualityClearedQty
                          End If
                        Else
                          If dmItem.t_qnty < bItem.QuantityDespatched Then
                            dmItem.t_qnty = bItem.QuantityDespatched
                          End If
                        End If
                        .Quantity = dmItem.t_qnty
                        .UOMWeight = SIS.PAK.pakUnits.pakUnitsGetByDescription("kg").UnitID
                        If dmItem.t_wght >= 0.0001 AndAlso dmItem.t_qnty >= 0.0001 Then
                          .WeightPerUnit = dmItem.t_wght / dmItem.t_qnty
                        Else
                          .WeightPerUnit = 0
                        End If
                      Catch ex As Exception
                        Throw New Exception("Err Updating: " & ex.Message)
                      End Try
                    End With
                    If Found Then
                      bItem = SIS.PAK.pakPOBItems.UpdateData(bItem)
                    Else
                      bItem = SIS.PAK.pakPOBItems.InsertData(bItem)
                    End If
                  Next
                  'Update Document Revision
                  SIS.PAK.pakPOBItems.CreateNewRevDocumentAndLink(PO.SerialNo, ePO.POLine, dwg.DocumentID, dwg.RevisionNo)
                  If MoreThanOnePO Then
                    Continue For 'Do Not Reverse Check to delete
                  End If
                  'Reverse check to Delete Items Removed from Drawing
                  For Each bItem As SIS.PAK.pakPOBItems In bItems
                    Found = False
                    For Each dmItem As dmisg002 In dmItems
                      If bItem.ItemCode.Trim = dmItem.t_item.Trim Then
                        Found = True
                        Exit For
                      End If
                    Next
                    If Not Found Then
                      If SIS.PAK.pakPOBItems.CanBeDeleted(bItem) Then
                        SIS.PAK.pakPOBItems.pakPOBItemsDeleteRecursive(bItem)
                      Else
                        bItem.DeletedInERP = True
                        SIS.PAK.pakPOBItems.UpdateData(bItem)
                      End If
                    End If
                  Next
                  'End of Reverse Check
                Next 'Drawing
            End Select
          Catch ex As Exception
            msg(ex.Message)
          End Try
        End If
      End Sub

      Public Shared Function UpdateIssuedPOFromERPPO(PONumber As String, Comp As String, msg As showMsg) As SIS.PAK.pakPO
        'It is stopped and on transmittal dwg base is used
        Dim oPO As SIS.PAK.pakPO = SIS.PAK.erpData.erpPO.pakPOGetByPONumber(PONumber, "PK")
        If oPO IsNot Nothing Then
          Try
            ImportFromERP(oPO.PONumber, False, IIf(oPO.POTypeID = pakErpPOTypes.ISGECEngineered, True, False), IIf(oPO.POTypeID = pakErpPOTypes.Boughtout, True, False), Comp)
          Catch ex As Exception
            msg(ex.Message)
          End Try
        End If
        Return oPO
      End Function
      'Main Function PO Import
      Public Shared Function ImportFromERP(PONumber As String, ForTC As Boolean, AsIsgecEngineered As Boolean, AsBoughtOut As Boolean, comp As String) As SIS.PAK.pakPO
        Dim oPO As SIS.PAK.pakPO = Nothing
        Dim oePO As SIS.PAK.pakPO = Nothing
        'Get ERP PO
        Try
          oePO = SIS.PAK.erpData.erpPO.GetFromERP(PONumber, comp)
        Catch ex As Exception
          Throw New Exception("Error when fetching PO from ERP: " & ex.Message)
        End Try
        If oePO Is Nothing Then
          Throw New Exception("PO Not found in ERP Company [PO Must be locked by Audit]: " & comp)
        End If
        'Get PAK PO
        oPO = SIS.PAK.erpData.erpPO.pakPOGetByPONumber(PONumber, ForTC)
        oPO.PORevision = oePO.PORevision
        If oePO.erpPOStatus = "30" Then
          oPO.POStatusID = 5
        End If
        oePO = SIS.PAK.pakPO.UpdateData(oPO)
        Select Case oePO.POTypeID
          Case pakErpPOTypes.ISGECEngineered, pakErpPOTypes.Boughtout
            Dim xPOLines As List(Of SIS.PAK.pakERPPOLine) = SIS.PAK.erpData.erpPO.GetPOLine(oePO, comp)
            For Each POL As SIS.PAK.pakERPPOLine In xPOLines
              '1. Check Element
              'New with Pack Elements, For Tree Using Elements
              Dim oEle As SIS.PAK.pakElements = SIS.PAK.pakElements.pakElementsGetByID(POL.t_cspa)
              If oEle Is Nothing Then
                oEle = SIS.PAK.erpData.erpPAKElement.GetFromERP(POL.t_cspa, comp)
                oEle.ElementLevel = 1
                oEle.ResponsibleAgencyID = 1
                oEle = SIS.PAK.pakElements.InsertData(oEle)
              End If
              '2. Check POLine exists in POBOM
              Dim POBOM As SIS.PAK.pakPOBOM = SIS.PAK.erpData.erpPO.pakPOBOMGetByERPPOLine(oePO.SerialNo, POL.t_pono)
              Dim Found As Boolean = True
              Dim BomNo As Integer = 0
              If POBOM Is Nothing Then
                Found = False
              Else
                BomNo = POBOM.BOMNo
              End If
              POBOM = SIS.PAK.pakERPPOLine.GetPOBOM(POL, oePO)
              POBOM.BOMNo = BomNo
              If Found Then
                POBOM = SIS.PAK.pakPOBOM.UpdateData(POBOM)
              Else
                POBOM = SIS.PAK.pakPOBOM.InsertData(POBOM)
              End If
              '3. Insert POBOM in BItem
              Dim POBItem As SIS.PAK.pakPOBItems = SIS.PAK.pakPOBOM.GetPOBItem(POBOM)
              If Found Then
                POBItem = SIS.PAK.pakPOBItems.UpdateData(POBItem)
              Else
                POBItem = SIS.PAK.pakPOBItems.InsertData(POBItem)
              End If
              '4. Child Items of POBOM
              Dim xPOLChild As List(Of SIS.PAK.pakERPPOLChild) = SIS.PAK.erpData.erpPO.GetPOLChild(oePO, POBOM, comp)
              For Each POLC As SIS.PAK.pakERPPOLChild In xPOLChild
                POBItem = SIS.PAK.pakERPPOLChild.GetPOBItems(POLC, oePO, POBOM)
                Found = True
                Dim xPOBItem As SIS.PAK.pakPOBItems = SIS.PAK.pakPOBItems.pakPOBItemsGetByItemCode(POBItem.SerialNo, POBItem.BOMNo, POBItem.ItemCode)
                If xPOBItem Is Nothing Then
                  xPOBItem = New SIS.PAK.pakPOBItems
                  Found = False
                End If
                If Found Then
                  With xPOBItem
                    Try
                      .ItemDescription = POLC.t_desc
                      If POLC.t_quom <> "" Then
                        .UOMQuantity = SIS.PAK.pakUnits.pakUnitsGetByDescription(POLC.t_quom).UnitID
                      End If
                      If POLC.t_qnty < 0.0001 Then POLC.t_qnty = 0
                      If oPO.QCRequired Then
                        If POLC.t_qnty < xPOBItem.QualityClearedQty Then
                          POLC.t_qnty = xPOBItem.QualityClearedQty
                        End If
                      Else
                        If POLC.t_qnty < xPOBItem.QuantityDespatched Then
                          POLC.t_qnty = xPOBItem.QuantityDespatched
                        End If
                      End If
                      .Quantity = POLC.t_qnty
                      .UOMWeight = SIS.PAK.pakUnits.pakUnitsGetByDescription("kg").UnitID
                      If POLC.t_wght >= 0.0001 AndAlso POLC.t_qnty >= 0.0001 Then
                        .WeightPerUnit = POLC.t_wght / POLC.t_qnty
                      Else
                        .WeightPerUnit = 0
                      End If
                      .ItemReference = POLC.t_iref
                      .SubItem = POLC.t_sitm
                      .SubItem2 = POLC.t_sub2
                      .SubItem3 = POLC.t_sub3
                      .SubItem4 = POLC.t_sub4
                    Catch ex As Exception
                      Throw New Exception("Err Assign Child Item: " & ex.Message)
                    End Try
                    'Create or Get DocumentNO
                    Dim tmpD As SIS.PAK.pakDocuments = Nothing
                    Dim tmpDoc As SIS.PAK.pakDocuments = SIS.PAK.pakDocuments.pakDocumentsSelectByDocRevID(POLC.t_docn, POLC.t_revi)
                    If tmpDoc Is Nothing Then
                      tmpD = New SIS.PAK.pakDocuments
                      With tmpD
                        .DocumentID = POLC.t_docn
                        .DocumentRevision = POLC.t_revi
                        .Description = POLC.DocumentDescription
                        .ExternalDocument = False
                        .DisskFile = ""
                        .Filename = ""
                      End With
                      tmpD = SIS.PAK.pakDocuments.InsertData(tmpD)
                      .DocumentNo = tmpD.DocumentNo
                    Else
                      tmpD = tmpDoc
                      With tmpD
                        .Description = POLC.DocumentDescription
                      End With
                      tmpD = SIS.PAK.pakDocuments.UpdateData(tmpD)
                      .DocumentNo = tmpD.DocumentNo
                    End If
                  End With
                  xPOBItem = SIS.PAK.pakPOBItems.UpdateData(xPOBItem)
                Else
                  'Find Next AVBL ItemNo
                  Dim mNextItemNo As Integer = GetNextItemNo(POBItem.SerialNo, POBItem.BOMNo)
                  POBItem.ItemNo = mNextItemNo
                  POBItem = SIS.PAK.pakPOBItems.InsertData(POBItem)
                End If
              Next
              '5. Reverse check POBOMItems, if Not found in POLchild then delete, only when there is no despatch
              '   else mark Deleted in ERP
              Dim rChilds As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.UZ_pakPOBItemsSelectList(0, 99999, "", False, "", POBOM.BOMNo, POBOM.SerialNo)
              For Each rChild As SIS.PAK.pakPOBItems In rChilds
                If rChild.Bottom Then
                  Dim biFound As Boolean = False
                  For Each POLC As SIS.PAK.pakERPPOLChild In xPOLChild
                    If POLC.t_pono = rChild.ParentItemNo AndAlso rChild.ItemCode = POLC.t_item.Trim Then
                      biFound = True
                      Exit For
                    End If
                  Next
                  If Not biFound Then
                    If oPO.QCRequired Then
                      If rChild.QualityClearedQty <= 0.00 Then
                        'Delete when there is no quality cleared qty
                        SIS.PAK.pakPOBItems.pakPOBItemsDeleteRecursive(rChild)
                      Else
                        rChild.DeletedInERP = True
                        SIS.PAK.pakPOBItems.UpdateData(rChild)
                      End If
                    Else
                      If rChild.QuantityDespatched <= 0.00 Then
                        'Delete when there is no despatch
                        SIS.PAK.pakPOBItems.pakPOBItemsDeleteRecursive(rChild)
                      Else
                        rChild.DeletedInERP = True
                        SIS.PAK.pakPOBItems.UpdateData(rChild)
                      End If
                    End If
                  End If
                End If
              Next
            Next
            'Reverse Check POBOM
            Dim rBOMs As List(Of SIS.PAK.pakPOBOM) = SIS.PAK.pakPOBOM.UZ_pakPOBOMSelectList(0, 9999, "", False, "", oPO.SerialNo)
            For Each rBOM As SIS.PAK.pakPOBOM In rBOMs
              Dim Found As Boolean = False
              For Each POL As SIS.PAK.pakERPPOLine In xPOLines
                If rBOM.ItemNo = POL.t_pono AndAlso rBOM.ItemCode = POL.t_item Then
                  Found = True
                  Exit For
                End If
              Next
              If Not Found Then
                Dim rChilds As List(Of SIS.PAK.pakPOBItems) = SIS.PAK.pakPOBItems.UZ_pakPOBItemsSelectList(0, 99999, "", False, "", rBOM.BOMNo, rBOM.SerialNo)
                Dim ExceptionFound As Boolean = False
                For Each rChild As SIS.PAK.pakPOBItems In rChilds
                  If rChild.Bottom Then
                    If oPO.QCRequired Then
                      If rChild.QualityClearedQty > 0.00 Then
                        ExceptionFound = True
                        Exit For
                      End If
                    Else
                      If rChild.QuantityDespatched > 0.00 Then
                        ExceptionFound = True
                        Exit For
                      End If
                    End If
                  End If
                Next
                If Not ExceptionFound Then
                  For Each rChild As SIS.PAK.pakPOBItems In rChilds
                    SIS.PAK.pakPOBItems.pakPOBItemsDeleteRecursive(rChild)
                  Next
                  SIS.PAK.pakPOBOM.pakPOBOMDelete(rBOM)
                End If
              End If
            Next
        End Select
        Return oePO
      End Function
      Private Shared Function GetNextItemNo(ByVal SerialNo As Integer, ByVal BomNo As Integer) As Integer
        Dim tmp As Integer = 0
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString)
          Using Cmd As SqlCommand = Con.CreateCommand()
            Dim mSql As String = "SELECT MAX(ISNULL(ItemNo,0)) FROM [PAK_POBitems] WHERE SerialNo = " & SerialNo & " and BOMNo=" & BomNo
            Cmd.CommandType = System.Data.CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            Try
              tmp = Cmd.ExecuteScalar()
              tmp += 1
            Catch ex As Exception
              tmp = 1
            End Try
          End Using
        End Using
        Return tmp
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
      Public Shared Function GetPOLine(ByVal po As SIS.PAK.pakPO, comp As String) As List(Of SIS.PAK.pakERPPOLine)
        Dim Results As New List(Of SIS.PAK.pakERPPOLine)
        Dim Sql As String = ""
        Sql &= "select "
        Sql &= "  [po].*,"
        Sql &= "  tc.t_dsca as ItemDescription "
        Sql &= "  from ttdpur401" & Comp & " as po "
        Sql &= "  inner join ttcibd001" & Comp & " as tc on po.t_item = tc.t_item "
        Sql &= "  where po.t_orno =  '" & po.PONumber & "'"
        Sql &= "  and po.t_oltp in (2,4) "     '2=>Detail, 4=>Order Line
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Sql
            Results = New List(Of SIS.PAK.pakERPPOLine)
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            While (Reader.Read())
              Results.Add(New SIS.PAK.pakERPPOLine(Reader))
            End While
            Reader.Close()
          End Using
        End Using
        Return Results
      End Function
      Public Shared Function GetPOLChild(ByVal po As SIS.PAK.pakPO, ByVal poBom As SIS.PAK.pakPOBOM, comp As String) As List(Of SIS.PAK.pakERPPOLChild)
        Dim Results As New List(Of SIS.PAK.pakERPPOLChild)
        Dim Sql As String = ""
        Sql &= "select isnull(cDoc.t_iref,'') as t_iref, isnull(dItm.t_sitm,'') as t_sitm, isnull(iRef.t_sub2,'') as t_sub2, isnull(iRef.t_sub3,'') as t_sub3, isnull(iRef.t_sub4,'') as t_sub4, cItm.*, dm.t_dttl as DocumentDescription "
        Sql &= "from ttdisg002" & comp & " as cItm "
        Sql &= "left outer join tdmisg140" & comp & " as cDoc "
        Sql &= "        on cItm.t_docn = cDoc.t_docn "
        Sql &= "left outer join tdmisg002" & comp & " as dItm "
        Sql &= "        on dItm.t_docn = cItm.t_docn "
        Sql &= "       and dItm.t_revn = cItm.t_revi "
        Sql &= "	   and dItm.t_item = cItm.t_item "
        Sql &= "left outer join tdmisg001" & comp & " as dm "
        Sql &= "        on dItm.t_docn = dm.t_docn "
        Sql &= "       and dItm.t_revn = dm.t_revn "
        Sql &= "left outer join ttpisg243" & comp & " as iRef "
        Sql &= "        on cDoc.t_pcod = iRef.t_cprd "
        Sql &= "	   and cDoc.t_iref = iRef.t_iref "
        Sql &= "	   and '' = isnull(iRef.t_sitm,'') "
        Sql &= "where cItm.t_orno='" & po.PONumber & "'"
        Sql &= "  and cItm.t_pono= " & poBom.ItemNo
        Sql &= "union all "
        Sql &= "select isnull(cDoc.t_iref,'') as t_iref, isnull(dItm.t_sitm,'') as t_sitm, isnull(iRef.t_sub2,'') as t_sub2, isnull(iRef.t_sub3,'') as t_sub3, isnull(iRef.t_sub4,'') as t_sub4, cItm.*, dm.t_dttl as DocumentDescription "
        Sql &= "from ttdisg002" & comp & " as cItm "
        Sql &= "left outer join tdmisg140" & comp & " as cDoc "
        Sql &= "        on cItm.t_docn = cDoc.t_docn "
        Sql &= "left outer join tdmisg002" & comp & " as dItm "
        Sql &= "        on dItm.t_docn = cItm.t_docn "
        Sql &= "       and dItm.t_revn = cItm.t_revi "
        Sql &= "	   and dItm.t_item = cItm.t_item "
        Sql &= "left outer join tdmisg001" & comp & " as dm "
        Sql &= "        on dItm.t_docn = dm.t_docn "
        Sql &= "       and dItm.t_revn = dm.t_revn "
        Sql &= "inner join ttpisg243" & comp & " as iRef "
        Sql &= "        on cDoc.t_pcod = iRef.t_cprd "
        Sql &= "	   and cDoc.t_iref = iRef.t_iref "
        Sql &= "	   and isnull(dItm.t_sitm,'') = isnull(iRef.t_sitm,'') "
        Sql &= "where cItm.t_orno='" & po.PONumber & "'"
        Sql &= "  and cItm.t_pono= " & poBom.ItemNo
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Sql
            Results = New List(Of SIS.PAK.pakERPPOLChild)
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            While (Reader.Read())
              Results.Add(New SIS.PAK.pakERPPOLChild(Reader))
            End While
            Reader.Close()
          End Using
        End Using
        Return Results
      End Function
      Public Shared Function IsChildInPO(PONumber As String, comp As String) As Boolean
        Dim mRet As Boolean = False
        Dim Sql As String = ""
        Sql &= "select isnull(count(*),0) as cnt "
        Sql &= "from ttdisg002" & Comp & "  "
        Sql &= "where t_orno='" & PONumber & "'"
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Sql
            Con.Open()
            Dim cnt As Integer = Cmd.ExecuteScalar
            If cnt > 0 Then
              mRet = True
            End If
          End Using
        End Using
        Return mRet
      End Function

      Public Shared Function GetFromERP(ByVal PONumber As String, Comp As String) As SIS.PAK.pakPO
        Dim Sql As String = ""
        '--t_work = 3 =>Completed
        '--t_hdst 10 =>Approved
        '--       30 =>Cancelled
        '--cdf_catg=1 Boughtout,2 Fabrication, 3 General
        Sql &= "select TOP 1 "
        Sql &= "  apo.t_orno as PONumber,"
        Sql &= "  apo.t_vrsn as PORevision,"
        Sql &= "  apo.t_apdt as PODate,"
        Sql &= "  lpo.t_cprj as ProjectID,"
        Sql &= "  hpo.t_otbp as SupplierID,"
        Sql &= "  hpo.t_otad as BuyFromAddressID,"
        Sql &= "  hpo.t_ccon as BuyerID,"
        Sql &= "  (select top 1 (case when xx.t_bptc='IN' then 0 else 1 end) as tmp from ttppdm740" & Comp & " as xx where xx.t_cprj = lpo.t_cprj ) As PortRequired,"
        Sql &= "  substring(apo.t_orno+'ZZ',2,1) as DivisionID,"
        Sql &= "  hpo.t_hdst as erpPOStatus, "
        Sql &= "  (select sum(case when t_wght=0 then t_qnty else t_wght end) from ttdisg002" & Comp & " where t_orno='" & PONumber & "') As POWeight "
        Sql &= "  FROM ttdmsl400" & Comp & " apo "
        Sql &= "  cross apply (select top 1 tmp.t_cprj from ttdpur401" & Comp & " tmp where tmp.t_orno=apo.t_orno   "
        Sql &= "              ) lpo "
        Sql &= "  inner join ttdpur400" & Comp & " as hpo on apo.t_orno = hpo.t_orno "
        Sql &= "  WHERE apo.t_work = 3 "
        Sql &= "  and apo.t_vrsn = (select max(tmp.t_vrsn) from ttdmsl400" & Comp & " tmp where tmp.t_orno=apo.t_orno) "
        If PONumber <> String.Empty Then
          Sql &= "  and apo.t_orno='" & PONumber & "'"
        End If
        Dim Results As SIS.PAK.pakPO = Nothing
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Sql
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            While (Reader.Read())
              Results = New SIS.PAK.pakPO(Reader)
            End While
            Reader.Close()
          End Using
        End Using
        Return Results
      End Function
      Sub New()
        'dummy
      End Sub
    End Class
    Public Class erpPAKElement
      Inherits SIS.PAK.pakElements
      Public Shared Function GetFromERP(ByVal ElementID As String, comp As String) As SIS.PAK.pakElements
        Dim Ret As SIS.PAK.pakElements = Nothing
        Dim Sql As String = ""
        Sql &= "select top 1  "
        Sql &= "  t_cspa as ElementID,  "
        Sql &= "  t_desc as Description "
        Sql &= "  from ttppdm090" & Comp & "  "
        Sql &= "  where t_cspa ='" & ElementID & "'"
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Sql
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            If (Reader.Read()) Then
              Ret = New SIS.PAK.pakElements(Reader)
            End If
            Reader.Close()
          End Using
        End Using
        Return Ret
      End Function
      Sub New(ByVal rd As SqlDataReader)
        MyBase.New(rd)
      End Sub
      Sub New()
        MyBase.New()
      End Sub
    End Class


    Public Class erpEnggGroup
      Public Property EngGroup As String = ""
      Public Property EmpID As String = ""
      Public Property EmpName As String = ""
      Public Property EMailID As String = ""
      Public Property ProjectID As String = ""
      Public Shared Function GetFromERP(ByVal eUnit As String, ByVal ProjectID As String, comp As String) As List(Of erpEnggGroup)
        Dim Ret As New List(Of erpEnggGroup)
        Dim Sql As String = ""
        Sql &= "  Select distinct en.t_engi as EngGroup,en.t_logn as EmpID,em.t_nama As EmpName,bp.t_mail As EMailID, en.t_cprj as ProjectID from tdmisg133" & Comp & " As en "
        Sql &= "  inner Join ttccom001" & Comp & " As em On en.t_logn = em.t_loco "
        Sql &= "  Left outer join tbpmdm001" & Comp & " as bp on em.t_emno=bp.t_emno "
        Sql &= "  where en.t_eunt ='" & eUnit & "'"
        Sql &= "  and en.t_cprj ='" & ProjectID & "'"
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = Sql
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            While (Reader.Read())
              Ret.Add(New erpEnggGroup(Reader))
            End While
            Reader.Close()
          End Using
        End Using
        Return Ret
      End Function
      Sub New(ByVal rd As SqlDataReader)
        EngGroup = rd("EngGroup")
        EmpID = rd("EmpID")
        EmpName = rd("EmpName")
        EMailID = IIf(Convert.IsDBNull(rd("EMailID")), "", rd("EMailID"))
        ProjectID = IIf(Convert.IsDBNull(rd("ProjectID")), "", rd("ProjectID"))
      End Sub
      Sub New()
        MyBase.New()
      End Sub
    End Class

  End Class

#End Region
End Namespace
Public Class TmtlDrawings
  Public Property DocumentID As String = ""
  Public Property RevisionNo As String = ""
End Class
Public Class EffectedPOByDrawings
  Public Property PONumber As String = ""
  Public Property POLine As Integer = 0
  Public Property Drawings As New List(Of TmtlDrawings)
  Public Shared Function GetPOList(Comp As String, t_tran As String) As List(Of EffectedPOByDrawings)
    Dim Sql As String = ""
    Sql &= " SELECT distinct a.t_orno,a.t_pono,a.t_docn,a.t_revi FROM ttdisg002" & Comp & " as a "
    Sql &= " INNER JOIN tdmisg132" & Comp & " as b "
    Sql &= " ON a.t_docn = b.t_docn AND a.t_revi = b.t_revn "
    Sql &= " WHERE b.t_tran='" & t_tran & "' "
    Dim mRet As New List(Of EffectedPOByDrawings)
    Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
      Using Cmd As SqlCommand = Con.CreateCommand()
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = Sql
        Con.Open()
        Dim Reader As SqlDataReader = Cmd.ExecuteReader()
        While (Reader.Read())
          Dim PO As String = Reader("t_orno")
          Dim LN As String = Reader("t_pono")
          Dim x As New TmtlDrawings
          x.DocumentID = Reader("t_docn")
          x.RevisionNo = Reader("t_revi")
          Dim Found As Boolean = False
          Dim f As EffectedPOByDrawings = Nothing
          For Each ret As EffectedPOByDrawings In mRet
            If ret.PONumber = PO AndAlso ret.POLine = LN Then
              Found = True
              f = ret
              Exit For
            End If
          Next
          If Found Then
            f.Drawings.Add(x)
          Else
            f = New EffectedPOByDrawings
            f.PONumber = PO
            f.POLine = LN
            f.Drawings.Add(x)
            mRet.Add(f)
          End If
        End While
        Reader.Close()
      End Using
    End Using
    Return mRet
  End Function

End Class

Public Class dmisg002
  Public Property t_docn As String = ""
  Public Property t_revn As String = ""
  Public Property t_srno As Integer = 0
  Public Property t_item As String = ""
  Public Property t_dsca As String = ""
  Public Property t_citg As String = ""
  Public Property t_qnty As Double = 0
  Public Property t_wght As Double = 0
  Public Property t_itmt As String = ""
  Public Property t_txta As Integer = 0
  Public Property t_txtb As Integer = 0
  Public Property t_cuni As String = ""
  Public Property t_oitm As String = ""
  Public Property t_elem As String = ""
  Public Property t_Refcntd As Integer = 0
  Public Property t_Refcntu As Integer = 0
  Public Property t_sitm As String = ""

  Public Shared Function GetItems(DocumentID As String, RevisionNo As String, comp As String) As List(Of dmisg002)
    Dim Ret As New List(Of dmisg002)
    Dim Sql As String = ""
    Sql &= "select a.*  "
    Sql &= "  from tdmisg002" & comp & "  "
    Sql &= "  where t_docn ='" & DocumentID & "'"
    Sql &= "    and t_revn ='" & RevisionNo & "'"
    Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
      Using Cmd As SqlCommand = Con.CreateCommand()
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = Sql
        Con.Open()
        Dim Reader As SqlDataReader = Cmd.ExecuteReader()
        Do While (Reader.Read())
          Ret.Add(New dmisg002(Reader))
        Loop
        Reader.Close()
      End Using
    End Using
    Return Ret
  End Function

  Sub New(ByVal rd As SqlDataReader)
    EDICommon.DBCommon.NewObj(Me, rd)
  End Sub
  Sub New()
    MyBase.New()
  End Sub

End Class