Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK
  <DataObject()> _
  Partial Public Class pakERPPOLChild
    Private Shared _RecordCount As Integer
    Public Property t_orno As String = ""
    Public Property t_vrsn As Int32 = 0
    Public Property t_pono As Int32 = 0
    Public Property t_item As String = ""
    Public Property t_desc As String = ""
    Public Property t_qnty As Double = 0
    Public Property t_quom As String = ""
    Public Property t_wght As Double = 0
    Public Property t_slct As Int32 = 0
    Public Property t_stat As Int32 = 0
    Public Property t_pric As Double = 0
    Public Property t_amnt As Double = 0
    Public Property t_qoor As Double = 0
    Public Property t_acht As Double = 0
    Public Property t_docn As String = ""
    Public Property t_revi As String = ""
    Public Property t_Refcntd As Int32 = 0
    Public Property t_Refcntu As Int32 = 0
    Public Property t_iref As String = ""
    Public Property t_sitm As String = ""
    Public Property t_sub2 As String = ""
    Public Property t_sub3 As String = ""
    Public Property t_sub4 As String = ""
    Public Property DocumentDescription As String = ""
    Public Shared Function GetPOBItems(ByVal ERPPOLChild As SIS.PAK.pakERPPOLChild, ByVal PO As SIS.PAK.pakPO, ByVal POBOM As SIS.PAK.pakPOBOM) As SIS.PAK.pakPOBItems
      Dim tmp As New SIS.PAK.pakPOBItems
      With tmp
        .ItemReference = ERPPOLChild.t_iref
        .SubItem = ERPPOLChild.t_sitm
        .SubItem2 = ERPPOLChild.t_sub2
        .SubItem3 = ERPPOLChild.t_sub3
        .SubItem4 = ERPPOLChild.t_sub4
        .SerialNo = PO.SerialNo
        .BOMNo = POBOM.BOMNo
        .ItemNo = 0 'Put Identity
        .ItemCode = ERPPOLChild.t_item.Trim
        .ItemDescription = ERPPOLChild.t_desc
        .SupplierItemCode = ""
        .DivisionID = PO.DivisionID
        .ElementID = POBOM.ElementID
        If ERPPOLChild.t_quom <> "" Then
          .UOMQuantity = SIS.PAK.pakUnits.pakUnitsGetByDescription(ERPPOLChild.t_quom).UnitID
        End If
        .Quantity = ERPPOLChild.t_qnty
        .UOMWeight = SIS.PAK.pakUnits.pakUnitsGetByDescription("kg").UnitID
        If ERPPOLChild.t_qnty > 0 Then
          .WeightPerUnit = ERPPOLChild.t_wght / ERPPOLChild.t_qnty
        Else
          .WeightPerUnit = ERPPOLChild.t_wght
        End If
        'Create or Get DocumentNO
        Dim tmpDoc As SIS.PAK.pakDocuments = SIS.PAK.pakDocuments.pakDocumentsSelectByDocRevID(ERPPOLChild.t_docn, ERPPOLChild.t_revi)
        If tmpDoc Is Nothing Then
          Dim tmpD = New SIS.PAK.pakDocuments
          With tmpD
            .DocumentID = ERPPOLChild.t_docn
            .DocumentRevision = ERPPOLChild.t_revi
            .Description = ERPPOLChild.DocumentDescription
            .ExternalDocument = False
            .DisskFile = ""
            .Filename = ""
          End With
          tmpD = SIS.PAK.pakDocuments.InsertData(tmpD)
          .DocumentNo = tmpD.DocumentNo
        Else
          With tmpDoc
            .DocumentID = ERPPOLChild.t_docn
            .DocumentRevision = ERPPOLChild.t_revi
            .Description = ERPPOLChild.DocumentDescription
            .ExternalDocument = False
            .DisskFile = ""
            .Filename = ""
          End With
          tmpDoc = SIS.PAK.pakDocuments.UpdateData(tmpDoc)
          .DocumentNo = tmpDoc.DocumentNo
        End If
        .ParentItemNo = POBOM.ItemNo
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
        .Root = False
        .Middle = False
        .Bottom = True
        .Free = False
        .ItemLevel = 1
        .Prefix = "»"
        .QuantityToPack = 0
        .TotalWeightToPack = 0
        .QuantityToDespatch = 0
        .TotalWeightToDespatch = 0
        .QuantityDespatched = 0
        .TotalWeightDespatched = 0
        .QuantityReceived = 0
        .TotalWeightReceived = 0
      End With
      Return tmp
    End Function

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
