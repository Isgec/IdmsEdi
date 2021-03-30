Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK
  Partial Public Class pakPOBIDocuments
    Public Shared Function DeleteWF(ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemNo As Int32, ByVal DocNo As Int32) As SIS.PAK.pakPOBIDocuments
      Dim Results As SIS.PAK.pakPOBIDocuments = pakPOBIDocumentsGetByID(SerialNo, BOMNo, ItemNo, DocNo)
      SIS.PAK.pakPOBIDocuments.UZ_pakPOBIDocumentsDelete(Results)
      Return Results
    End Function
    Public Shared Function UZ_pakPOBIDocumentsSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal SerialNo As Int32, ByVal BOMNo As Int32, ByVal ItemNo As Int32) As List(Of SIS.PAK.pakPOBIDocuments)
      Dim Results As List(Of SIS.PAK.pakPOBIDocuments) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "sppak_LG_POBIDocumentsSelectListSearch"
            Cmd.CommandText = "sppakPOBIDocumentsSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "sppak_LG_POBIDocumentsSelectListFilteres"
            Cmd.CommandText = "sppakPOBIDocumentsSelectListFilteres"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_SerialNo", SqlDbType.Int, 10, IIf(SerialNo = Nothing, 0, SerialNo))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_BOMNo", SqlDbType.Int, 10, IIf(BOMNo = Nothing, 0, BOMNo))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_ItemNo", SqlDbType.Int, 10, IIf(ItemNo = Nothing, 0, ItemNo))
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakPOBIDocuments)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakPOBIDocuments(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function UZ_pakPOBIDocumentsInsert(ByVal Record As SIS.PAK.pakPOBIDocuments) As SIS.PAK.pakPOBIDocuments
      Dim _Result As SIS.PAK.pakPOBIDocuments = pakPOBIDocumentsInsert(Record)
      Return _Result
    End Function
    Public Shared Function UZ_pakPOBIDocumentsUpdate(ByVal Record As SIS.PAK.pakPOBIDocuments) As SIS.PAK.pakPOBIDocuments
      Dim _Result As SIS.PAK.pakPOBIDocuments = pakPOBIDocumentsUpdate(Record)
      Return _Result
    End Function
    Public Shared Function UZ_pakPOBIDocumentsDelete(ByVal Record As SIS.PAK.pakPOBIDocuments) As Integer
      Dim _Result As Integer = pakPOBIDocumentsDelete(Record)
      Return _Result
    End Function
  End Class
End Namespace
