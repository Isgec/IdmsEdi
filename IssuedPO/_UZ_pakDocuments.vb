Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.PAK
  Partial Public Class pakDocuments
    Public Shared Function UZ_pakDocumentsSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String) As List(Of SIS.PAK.pakDocuments)
      Dim Results As List(Of SIS.PAK.pakDocuments) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "sppak_LG_DocumentsSelectListSearch"
            Cmd.CommandText = "sppakDocumentsSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "sppak_LG_DocumentsSelectListFilteres"
            Cmd.CommandText = "sppakDocumentsSelectListFilteres"
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.PAK.pakDocuments)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakDocuments(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function UZ_pakDocumentsInsert(ByVal Record As SIS.PAK.pakDocuments) As SIS.PAK.pakDocuments
      Dim _Result As SIS.PAK.pakDocuments = pakDocumentsInsert(Record)
      Return _Result
    End Function
    Public Shared Function UZ_pakDocumentsUpdate(ByVal Record As SIS.PAK.pakDocuments) As SIS.PAK.pakDocuments
      Dim _Result As SIS.PAK.pakDocuments = pakDocumentsUpdate(Record)
      Return _Result
    End Function
    Public Shared Function UZ_pakDocumentsDelete(ByVal Record As SIS.PAK.pakDocuments) As Integer
      Dim _Result As Integer = pakDocumentsDelete(Record)
      Return _Result
    End Function
    Public Shared Function pakDocumentsSelectByDocRevID(ByVal DocumentID As String, ByVal DocumentRevision As String) As SIS.PAK.pakDocuments
      Dim Results As SIS.PAK.pakDocuments = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppak_LG_DocumentsSelectByDocRevID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentID", SqlDbType.NVarChar, DocumentID.Length, DocumentID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentRevision", SqlDbType.NVarChar, DocumentRevision.Length, DocumentRevision)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.PAK.pakDocuments(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function pakDocumentsSelectByDocID(ByVal DocumentID As String) As List(Of SIS.PAK.pakDocuments)
      Dim Results As List(Of SIS.PAK.pakDocuments) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppak_LG_DocumentsSelectByDocID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DocumentID", SqlDbType.NVarChar, DocumentID.Length, DocumentID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Results = New List(Of SIS.PAK.pakDocuments)
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.PAK.pakDocuments(Reader))
          End While
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
  End Class
End Namespace
