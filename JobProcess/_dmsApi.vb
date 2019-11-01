Imports System
Imports System.Web
Imports System.Xml
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.DMS
  Public Class UI

    Public Class apiItem
      Public Property IsAdmin As Boolean = False
      Public Property ChildCount As Integer = 0


#Region "********Properties**********"
      Public Property ItemID As Int32 = 0
      Public Property InheritFromParent As Boolean = False
      Public Property UserID As String = ""
      Public Property Description As String = ""
      Public Property RevisionNo As String = ""
      Public Property ItemTypeID As Int32 = 0
      Public Property StatusID As Int32 = 0
      Public Property StatusID_Description As String = ""
      Public Property CreatedBy As String = ""
      Private _CreatedOn As String = ""
      Public Property MaintainAllLog As Boolean = False
      Public Property BackwardLinkedItemID As String = ""
      Public Property MaintainVersions As Boolean = False
      Public Property MaintainStatusLog As Boolean = False
      Public Property LinkedItemID As String = ""
      Public Property LinkedItemTypeID As String = ""
      Public Property BackwardLinkedItemTypeID As String = ""
      Public Property IsMultiBackward As Boolean = False
      Public Property IsgecVaultID As String = ""
      Public Property DeleteFile As Integer = 1
      Public Property CreateFile As Integer = 1
      Public Property BrowseList As Integer = 1
      Public Property GrantAuthorization As Integer = 1
      Public Property CreateFolder As Integer = 1
      Public Property Publish As Integer = 1
      Public Property DeleteFolder As Integer = 1
      Public Property RenameFolder As Integer = 1
      Public Property ShowInList As Integer = 1
      Public Property CompanyID As String = ""
      Public Property ChildItemID As String = ""
      Public Property DepartmentID As String = ""
      Public Property DivisionID As String = ""
      Public Property IsMultiParent As Boolean = False
      Public Property ConvertedStatusID As String = ""
      Public Property IsMultiChild As Boolean = False
      Public Property ParentItemID As String = ""
      Public Property ProjectID As String = ""
      Public Property ForwardLinkedItemTypeID As String = ""
      Public Property IsMultiForward As Boolean = False
      Public Property IsMultiLinked As Boolean = False
      Public Property ForwardLinkedItemID As String = ""
      Public Property KeyWords As String = ""
      Public Property WBSID As String = ""
      Public Property FullDescription As String = ""
      Public Property EMailID As String = ""
      Public Property SearchInParent As Boolean = False
      Public Property Approved As Boolean = False
      Public Property Rejected As Boolean = False
      Public Property ActionRemarks As String = ""
      Public Property ActionBy As String = ""
      Private _ActionOn As String = ""
      Public Property IsError As String = ""
      Public Property ErrorMessage As String = ""
      Public Property CreatedOn() As String
        Get
          If Not _CreatedOn = String.Empty Then
            Return Convert.ToDateTime(_CreatedOn).ToString("dd/MM/yyyy HH:mm")
          End If
          Return _CreatedOn
        End Get
        Set(ByVal value As String)
          _CreatedOn = value
        End Set
      End Property
      Public Property ActionOn() As String
        Get
          If Not _ActionOn = String.Empty Then
            Return Convert.ToDateTime(_ActionOn).ToString("dd/MM/yyyy")
          End If
          Return _ActionOn
        End Get
        Set(ByVal value As String)
          If Convert.IsDBNull(value) Then
            _ActionOn = ""
          Else
            _ActionOn = value
          End If
        End Set
      End Property

#End Region
#Region "********* INSERT / UPDATE ***********"
      Public Shared Function InsertData(ByVal Record As SIS.DMS.UI.apiItem) As SIS.DMS.UI.apiItem
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.StoredProcedure
            Cmd.CommandText = "spdmsItemsInsert"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@InheritFromParent", SqlDbType.Bit, 3, Record.InheritFromParent)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@UserID", SqlDbType.NVarChar, 9, IIf(Record.UserID = "", Convert.DBNull, Record.UserID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 251, Record.Description)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@RevisionNo", SqlDbType.NVarChar, 51, Record.RevisionNo)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemTypeID", SqlDbType.Int, 11, Record.ItemTypeID)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@StatusID", SqlDbType.Int, 11, Record.StatusID)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedBy", SqlDbType.NVarChar, 9, Record.CreatedBy)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedOn", SqlDbType.DateTime, 21, Record.CreatedOn)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@MaintainAllLog", SqlDbType.Bit, 3, Record.MaintainAllLog)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@BackwardLinkedItemID", SqlDbType.Int, 11, IIf(Record.BackwardLinkedItemID = "", Convert.DBNull, Record.BackwardLinkedItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@MaintainVersions", SqlDbType.Bit, 3, Record.MaintainVersions)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@MaintainStatusLog", SqlDbType.Bit, 3, Record.MaintainStatusLog)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@LinkedItemID", SqlDbType.Int, 11, IIf(Record.LinkedItemID = "", Convert.DBNull, Record.LinkedItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@LinkedItemTypeID", SqlDbType.Int, 11, IIf(Record.LinkedItemTypeID = "", Convert.DBNull, Record.LinkedItemTypeID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@BackwardLinkedItemTypeID", SqlDbType.Int, 11, IIf(Record.BackwardLinkedItemTypeID = "", Convert.DBNull, Record.BackwardLinkedItemTypeID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiBackward", SqlDbType.Bit, 3, Record.IsMultiBackward)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsgecVaultID", SqlDbType.NVarChar, 51, IIf(Record.IsgecVaultID = "", Convert.DBNull, Record.IsgecVaultID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DeleteFile", SqlDbType.Int, 11, Record.DeleteFile)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreateFile", SqlDbType.Int, 11, Record.CreateFile)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@BrowseList", SqlDbType.Int, 11, Record.BrowseList)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@GrantAuthorization", SqlDbType.Int, 11, Record.GrantAuthorization)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreateFolder", SqlDbType.Int, 11, Record.CreateFolder)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Publish", SqlDbType.Int, 11, Record.Publish)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DeleteFolder", SqlDbType.Int, 11, Record.DeleteFolder)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@RenameFolder", SqlDbType.Int, 11, Record.RenameFolder)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ShowInList", SqlDbType.Int, 11, Record.ShowInList)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CompanyID", SqlDbType.NVarChar, 7, IIf(Record.CompanyID = "", Convert.DBNull, Record.CompanyID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ChildItemID", SqlDbType.Int, 11, IIf(Record.ChildItemID = "", Convert.DBNull, Record.ChildItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DepartmentID", SqlDbType.NVarChar, 7, IIf(Record.DepartmentID = "", Convert.DBNull, Record.DepartmentID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DivisionID", SqlDbType.NVarChar, 7, IIf(Record.DivisionID = "", Convert.DBNull, Record.DivisionID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiParent", SqlDbType.Bit, 3, Record.IsMultiParent)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ConvertedStatusID", SqlDbType.Int, 11, IIf(Record.ConvertedStatusID = "", Convert.DBNull, Record.ConvertedStatusID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiChild", SqlDbType.Bit, 3, Record.IsMultiChild)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentItemID", SqlDbType.Int, 11, IIf(Record.ParentItemID = "", Convert.DBNull, Record.ParentItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ProjectID", SqlDbType.NVarChar, 7, IIf(Record.ProjectID = "", Convert.DBNull, Record.ProjectID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ForwardLinkedItemTypeID", SqlDbType.Int, 11, IIf(Record.ForwardLinkedItemTypeID = "", Convert.DBNull, Record.ForwardLinkedItemTypeID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiForward", SqlDbType.Bit, 3, Record.IsMultiForward)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiLinked", SqlDbType.Bit, 3, Record.IsMultiLinked)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ForwardLinkedItemID", SqlDbType.Int, 11, IIf(Record.ForwardLinkedItemID = "", Convert.DBNull, Record.ForwardLinkedItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWords", SqlDbType.NVarChar, 251, IIf(Record.KeyWords = "", Convert.DBNull, Record.KeyWords))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@WBSID", SqlDbType.NVarChar, 9, IIf(Record.WBSID = "", Convert.DBNull, Record.WBSID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@FullDescription", SqlDbType.NVarChar, 1001, IIf(Record.FullDescription = "", Convert.DBNull, Record.FullDescription))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@EMailID", SqlDbType.NVarChar, 251, IIf(Record.EMailID = "", Convert.DBNull, Record.EMailID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@SearchInParent", SqlDbType.Bit, 3, Record.SearchInParent)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Approved", SqlDbType.Bit, 3, Record.Approved)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Rejected", SqlDbType.Bit, 3, Record.Rejected)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ActionRemarks", SqlDbType.NVarChar, 251, IIf(Record.ActionRemarks = "", Convert.DBNull, Record.ActionRemarks))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ActionBy", SqlDbType.NVarChar, 9, IIf(Record.ActionBy = "", Convert.DBNull, Record.ActionBy))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ActionOn", SqlDbType.DateTime, 21, IIf(Record.ActionOn = "", Convert.DBNull, Record.ActionOn))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsError", SqlDbType.Bit, 3, IIf(Record.IsError = "", Convert.DBNull, Record.IsError))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ErrorMessage", SqlDbType.NVarChar, 501, IIf(Record.ErrorMessage = "", Convert.DBNull, Record.ErrorMessage))
            Cmd.Parameters.Add("@Return_ItemID", SqlDbType.Int, 11)
            Cmd.Parameters("@Return_ItemID").Direction = ParameterDirection.Output
            Con.Open()
            Cmd.ExecuteNonQuery()
            Record.ItemID = Cmd.Parameters("@Return_ItemID").Value
          End Using
        End Using
        Return Record
      End Function
      Public Shared Function UpdateData(ByVal Record As SIS.DMS.UI.apiItem) As SIS.DMS.UI.apiItem
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.StoredProcedure
            Cmd.CommandText = "spdmsItemsUpdate"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_ItemID", SqlDbType.Int, 11, Record.ItemID)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@InheritFromParent", SqlDbType.Bit, 3, Record.InheritFromParent)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@UserID", SqlDbType.NVarChar, 9, IIf(Record.UserID = "", Convert.DBNull, Record.UserID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 251, Record.Description)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@RevisionNo", SqlDbType.NVarChar, 51, Record.RevisionNo)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemTypeID", SqlDbType.Int, 11, Record.ItemTypeID)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@StatusID", SqlDbType.Int, 11, Record.StatusID)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedBy", SqlDbType.NVarChar, 9, Record.CreatedBy)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedOn", SqlDbType.DateTime, 21, Record.CreatedOn)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@MaintainAllLog", SqlDbType.Bit, 3, Record.MaintainAllLog)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@BackwardLinkedItemID", SqlDbType.Int, 11, IIf(Record.BackwardLinkedItemID = "", Convert.DBNull, Record.BackwardLinkedItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@MaintainVersions", SqlDbType.Bit, 3, Record.MaintainVersions)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@MaintainStatusLog", SqlDbType.Bit, 3, Record.MaintainStatusLog)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@LinkedItemID", SqlDbType.Int, 11, IIf(Record.LinkedItemID = "", Convert.DBNull, Record.LinkedItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@LinkedItemTypeID", SqlDbType.Int, 11, IIf(Record.LinkedItemTypeID = "", Convert.DBNull, Record.LinkedItemTypeID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@BackwardLinkedItemTypeID", SqlDbType.Int, 11, IIf(Record.BackwardLinkedItemTypeID = "", Convert.DBNull, Record.BackwardLinkedItemTypeID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiBackward", SqlDbType.Bit, 3, Record.IsMultiBackward)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsgecVaultID", SqlDbType.NVarChar, 51, IIf(Record.IsgecVaultID = "", Convert.DBNull, Record.IsgecVaultID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DeleteFile", SqlDbType.Int, 11, Record.DeleteFile)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreateFile", SqlDbType.Int, 11, Record.CreateFile)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@BrowseList", SqlDbType.Int, 11, Record.BrowseList)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@GrantAuthorization", SqlDbType.Int, 11, Record.GrantAuthorization)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreateFolder", SqlDbType.Int, 11, Record.CreateFolder)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Publish", SqlDbType.Int, 11, Record.Publish)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DeleteFolder", SqlDbType.Int, 11, Record.DeleteFolder)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@RenameFolder", SqlDbType.Int, 11, Record.RenameFolder)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ShowInList", SqlDbType.Int, 11, Record.ShowInList)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CompanyID", SqlDbType.NVarChar, 7, IIf(Record.CompanyID = "", Convert.DBNull, Record.CompanyID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ChildItemID", SqlDbType.Int, 11, IIf(Record.ChildItemID = "", Convert.DBNull, Record.ChildItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DepartmentID", SqlDbType.NVarChar, 7, IIf(Record.DepartmentID = "", Convert.DBNull, Record.DepartmentID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DivisionID", SqlDbType.NVarChar, 7, IIf(Record.DivisionID = "", Convert.DBNull, Record.DivisionID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiParent", SqlDbType.Bit, 3, Record.IsMultiParent)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ConvertedStatusID", SqlDbType.Int, 11, IIf(Record.ConvertedStatusID = "", Convert.DBNull, Record.ConvertedStatusID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiChild", SqlDbType.Bit, 3, Record.IsMultiChild)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentItemID", SqlDbType.Int, 11, IIf(Record.ParentItemID = "", Convert.DBNull, Record.ParentItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ProjectID", SqlDbType.NVarChar, 7, IIf(Record.ProjectID = "", Convert.DBNull, Record.ProjectID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ForwardLinkedItemTypeID", SqlDbType.Int, 11, IIf(Record.ForwardLinkedItemTypeID = "", Convert.DBNull, Record.ForwardLinkedItemTypeID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiForward", SqlDbType.Bit, 3, Record.IsMultiForward)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiLinked", SqlDbType.Bit, 3, Record.IsMultiLinked)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ForwardLinkedItemID", SqlDbType.Int, 11, IIf(Record.ForwardLinkedItemID = "", Convert.DBNull, Record.ForwardLinkedItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWords", SqlDbType.NVarChar, 251, IIf(Record.KeyWords = "", Convert.DBNull, Record.KeyWords))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@WBSID", SqlDbType.NVarChar, 9, IIf(Record.WBSID = "", Convert.DBNull, Record.WBSID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@FullDescription", SqlDbType.NVarChar, 1001, IIf(Record.FullDescription = "", Convert.DBNull, Record.FullDescription))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@EMailID", SqlDbType.NVarChar, 251, IIf(Record.EMailID = "", Convert.DBNull, Record.EMailID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@SearchInParent", SqlDbType.Bit, 3, Record.SearchInParent)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Approved", SqlDbType.Bit, 3, Record.Approved)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Rejected", SqlDbType.Bit, 3, Record.Rejected)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ActionRemarks", SqlDbType.NVarChar, 251, IIf(Record.ActionRemarks = "", Convert.DBNull, Record.ActionRemarks))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ActionBy", SqlDbType.NVarChar, 9, IIf(Record.ActionBy = "", Convert.DBNull, Record.ActionBy))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ActionOn", SqlDbType.DateTime, 21, IIf(Record.ActionOn = "", Convert.DBNull, Record.ActionOn))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsError", SqlDbType.Bit, 3, IIf(Record.IsError = "", Convert.DBNull, Record.IsError))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ErrorMessage", SqlDbType.NVarChar, 501, IIf(Record.ErrorMessage = "", Convert.DBNull, Record.ErrorMessage))
            Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
            Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
            Con.Open()
            Cmd.ExecuteNonQuery()
          End Using
        End Using
        Return Record
      End Function

#End Region
      Public Shared Function InsertHistory(ByVal Record As SIS.DMS.UI.apiItem) As SIS.DMS.UI.apiItem
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.StoredProcedure
            Cmd.CommandText = "spdmsHistoryInsert"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemID", SqlDbType.Int, 11, Record.ItemID)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@InheritFromParent", SqlDbType.Bit, 3, Record.InheritFromParent)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@UserID", SqlDbType.NVarChar, 9, IIf(Record.UserID = "", Convert.DBNull, Record.UserID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Description", SqlDbType.NVarChar, 251, Record.Description)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@RevisionNo", SqlDbType.NVarChar, 51, Record.RevisionNo)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemTypeID", SqlDbType.Int, 11, Record.ItemTypeID)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@StatusID", SqlDbType.Int, 11, Record.StatusID)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedBy", SqlDbType.NVarChar, 9, Record.CreatedBy)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedOn", SqlDbType.DateTime, 21, Record.CreatedOn)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@MaintainAllLog", SqlDbType.Bit, 3, Record.MaintainAllLog)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@BackwardLinkedItemID", SqlDbType.Int, 11, IIf(Record.BackwardLinkedItemID = "", Convert.DBNull, Record.BackwardLinkedItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@MaintainVersions", SqlDbType.Bit, 3, Record.MaintainVersions)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@MaintainStatusLog", SqlDbType.Bit, 3, Record.MaintainStatusLog)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@LinkedItemID", SqlDbType.Int, 11, IIf(Record.LinkedItemID = "", Convert.DBNull, Record.LinkedItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@LinkedItemTypeID", SqlDbType.Int, 11, IIf(Record.LinkedItemTypeID = "", Convert.DBNull, Record.LinkedItemTypeID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@BackwardLinkedItemTypeID", SqlDbType.Int, 11, IIf(Record.BackwardLinkedItemTypeID = "", Convert.DBNull, Record.BackwardLinkedItemTypeID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiBackward", SqlDbType.Bit, 3, Record.IsMultiBackward)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsgecVaultID", SqlDbType.NVarChar, 51, IIf(Record.IsgecVaultID = "", Convert.DBNull, Record.IsgecVaultID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DeleteFile", SqlDbType.Int, 11, Record.DeleteFile)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreateFile", SqlDbType.Int, 11, Record.CreateFile)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@BrowseList", SqlDbType.Int, 11, Record.BrowseList)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@GrantAuthorization", SqlDbType.Int, 11, Record.GrantAuthorization)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreateFolder", SqlDbType.Int, 11, Record.CreateFolder)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Publish", SqlDbType.Int, 11, Record.Publish)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DeleteFolder", SqlDbType.Int, 11, Record.DeleteFolder)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@RenameFolder", SqlDbType.Int, 11, Record.RenameFolder)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ShowInList", SqlDbType.Int, 11, Record.ShowInList)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CompanyID", SqlDbType.NVarChar, 7, IIf(Record.CompanyID = "", Convert.DBNull, Record.CompanyID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ChildItemID", SqlDbType.Int, 11, IIf(Record.ChildItemID = "", Convert.DBNull, Record.ChildItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DepartmentID", SqlDbType.NVarChar, 7, IIf(Record.DepartmentID = "", Convert.DBNull, Record.DepartmentID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@DivisionID", SqlDbType.NVarChar, 7, IIf(Record.DivisionID = "", Convert.DBNull, Record.DivisionID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiParent", SqlDbType.Bit, 3, Record.IsMultiParent)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ConvertedStatusID", SqlDbType.Int, 11, IIf(Record.ConvertedStatusID = "", Convert.DBNull, Record.ConvertedStatusID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiChild", SqlDbType.Bit, 3, Record.IsMultiChild)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ParentItemID", SqlDbType.Int, 11, IIf(Record.ParentItemID = "", Convert.DBNull, Record.ParentItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ProjectID", SqlDbType.NVarChar, 7, IIf(Record.ProjectID = "", Convert.DBNull, Record.ProjectID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ForwardLinkedItemTypeID", SqlDbType.Int, 11, IIf(Record.ForwardLinkedItemTypeID = "", Convert.DBNull, Record.ForwardLinkedItemTypeID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiForward", SqlDbType.Bit, 3, Record.IsMultiForward)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsMultiLinked", SqlDbType.Bit, 3, Record.IsMultiLinked)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ForwardLinkedItemID", SqlDbType.Int, 11, IIf(Record.ForwardLinkedItemID = "", Convert.DBNull, Record.ForwardLinkedItemID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWords", SqlDbType.NVarChar, 251, IIf(Record.KeyWords = "", Convert.DBNull, Record.KeyWords))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@WBSID", SqlDbType.NVarChar, 9, IIf(Record.WBSID = "", Convert.DBNull, Record.WBSID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@FullDescription", SqlDbType.NVarChar, 1001, IIf(Record.FullDescription = "", Convert.DBNull, Record.FullDescription))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@EMailID", SqlDbType.NVarChar, 251, IIf(Record.EMailID = "", Convert.DBNull, Record.EMailID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@SearchInParent", SqlDbType.Bit, 3, Record.SearchInParent)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Approved", SqlDbType.Bit, 3, Record.Approved)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Rejected", SqlDbType.Bit, 3, Record.Rejected)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ActionRemarks", SqlDbType.NVarChar, 251, IIf(Record.ActionRemarks = "", Convert.DBNull, Record.ActionRemarks))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ActionBy", SqlDbType.NVarChar, 9, IIf(Record.ActionBy = "", Convert.DBNull, Record.ActionBy))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ActionOn", SqlDbType.DateTime, 21, IIf(Record.ActionOn = "", Convert.DBNull, Record.ActionOn))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IsError", SqlDbType.Bit, 3, IIf(Record.IsError = "", Convert.DBNull, Record.IsError))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ErrorMessage", SqlDbType.NVarChar, 501, IIf(Record.ErrorMessage = "", Convert.DBNull, Record.ErrorMessage))
            Cmd.Parameters.Add("@Return_HistoryID", SqlDbType.Int, 11)
            Cmd.Parameters("@Return_HistoryID").Direction = ParameterDirection.Output
            Con.Open()
            Cmd.ExecuteNonQuery()
            Dim HistoryID As String = Cmd.Parameters("@Return_HistoryID").Value
          End Using
        End Using
        Return Nothing
      End Function

      Sub New(rd As SqlDataReader)
        SIS.DMS.UI.NewObj(Me, rd)
      End Sub
      Sub New()
        'dummy
      End Sub
    End Class

    Public Shared Function NewObj(this As Object, Reader As SqlDataReader) As Object
      Try
        For Each pi As System.Reflection.PropertyInfo In this.GetType.GetProperties
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
                      CallByName(this, pi.Name, CallType.Let, "0.00")
                    Case "bit"
                      CallByName(this, pi.Name, CallType.Let, Boolean.FalseString)
                    Case Else
                      CallByName(this, pi.Name, CallType.Let, String.Empty)
                  End Select
                Else
                  CallByName(this, pi.Name, CallType.Let, Reader(pi.Name))
                End If
              End If
            Catch ex As Exception
            End Try
          End If
        Next
      Catch ex As Exception
        Return Nothing
      End Try
      Return this
    End Function

    Public Shared Function GetItem(ByVal ItemID As String) As SIS.DMS.UI.apiItem
      Dim Results As apiItem = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Dim Sql As String = ""
        Sql &= "   select dms_items.*,  "
        Sql &= "   dms_states.Description as StatusID_Description "
        Sql &= "   from dms_items "
        Sql &= "   inner join dms_states on dms_states.statusid = dms_items.statusid "
        Sql &= " WHERE "
        Sql &= " [DMS_Items].[ItemID]=" & ItemID
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = Sql
          Con.Open()
          Dim rd As SqlDataReader = Cmd.ExecuteReader()
          If rd.Read() Then
            Results = New apiItem(rd)
          End If
          rd.Close()
        End Using
      End Using
      Return Results
    End Function

    Public Shared Function GetAssociatedWF(ItemID As String) As List(Of SIS.DMS.UI.apiItem)
      Dim tmp As New List(Of SIS.DMS.UI.apiItem)
      Dim mRet As New ArrayList
      mRet = getWorkFlow(ItemID, mRet)
      For Each str As String In mRet
        If str <> "" Then
          tmp.Add(SIS.DMS.UI.GetItem(str))
        End If
      Next
      Return tmp
    End Function
    Private Shared Function getWorkFlow(ItemID As Integer, mRet As ArrayList) As ArrayList
      Dim Parent As String = ""
      Dim Sql As String = ""
      Sql &= " select isnull(MultiItemID,'') as fld  from dms_multiItems where itemid=" & ItemID
      Sql &= " and multitypeID=" & 6
      Sql &= " and multiItemtypeID=" & 10
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Con.Open()
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = Sql
          Dim rd As SqlDataReader = Cmd.ExecuteReader()
          While rd.Read
            mRet.Add(rd("fld"))
          End While
          rd.Close()
        End Using
        Sql = " select isnull(ParentItemID,'') from dms_Items where itemid=" & ItemID
        Parent = ""
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = Sql
          Parent = Cmd.ExecuteScalar()
        End Using
      End Using
      If Parent <> "" Then
        mRet = getWorkFlow(Parent, mRet)
      End If
      Return mRet
    End Function

    Public Shared Function GetTopOneChild(itm As SIS.DMS.UI.apiItem) As SIS.DMS.UI.apiItem
      Dim mRet As SIS.DMS.UI.apiItem = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Con.Open()
        Dim Sql As String = " select top 1 * from dms_items "
        Sql &= "   WHERE "
        Sql &= "   ParentItemID=" & itm.ItemID
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = Sql
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            mRet = New SIS.DMS.UI.apiItem(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return mRet

    End Function

    Public Class ItemProperty
      Public Property ItemID As String = ""
      Public Property TransmittalID As String = ""
      Public Property ProjectID As String = ""
      Public Property ProjectName As String = ""
      Public Property TransmittalType As String = ""
      Public Property TransmittalSubject As String = ""
      Public Property Remarks As String = ""
      Public Property CreatedBy As String = ""
      Public Property CreatedName As String = ""
      Public Property CreatedOn As String = ""
      Public Property ApprovedBy As String = ""
      Public Property ApproverName As String = ""
      Public Property ApprovedOn As String = ""
      Public Property IssuedBy As String = ""
      Public Property IssuerName As String = ""
      Public Property IssuedOn As String = ""
      Public Shared Function InsertData(Record As SIS.DMS.UI.ItemProperty) As SIS.DMS.UI.ItemProperty
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.StoredProcedure
            Cmd.CommandText = "spDMS_ItemPropertyInsert"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ItemID", SqlDbType.Int, 11, Record.ItemID)
            EDICommon.DBCommon.AddDBParameter(Cmd, "@TransmittalID", SqlDbType.NVarChar, 51, IIf(Record.TransmittalID = "", Convert.DBNull, Record.TransmittalID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ProjectID", SqlDbType.NVarChar, 51, IIf(Record.ProjectID = "", Convert.DBNull, Record.ProjectID))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ProjectName", SqlDbType.NVarChar, 251, IIf(Record.ProjectName = "", Convert.DBNull, Record.ProjectName))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@TransmittalType", SqlDbType.NVarChar, 51, IIf(Record.TransmittalType = "", Convert.DBNull, Record.TransmittalType))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@TransmittalSubject", SqlDbType.NVarChar, 251, IIf(Record.TransmittalSubject = "", Convert.DBNull, Record.TransmittalSubject))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Remarks", SqlDbType.NVarChar, 501, IIf(Record.Remarks = "", Convert.DBNull, Record.Remarks))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedBy", SqlDbType.NVarChar, 51, IIf(Record.CreatedBy = "", Convert.DBNull, Record.CreatedBy))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedName", SqlDbType.NVarChar, 51, IIf(Record.CreatedName = "", Convert.DBNull, Record.CreatedName))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedOn", SqlDbType.DateTime, 21, IIf(Record.CreatedOn = "", Convert.DBNull, Record.CreatedOn))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ApprovedBy", SqlDbType.NVarChar, 51, IIf(Record.ApprovedBy = "", Convert.DBNull, Record.ApprovedBy))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ApproverName", SqlDbType.NVarChar, 51, IIf(Record.ApproverName = "", Convert.DBNull, Record.ApproverName))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@ApprovedOn", SqlDbType.DateTime, 21, IIf(Record.ApprovedOn = "", Convert.DBNull, Record.ApprovedOn))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IssuedBy", SqlDbType.NVarChar, 51, IIf(Record.IssuedBy = "", Convert.DBNull, Record.IssuedBy))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IssuerName", SqlDbType.NVarChar, 51, IIf(Record.IssuerName = "", Convert.DBNull, Record.IssuerName))
            EDICommon.DBCommon.AddDBParameter(Cmd, "@IssuedOn", SqlDbType.DateTime, 21, IIf(Record.IssuedOn = "", Convert.DBNull, Record.IssuedOn))
            Cmd.Parameters.Add("@Return_ItemID", SqlDbType.Int, 11)
            Cmd.Parameters("@Return_ItemID").Direction = ParameterDirection.Output
            Con.Open()
            Cmd.ExecuteNonQuery()
            Record.ItemID = Cmd.Parameters("@Return_ItemID").Value
          End Using
        End Using
        Return Record
      End Function

    End Class
  End Class

End Namespace
