Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Text
Imports System.Xml
Namespace SIS.EDI
  <DataObject()>
  Partial Public Class ediTmtlH
    Private Shared _RecordCount As Integer
    Private _t_tran As String = ""
    Private _t_type As Int32 = 0
    Private _t_bpid As String = ""
    Private _t_cadr As String = ""
    Private _t_cprj As String = ""
    Private _t_logn As String = ""
    Private _t_subj As String = ""
    Private _t_remk As String = ""
    Private _t_issu As String = ""
    Private _t_stat As Int32 = 0
    Private _t_ofbp As String = ""
    Private _t_vadr As String = ""
    Private _t_padr As String = ""
    Private _t_dprj As String = ""
    Private _t_user As String = ""
    Private _t_date As String = ""
    Private _t_subt As Int32 = 0
    Private _t_refr As String = ""
    Private _t_appr As Int32 = 0
    Private _t_rejc As Int32 = 0
    Private _t_rekm As String = ""
    Private _t_apdt As String = ""
    Private _t_apsu As String = ""
    Private _t_irmk As String = ""
    Private _t_iisu As Int32 = 0
    Private _t_retn As Int32 = 0
    Private _t_smdt As String = ""
    Private _t_isby As String = ""
    Private _t_isdt As String = ""
    Private _t_Refcntd As Int32 = 0
    Private _t_Refcntu As Int32 = 0
    Private _t_edif As Int32 = 0
    Public Property t_tran() As String
      Get
        Return _t_tran
      End Get
      Set(ByVal value As String)
        _t_tran = value
      End Set
    End Property
    Public Property t_type() As Int32
      Get
        Return _t_type
      End Get
      Set(ByVal value As Int32)
        _t_type = value
      End Set
    End Property
    Public Property t_bpid() As String
      Get
        Return _t_bpid
      End Get
      Set(ByVal value As String)
        _t_bpid = value
      End Set
    End Property
    Public Property t_cadr() As String
      Get
        Return _t_cadr
      End Get
      Set(ByVal value As String)
        _t_cadr = value
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
    Public Property t_logn() As String
      Get
        Return _t_logn
      End Get
      Set(ByVal value As String)
        _t_logn = value
      End Set
    End Property
    Public Property t_subj() As String
      Get
        Return _t_subj
      End Get
      Set(ByVal value As String)
        _t_subj = value
      End Set
    End Property
    Public Property t_remk() As String
      Get
        Return _t_remk
      End Get
      Set(ByVal value As String)
        _t_remk = value
      End Set
    End Property
    Public Property t_issu() As String
      Get
        Return _t_issu
      End Get
      Set(ByVal value As String)
        _t_issu = value
      End Set
    End Property
    Public Property t_stat() As Int32
      Get
        Return _t_stat
      End Get
      Set(ByVal value As Int32)
        _t_stat = value
      End Set
    End Property
    Public Property t_ofbp() As String
      Get
        Return _t_ofbp
      End Get
      Set(ByVal value As String)
        _t_ofbp = value
      End Set
    End Property
    Public Property t_vadr() As String
      Get
        Return _t_vadr
      End Get
      Set(ByVal value As String)
        _t_vadr = value
      End Set
    End Property
    Public Property t_padr() As String
      Get
        Return _t_padr
      End Get
      Set(ByVal value As String)
        _t_padr = value
      End Set
    End Property
    Public Property t_dprj() As String
      Get
        Return _t_dprj
      End Get
      Set(ByVal value As String)
        _t_dprj = value
      End Set
    End Property
    Public Property t_user() As String
      Get
        Return _t_user
      End Get
      Set(ByVal value As String)
        _t_user = value
      End Set
    End Property
    Public Property t_date() As String
      Get
        If Not _t_date = String.Empty Then
          Return Convert.ToDateTime(_t_date).ToString("dd/MM/yyyy")
        End If
        Return _t_date
      End Get
      Set(ByVal value As String)
        _t_date = value
      End Set
    End Property
    Public Property t_subt() As Int32
      Get
        Return _t_subt
      End Get
      Set(ByVal value As Int32)
        _t_subt = value
      End Set
    End Property
    Public Property t_refr() As String
      Get
        Return _t_refr
      End Get
      Set(ByVal value As String)
        _t_refr = value
      End Set
    End Property
    Public Property t_appr() As Int32
      Get
        Return _t_appr
      End Get
      Set(ByVal value As Int32)
        _t_appr = value
      End Set
    End Property
    Public Property t_rejc() As Int32
      Get
        Return _t_rejc
      End Get
      Set(ByVal value As Int32)
        _t_rejc = value
      End Set
    End Property
    Public Property t_rekm() As String
      Get
        Return _t_rekm
      End Get
      Set(ByVal value As String)
        _t_rekm = value
      End Set
    End Property
    Public Property t_apdt() As String
      Get
        If Not _t_apdt = String.Empty Then
          Return Convert.ToDateTime(_t_apdt).ToString("dd/MM/yyyy")
        End If
        Return _t_apdt
      End Get
      Set(ByVal value As String)
        _t_apdt = value
      End Set
    End Property
    Public Property t_apsu() As String
      Get
        Return _t_apsu
      End Get
      Set(ByVal value As String)
        _t_apsu = value
      End Set
    End Property
    Public Property t_irmk() As String
      Get
        Return _t_irmk
      End Get
      Set(ByVal value As String)
        _t_irmk = value
      End Set
    End Property
    Public Property t_iisu() As Int32
      Get
        Return _t_iisu
      End Get
      Set(ByVal value As Int32)
        _t_iisu = value
      End Set
    End Property
    Public Property t_retn() As Int32
      Get
        Return _t_retn
      End Get
      Set(ByVal value As Int32)
        _t_retn = value
      End Set
    End Property
    Public Property t_smdt() As String
      Get
        If Not _t_smdt = String.Empty Then
          Return Convert.ToDateTime(_t_smdt).ToString("dd/MM/yyyy")
        End If
        Return _t_smdt
      End Get
      Set(ByVal value As String)
        _t_smdt = value
      End Set
    End Property
    Public Property t_isby() As String
      Get
        Return _t_isby
      End Get
      Set(ByVal value As String)
        _t_isby = value
      End Set
    End Property
    Public Property t_isdt() As String
      Get
        If Not _t_isdt = String.Empty Then
          Return Convert.ToDateTime(_t_isdt).ToString("dd/MM/yyyy")
        End If
        Return _t_isdt
      End Get
      Set(ByVal value As String)
        _t_isdt = value
      End Set
    End Property
    Public Property t_Refcntd() As Int32
      Get
        Return _t_Refcntd
      End Get
      Set(ByVal value As Int32)
        _t_Refcntd = value
      End Set
    End Property
    Public Property t_Refcntu() As Int32
      Get
        Return _t_Refcntu
      End Get
      Set(ByVal value As Int32)
        _t_Refcntu = value
      End Set
    End Property
    Public Property t_edif() As Int32
      Get
        Return _t_edif
      End Get
      Set(ByVal value As Int32)
        _t_edif = value
      End Set
    End Property
    Public ReadOnly Property DisplayField() As String
      Get
        Return ""
      End Get
    End Property
    Public ReadOnly Property PrimaryKey() As String
      Get
        Return _t_tran
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
    Public Class PKediTmtlH
      Private _t_tran As String = ""
      Public Property t_tran() As String
        Get
          Return _t_tran
        End Get
        Set(ByVal value As String)
          _t_tran = value
        End Set
      End Property
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediTmtlHGetNewRecord() As SIS.EDI.ediTmtlH
      Return New SIS.EDI.ediTmtlH()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediTmtlHGetByID(ByVal t_tran As String) As SIS.EDI.ediTmtlH
      Dim Results As SIS.EDI.ediTmtlH = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediTmtlHSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_tran", SqlDbType.VarChar, t_tran.ToString.Length, t_tran)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "0340")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.EDI.ediTmtlH(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediTmtlHSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal t_tran As String) As List(Of SIS.EDI.ediTmtlH)
      Dim Results As List(Of SIS.EDI.ediTmtlH) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "spediTmtlHSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "spediTmtlHSelectListFilteres"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@Filter_t_tran", SqlDbType.VarChar, 9, IIf(t_tran Is Nothing, String.Empty, t_tran))
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "0340")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.EDI.ediTmtlH)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.EDI.ediTmtlH(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function ediTmtlHSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String, ByVal t_tran As String) As Integer
      Return _RecordCount
    End Function
    'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediTmtlHGetByID(ByVal t_tran As String, ByVal Filter_t_tran As String) As SIS.EDI.ediTmtlH
      Return ediTmtlHGetByID(t_tran)
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)>
    Public Shared Function ediTmtlHUpdate(ByVal Record As SIS.EDI.ediTmtlH) As SIS.EDI.ediTmtlH
      Dim _Rec As SIS.EDI.ediTmtlH = SIS.EDI.ediTmtlH.ediTmtlHGetByID(Record.t_tran)
      With _Rec
        .t_type = Record.t_type
        .t_bpid = Record.t_bpid
        .t_cadr = Record.t_cadr
        .t_cprj = Record.t_cprj
        .t_logn = Record.t_logn
        .t_subj = Record.t_subj
        .t_remk = Record.t_remk
        .t_issu = Record.t_issu
        .t_stat = Record.t_stat
        .t_ofbp = Record.t_ofbp
        .t_vadr = Record.t_vadr
        .t_padr = Record.t_padr
        .t_dprj = Record.t_dprj
        .t_user = Record.t_user
        .t_date = Record.t_date
        .t_subt = Record.t_subt
        .t_refr = Record.t_refr
        .t_appr = Record.t_appr
        .t_rejc = Record.t_rejc
        .t_rekm = Record.t_rekm
        .t_apdt = Record.t_apdt
        .t_apsu = Record.t_apsu
        .t_irmk = Record.t_irmk
        .t_iisu = Record.t_iisu
        .t_retn = Record.t_retn
        .t_smdt = Record.t_smdt
        .t_isby = Record.t_isby
        .t_isdt = Record.t_isdt
        .t_Refcntd = Record.t_Refcntd
        .t_Refcntu = Record.t_Refcntu
        .t_edif = Record.t_edif
      End With
      Return SIS.EDI.ediTmtlH.UpdateData(_Rec)
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.EDI.ediTmtlH) As SIS.EDI.ediTmtlH
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediTmtlHUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_t_tran", SqlDbType.VarChar, 10, Record.t_tran)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_tran", SqlDbType.VarChar, 10, Record.t_tran)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_type", SqlDbType.Int, 11, Record.t_type)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_bpid", SqlDbType.VarChar, 10, Record.t_bpid)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_cadr", SqlDbType.VarChar, 10, Record.t_cadr)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_cprj", SqlDbType.VarChar, 10, Record.t_cprj)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_logn", SqlDbType.VarChar, 17, Record.t_logn)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_subj", SqlDbType.VarChar, 101, Record.t_subj)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_remk", SqlDbType.VarChar, 101, Record.t_remk)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_issu", SqlDbType.VarChar, 4, Record.t_issu)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_stat", SqlDbType.Int, 11, Record.t_stat)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_ofbp", SqlDbType.VarChar, 10, Record.t_ofbp)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_vadr", SqlDbType.VarChar, 10, Record.t_vadr)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_padr", SqlDbType.VarChar, 10, Record.t_padr)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_dprj", SqlDbType.VarChar, 10, Record.t_dprj)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_user", SqlDbType.VarChar, 17, Record.t_user)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_date", SqlDbType.DateTime, 21, Record.t_date)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_subt", SqlDbType.Int, 11, Record.t_subt)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_refr", SqlDbType.VarChar, 33, Record.t_refr)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_appr", SqlDbType.Int, 11, Record.t_appr)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_rejc", SqlDbType.Int, 11, Record.t_rejc)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_rekm", SqlDbType.VarChar, 101, Record.t_rekm)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_apdt", SqlDbType.DateTime, 21, Record.t_apdt)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_apsu", SqlDbType.VarChar, 17, Record.t_apsu)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_irmk", SqlDbType.VarChar, 101, Record.t_irmk)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_iisu", SqlDbType.Int, 11, Record.t_iisu)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_retn", SqlDbType.Int, 11, Record.t_retn)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_smdt", SqlDbType.DateTime, 21, Record.t_smdt)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_isby", SqlDbType.VarChar, 17, Record.t_isby)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_isdt", SqlDbType.DateTime, 21, Record.t_isdt)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_Refcntd", SqlDbType.Int, 11, Record.t_Refcntd)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_Refcntu", SqlDbType.Int, 11, Record.t_Refcntu)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_edif", SqlDbType.Int, 11, Record.t_edif)
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
    Public Shared Function GetHTML(ByVal TransmittalID As String) As String
      Dim form1 As New WebControls.Panel
      Dim t_tran As String = TransmittalID
      Dim oVar As SIS.EDI.ediTmtlH = SIS.EDI.ediTmtlH.ediTmtlHGetByID(t_tran)
      Dim oTblediWTmtlH As New Table
      oTblediWTmtlH.Width = 1000
      oTblediWTmtlH.GridLines = GridLines.Both
      oTblediWTmtlH.Style.Add("margin-top", "15px")
      oTblediWTmtlH.Style.Add("margin-left", "10px")
      Dim oColediWTmtlH As TableCell = Nothing
      Dim oRowediWTmtlH As TableRow = Nothing
      oRowediWTmtlH = New TableRow
      oRowediWTmtlH.BackColor = System.Drawing.Color.CadetBlue
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Transmittal No."
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = oVar.t_tran
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Reference No."
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = oVar.t_refr
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oTblediWTmtlH.Rows.Add(oRowediWTmtlH)
      oRowediWTmtlH = New TableRow
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Tmtl. Type"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      Dim X As String = "CUSTOMER"
      Select Case oVar.t_type
        Case 1
          X = "CUSTOMER"
        Case 2
          X = "INTERNAL"
        Case 3
          X = "SITE"
        Case 4
          X = "VENDOR"
      End Select
      oColediWTmtlH.Text = X
      oColediWTmtlH.Style.Add("text-align", "center")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Tmtl. Project"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      If oVar.t_cprj <> "" Then
        oColediWTmtlH.Text = "[" & oVar.t_cprj & "] " & emp.GetProjectName(oVar.t_cprj)
      End If
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oTblediWTmtlH.Rows.Add(oRowediWTmtlH)
      oRowediWTmtlH = New TableRow
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Customer"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      If oVar.t_bpid <> "" Then
        oColediWTmtlH.Text = "[" & oVar.t_bpid & "] " & emp.GetBPName(oVar.t_bpid)
      End If
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Vendor"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      If oVar.t_ofbp <> "" Then
        oColediWTmtlH.Text = "[" & oVar.t_ofbp & "] " & emp.GetBPName(oVar.t_ofbp)
      End If
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oTblediWTmtlH.Rows.Add(oRowediWTmtlH)
      oRowediWTmtlH = New TableRow
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Employee"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      If oVar.t_logn <> "" Then
        Try
          oColediWTmtlH.Text = "[" & oVar.t_logn & "] " & emp.GetEmp(oVar.t_logn).empName
        Catch ex As Exception
        End Try
      End If
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Project"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      If oVar.t_dprj <> "" Then
        oColediWTmtlH.Text = "[" & oVar.t_dprj & "] " & emp.GetProjectName(oVar.t_dprj)
      End If
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oTblediWTmtlH.Rows.Add(oRowediWTmtlH)
      oRowediWTmtlH = New TableRow
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Issued Via"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = emp.GetIssuedVia(oVar.t_issu)
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "5"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oTblediWTmtlH.Rows.Add(oRowediWTmtlH)
      oRowediWTmtlH = New TableRow
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Subject"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = oVar.t_subj
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "5"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oTblediWTmtlH.Rows.Add(oRowediWTmtlH)
      oRowediWTmtlH = New TableRow
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Remarks"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = oVar.t_remk
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "5"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oTblediWTmtlH.Rows.Add(oRowediWTmtlH)
      oRowediWTmtlH = New TableRow
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Created By"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      If oVar.t_user <> "" Then
        Try
          oColediWTmtlH.Text = "[" & oVar.t_user & "] " & emp.GetEmp(oVar.t_user).empName
        Catch ex As Exception
        End Try
      End If
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Created On"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = oVar.t_date
      oColediWTmtlH.Style.Add("text-align", "center")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oTblediWTmtlH.Rows.Add(oRowediWTmtlH)
      oRowediWTmtlH = New TableRow
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Approved By"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      If oVar.t_apsu <> "" Then
        Try
          oColediWTmtlH.Text = "[" & oVar.t_apsu & "] " & emp.GetEmp(oVar.t_apsu).empName
        Catch ex As Exception
        End Try
      End If
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Approved On"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = oVar.t_apdt
      oColediWTmtlH.Style.Add("text-align", "center")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oTblediWTmtlH.Rows.Add(oRowediWTmtlH)
      oRowediWTmtlH = New TableRow
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Issued By"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      If oVar.t_isby <> "" Then
        Try
          oColediWTmtlH.Text = "[" & oVar.t_isby & "] " & emp.GetEmp(oVar.t_isby).empName
        Catch ex As Exception
        End Try
      End If
      oColediWTmtlH.Style.Add("text-align", "left")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = "Issued On"
      oColediWTmtlH.Font.Bold = True
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oColediWTmtlH = New TableCell
      oColediWTmtlH.Text = oVar.t_isdt
      oColediWTmtlH.Style.Add("text-align", "center")
      oColediWTmtlH.ColumnSpan = "2"
      oRowediWTmtlH.Cells.Add(oColediWTmtlH)
      oTblediWTmtlH.Rows.Add(oRowediWTmtlH)
      form1.Controls.Add(oTblediWTmtlH)
      Dim oTblediWTmtlD As Table = Nothing
      Dim oRowediWTmtlD As TableRow = Nothing
      Dim oColediWTmtlD As TableCell = Nothing
      Dim oediWTmtlDs As List(Of SIS.EDI.ediTmtlD) = SIS.EDI.ediTmtlD.ediTmtlDSelectList(0, 9999, "", False, "", oVar.t_tran, "", "")
      If oediWTmtlDs.Count > 0 Then
        Dim lbl As New System.Web.UI.WebControls.Label
        With lbl
          .Font.Bold = True
          .Font.Size = FontUnit.Point(14)
          .Text = "Standard Document in this transmittal will NOT be available for download. Please download/get it from Vault/Design department."
          .Style.Add("margin-top", "15px")
          .Style.Add("margin-left", "10px")
        End With
        form1.Controls.Add(lbl)
        Dim oTblhediWTmtlD As Table = New Table
        oTblhediWTmtlD.Width = 1000
        oTblhediWTmtlD.Style.Add("margin-top", "15px")
        oTblhediWTmtlD.Style.Add("margin-left", "10px")
        oTblhediWTmtlD.Style.Add("margin-right", "10px")
        Dim oRowhediWTmtlD As TableRow = New TableRow
        Dim oColhediWTmtlD As TableCell = New TableCell
        oColhediWTmtlD.Font.Bold = True
        oColhediWTmtlD.Font.Underline = True
        oColhediWTmtlD.Font.Size = 10
        oColhediWTmtlD.CssClass = "grpHD"
        oColhediWTmtlD.BackColor = System.Drawing.Color.CadetBlue
        oColhediWTmtlD.Text = "Transmittal Detail"
        oRowhediWTmtlD.Cells.Add(oColhediWTmtlD)
        oTblhediWTmtlD.Rows.Add(oRowhediWTmtlD)
        form1.Controls.Add(oTblhediWTmtlD)
        oTblediWTmtlD = New Table
        oTblediWTmtlD.Width = 1000
        oTblediWTmtlD.GridLines = GridLines.Both
        oTblediWTmtlD.Style.Add("margin-left", "10px")
        oTblediWTmtlD.Style.Add("margin-right", "10px")
        oRowediWTmtlD = New TableRow
        oRowediWTmtlD.BackColor = System.Drawing.Color.CornflowerBlue
        'oColediWTmtlD = New TableCell
        'oColediWTmtlD.Text = "Transmittal No."
        'oColediWTmtlD.Font.Bold = True
        'oColediWTmtlD.CssClass = "colHD"
        'oColediWTmtlD.Style.Add("text-align", "left")
        'oRowediWTmtlD.Cells.Add(oColediWTmtlD)
        oColediWTmtlD = New TableCell
        oColediWTmtlD.Text = "Document ID"
        oColediWTmtlD.Font.Bold = True
        oColediWTmtlD.CssClass = "colHD"
        oColediWTmtlD.Style.Add("text-align", "left")
        oRowediWTmtlD.Cells.Add(oColediWTmtlD)
        oColediWTmtlD = New TableCell
        oColediWTmtlD.Text = "Revision No."
        oColediWTmtlD.Font.Bold = True
        oColediWTmtlD.CssClass = "colHD"
        oColediWTmtlD.Style.Add("text-align", "left")
        oRowediWTmtlD.Cells.Add(oColediWTmtlD)
        oColediWTmtlD = New TableCell
        oColediWTmtlD.Text = "Title"
        oColediWTmtlD.Font.Bold = True
        oColediWTmtlD.CssClass = "colHD"
        oColediWTmtlD.Style.Add("text-align", "left")
        oRowediWTmtlD.Cells.Add(oColediWTmtlD)
        oColediWTmtlD = New TableCell
        oColediWTmtlD.Text = "Remarks"
        oColediWTmtlD.Font.Bold = True
        oColediWTmtlD.CssClass = "colHD"
        oColediWTmtlD.Style.Add("text-align", "left")
        oRowediWTmtlD.Cells.Add(oColediWTmtlD)
        oTblediWTmtlD.Rows.Add(oRowediWTmtlD)
        For Each oediWTmtlD As SIS.EDI.ediTmtlD In oediWTmtlDs
          oRowediWTmtlD = New TableRow
          'oColediWTmtlD = New TableCell
          'oColediWTmtlD.CssClass = "rowHD"
          'oColediWTmtlD.Text = oediWTmtlD.t_tran
          'oColediWTmtlD.Style.Add("text-align", "left")
          'oRowediWTmtlD.Cells.Add(oColediWTmtlD)
          oColediWTmtlD = New TableCell
          oColediWTmtlD.CssClass = "rowHD"
          oColediWTmtlD.Text = oediWTmtlD.t_docn
          oColediWTmtlD.Style.Add("text-align", "left")
          oRowediWTmtlD.Cells.Add(oColediWTmtlD)
          oColediWTmtlD = New TableCell
          oColediWTmtlD.CssClass = "rowHD"
          oColediWTmtlD.Text = oediWTmtlD.t_revn
          oColediWTmtlD.Style.Add("text-align", "left")
          oRowediWTmtlD.Cells.Add(oColediWTmtlD)
          oColediWTmtlD = New TableCell
          oColediWTmtlD.CssClass = "rowHD"
          oColediWTmtlD.Text = oediWTmtlD.t_dttl
          oColediWTmtlD.Style.Add("text-align", "left")
          oRowediWTmtlD.Cells.Add(oColediWTmtlD)
          oColediWTmtlD = New TableCell
          oColediWTmtlD.CssClass = "rowHD"
          oColediWTmtlD.Text = oediWTmtlD.t_remk
          oColediWTmtlD.Style.Add("text-align", "left")
          oRowediWTmtlD.Cells.Add(oColediWTmtlD)
          oTblediWTmtlD.Rows.Add(oRowediWTmtlD)
        Next
        form1.Controls.Add(oTblediWTmtlD)
      End If
      Dim sb As StringBuilder = New StringBuilder()
      Dim sw As IO.StringWriter = New IO.StringWriter(sb)
      Dim writer As HtmlTextWriter = New HtmlTextWriter(sw)
      Try
        form1.RenderControl(writer)
      Catch ex As Exception

      End Try
      Return sb.ToString
    End Function

    Private Class emp
      Public Property empID As Integer = 0
      Public Property empName As String = ""
      Public Property empEMail As String = ""

      Public Shared Function GetEmp(ByVal empID As Integer) As emp
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " emp1.t_nama as empName,"
        mSql = mSql & " bpe1.t_mail as empEMail "
        mSql = mSql & " from ttccom001200 as emp1 "
        mSql = mSql & " left outer join tbpmdm001200 as bpe1 on emp1.t_emno=bpe1.t_emno "
        mSql = mSql & " where emp1.t_emno = '" & empID & "'"
        Dim tmp As emp = Nothing
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            tmp = New emp
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            If (Reader.Read()) Then
              With tmp
                .empID = empID
                If Not Convert.IsDBNull(Reader("empName")) Then .empName = Reader("empName")
                If Not Convert.IsDBNull(Reader("empEMail")) Then .empEMail = Reader("empEMail")
              End With
            End If
            Reader.Close()
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetProjectName(ByVal ProjectID As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " t_dsca "
        mSql = mSql & " from ttcmcs052200 "
        mSql = mSql & " where t_cprj = '" & ProjectID & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
          End Using
        End Using
        Return tmp
      End Function

      Public Shared Function GetBPName(ByVal BPID As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " t_nama "
        mSql = mSql & " from ttccom100200 "
        mSql = mSql & " where t_bpid = '" & BPID & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetIssuedVia(ByVal IssueID As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " t_dsca "
        mSql = mSql & " from tdmisg125200 "
        mSql = mSql & " where t_issu = '" & IssueID & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
          End Using
        End Using
        Return tmp
      End Function

      Public Shared Function GetReceiptCreator(ByVal tmtlID As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " t_user as [user] "
        mSql = mSql & " from tdmisg134200 "
        mSql = mSql & " where t_trno = '" & tmtlID & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            If (Reader.Read()) Then
              If Not Convert.IsDBNull(Reader("user")) Then tmp = Reader("user")
            End If
            Reader.Close()
          End Using
        End Using
        Return tmp
      End Function
    End Class

  End Class

End Namespace
