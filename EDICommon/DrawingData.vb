Imports System.Xml
Imports System.Xml.Serialization
Imports EDICommon

<Serializable>
Public Class DrawingData
  Implements IDisposable
  Implements ICloneable
  Public Function Clone() As Object Implements ICloneable.Clone
    Dim tmp As DrawingData = MyBase.MemberwiseClone()
    With tmp
      tmp.Items = New List(Of ItemData)
      tmp.AllParts = New List(Of PartData)
      tmp.AllColumns = New List(Of Columns)
    End With
    Return tmp
  End Function

  Public Property SourceDocument As String = ""

  Public Property DrawingNumber As String = ""
  Private _DrawingID As String = ""
  Public Property DrawingID As String
    Get
      If _DrawingID = "" Then
        Return DrawingNumber
      Else
        Return _DrawingID
      End If
    End Get
    Set(value As String)
      _DrawingID = value
    End Set
  End Property

  Public Property Title As String = ""
  Public Property Revision As String = ""
  Public Property ElementID As String = ""
  Public Property SheetSize As String = ""
  Public Property Scale As String = ""
  Public Property Weight As String = ""
  Public Property DrawnBy As String = ""
  Public Property CheckedBy As String = ""
  Public Property ApprovedBy As String = ""
  Public Property CreationDate As String = ""
  Public Property Department As String = ""
  Public Property ForApproval As String = ""
  Public Property ForInformation As String = ""
  Public Property ForErection As String = ""
  Public Property ForProduction As String = ""
  Public Property ForCustomer As String = ""
  Public Property DocumentType As String = ""
  Public Property SoftwareName As String = ""
  Public Property ProjectID As String = ""
  Public Property Service1 As String = ""
  Public Property Service2 As String = ""
  Public Property IWT As String = ""
  Public Property ProjectYear As String = ""
  Public Property Consultant As String = ""
  Public Property ClientName As String = ""
  Public Property Division As String = ""
  Public Property Para1 As String = ""
  Public Property Para2 As String = ""

  Public Property DrawingFileName As String = ""

  Public Property Items As List(Of ItemData)
  Public Property RefDwgs As List(Of RefDwgData)
  Public Property AllParts As List(Of PartData)
  Public Property AllColumns As List(Of Columns)
  Public ReadOnly Property PDFFileName As String
    Get
      Return IO.Path.GetFileNameWithoutExtension(DrawingFileName) & ".PDF"
    End Get
  End Property

  Public Shared Sub WriteXMLDocument(ByVal xData As DrawingData, ByVal Job As jobFile, ByVal xmlPath As String)
    Dim document As New XmlDocument
    With xData
      Try

        Dim tmpN As XmlNode = document.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
        document.AppendChild(tmpN)

        tmpN = document.CreateElement("Document")
        document.AppendChild(tmpN)

        Dim tmpA As XmlAttribute = document.CreateAttribute("PDF_filename")
        tmpA.Value = .PDFFileName
        tmpN.Attributes.Append(tmpA)


        tmpA = document.CreateAttribute("filename")
        tmpA.Value = .DrawingFileName
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("number")
        tmpA.Value = .DrawingNumber
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("title")
        tmpA.Value = .Title
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("rev")
        tmpA.Value = .Revision
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("el.id.")
        tmpA.Value = .ElementID
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("sheetsize")
        tmpA.Value = .SheetSize
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("scale")
        tmpA.Value = .Scale
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("weight")
        tmpA.Value = .Weight
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("drawn")
        tmpA.Value = .DrawnBy
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("chqd")
        tmpA.Value = .CheckedBy
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("appd")
        tmpA.Value = .ApprovedBy
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("date")
        tmpA.Value = .CreationDate
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("resp_dept")
        tmpA.Value = .Department
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("apr")
        tmpA.Value = .ForApproval
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("inf")
        tmpA.Value = .ForInformation
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("pro")
        tmpA.Value = .ForProduction
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("ere")
        tmpA.Value = .ForErection
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("drgid")
        tmpA.Value = .DrawingID
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("VaultDBName")
        tmpA.Value = Job.VaultDB
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("VaultUserName")
        tmpA.Value = Job.UserID
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("VaultClientMachine")
        tmpA.Value = Job.ClientMachineName
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("VaultSubmittedDate")
        tmpA.Value = Job.JobCreationDate & " " & Job.JobCreationTime
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("ISGEC_Datasource")
        tmpA.Value = Job.FileName
        tmpN.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("name_software")
        tmpA.Value = .SoftwareName
        tmpN.Attributes.Append(tmpA)


        Dim prjNode As XmlNode = document.CreateElement("project")
        tmpA = document.CreateAttribute("contract")
        tmpA.Value = .ProjectID
        prjNode.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("service1")
        tmpA.Value = .Service1
        prjNode.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("service2")
        tmpA.Value = .Service2
        prjNode.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("iwt")
        tmpA.Value = .IWT
        prjNode.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("year")
        tmpA.Value = .ProjectYear
        prjNode.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("consultant")
        tmpA.Value = .Consultant
        prjNode.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("client")
        tmpA.Value = .ClientName
        prjNode.Attributes.Append(tmpA)

        tmpA = document.CreateAttribute("group")
        tmpA.Value = .Division
        prjNode.Attributes.Append(tmpA)

        tmpN.AppendChild(prjNode)

        Dim ItemsNode As XmlNode = document.CreateElement("Items")
        Dim partsNode As XmlNode = Nothing
        tmpN.AppendChild(ItemsNode)
        For Each tmpc As Columns In xData.AllColumns
          For Each data As PartData In .AllParts
            If data.ColumnNumber <> tmpc.ColumnNumber Then Continue For
            Dim itemNode As XmlNode = Nothing
            If data.LineType = "ITEM" Then
              itemNode = document.CreateElement("Item")

              tmpA = document.CreateAttribute("remark")
              tmpA.Value = data.ItemRemarks
              itemNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("it.qty")
              tmpA.Value = data.ItemQuantity
              itemNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("it.wt")
              tmpA.Value = data.ItemWeight
              itemNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("item_g")
              tmpA.Value = data.ItemGroup
              itemNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("t")
              tmpA.Value = data.ItemType
              itemNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("item_code")
              tmpA.Value = data.ItemCode
              itemNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("item_descn")
              tmpA.Value = data.ItemDescription
              itemNode.Attributes.Append(tmpA)

              ItemsNode.AppendChild(itemNode)

              partsNode = document.CreateElement("PartItems")

              itemNode.AppendChild(partsNode)

            End If
            If data.LineType = "PART" Then
              Dim partNode As XmlNode = document.CreateElement("PartItem")
              tmpA = document.CreateAttribute("p.no")
              tmpA.Value = data.PartNumber
              partNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("p_desc")
              tmpA.Value = data.PartDescription
              partNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("spec")
              tmpA.Value = data.PartSpecification
              partNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("size")
              tmpA.Value = data.PartSize
              partNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("p_qty")
              tmpA.Value = data.PartQuantity
              partNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("p_wt")
              tmpA.Value = data.PartWeight
              partNode.Attributes.Append(tmpA)

              tmpA = document.CreateAttribute("remark")
              tmpA.Value = data.PartRemarks & " " & data.PartRemarks1 & " " & data.PartRemarks2
              partNode.Attributes.Append(tmpA)

              If partsNode IsNot Nothing Then
                partsNode.AppendChild(partNode)
              End If

            End If
          Next
        Next
        '====================Start For XL=================
        'If xData is extracted From XL File (ComponenXL or MCD)
        'Then Write from Items & Part Items within Item
        For Each data As ItemData In .Items
          Dim itemNode As XmlNode = document.CreateElement("Item")
          tmpA = document.CreateAttribute("remark")
          tmpA.Value = data.ItemRemarks
          itemNode.Attributes.Append(tmpA)

          tmpA = document.CreateAttribute("it.qty")
          tmpA.Value = data.ItemQuantity
          itemNode.Attributes.Append(tmpA)

          tmpA = document.CreateAttribute("it.wt")
          tmpA.Value = data.ItemWeight
          itemNode.Attributes.Append(tmpA)

          tmpA = document.CreateAttribute("item_g")
          tmpA.Value = data.ItemGroup
          itemNode.Attributes.Append(tmpA)

          tmpA = document.CreateAttribute("t")
          tmpA.Value = data.ItemType
          itemNode.Attributes.Append(tmpA)

          tmpA = document.CreateAttribute("item_code")
          tmpA.Value = data.ItemCode
          itemNode.Attributes.Append(tmpA)

          tmpA = document.CreateAttribute("item_descn")
          tmpA.Value = data.ItemDescription
          itemNode.Attributes.Append(tmpA)

          ItemsNode.AppendChild(itemNode)

          'PartItem
          partsNode = document.CreateElement("PartItems")

          itemNode.AppendChild(partsNode)

          For Each pData As PartData In data.PartItems

            Dim partNode As XmlNode = document.CreateElement("PartItem")
            tmpA = document.CreateAttribute("p.no")
            tmpA.Value = pData.PartNumber
            partNode.Attributes.Append(tmpA)

            tmpA = document.CreateAttribute("p_desc")
            tmpA.Value = pData.PartDescription
            partNode.Attributes.Append(tmpA)

            tmpA = document.CreateAttribute("spec")
            tmpA.Value = pData.PartSpecification
            partNode.Attributes.Append(tmpA)

            tmpA = document.CreateAttribute("size")
            tmpA.Value = pData.PartSize
            partNode.Attributes.Append(tmpA)

            tmpA = document.CreateAttribute("p_qty")
            tmpA.Value = pData.PartQuantity
            partNode.Attributes.Append(tmpA)

            tmpA = document.CreateAttribute("p_wt")
            tmpA.Value = pData.PartWeight
            partNode.Attributes.Append(tmpA)

            tmpA = document.CreateAttribute("remark")
            tmpA.Value = pData.PartRemarks
            partNode.Attributes.Append(tmpA)

            partsNode.AppendChild(partNode)
          Next
        Next
        '================End Of XL===================
        Dim rdocsNode As XmlNode = document.CreateElement("ReferenceDocuments")
        tmpN.AppendChild(rdocsNode)
        For Each refDoc As RefDwgData In .RefDwgs
          Dim rdocNode As XmlNode = document.CreateElement("ReferenceDocument")
          tmpA = document.CreateAttribute("drgno")
          tmpA.Value = refDoc.DrawingID
          rdocNode.Attributes.Append(tmpA)

          tmpA = document.CreateAttribute("drg_descn")
          tmpA.Value = refDoc.DrawingName
          rdocNode.Attributes.Append(tmpA)

          tmpA = document.CreateAttribute("rev")
          tmpA.Value = refDoc.Revision
          rdocNode.Attributes.Append(tmpA)

          tmpA = document.CreateAttribute("PDF_filename")
          tmpA.Value = (refDoc.DrawingID & ".pdf")
          rdocNode.Attributes.Append(tmpA)

          tmpA = document.CreateAttribute("filename")
          tmpA.Value = (refDoc.DrawingID & ".dwg")
          rdocNode.Attributes.Append(tmpA)

          rdocsNode.AppendChild(rdocNode)
        Next
        'If (Me.referenceDrawingsData.Count > 1) Then
        '  Me.WriterefDwgsFile(Me.referenceDrawingsData)
        'End If
        'Dim node10 As XmlNode = document.CreateElement("AdditionalDocuments")
        'tmpN.AppendChild(node10)
      Catch obj1 As System.Exception

      End Try
    End With
    document.Save(xmlPath & "\" & IO.Path.GetFileNameWithoutExtension(xData.DrawingFileName) & ".xml")
  End Sub

  Public Shared Sub Save(ByVal FileName As String, ByVal oApl As DrawingData)
    Dim oSrz As XmlSerializer = New XmlSerializer(oApl.GetType)
    Dim oSW As IO.StreamWriter = New IO.StreamWriter(FileName)
    oSrz.Serialize(oSW, oApl)
    oSW.Close()
  End Sub
  Public Shared Function Open(ByVal FileName As String) As DrawingData
    If IO.File.Exists(FileName) Then
      Dim oApl As DrawingData = New DrawingData
      Dim oFS As IO.FileStream = New IO.FileStream(FileName, IO.FileMode.Open)
      Dim oSrz As XmlSerializer = New XmlSerializer(oApl.GetType)
      oApl = CType(oSrz.Deserialize(oFS), DrawingData)
      oFS.Close()
      Return oApl
    End If
    Return Nothing
  End Function
  Public Shared Sub CalculateColumns(ByVal xData As DrawingData)
    xData.AllColumns = New List(Of Columns)
    For Each tmpP As PartData In xData.AllParts
      Dim Found As Boolean = False
      For Each tmpC As Columns In xData.AllColumns
        If tmpC.IsSameColumn(tmpP.posMinX) Then
          tmpP.ColumnNumber = tmpC.ColumnNumber
          Found = True
          Exit For
        End If
      Next
      If Not Found Then
        Dim tmpC As New Columns
        tmpC.ColumnNumber = xData.AllColumns.Count + 1
        tmpC.posMinX = tmpP.posMinX
        tmpC.posMinY = tmpP.posMinY
        tmpC.posMaxX = tmpP.posMaxX
        tmpC.posMaxY = tmpP.posMaxY
        xData.AllColumns.Add(tmpC)
        tmpP.ColumnNumber = tmpC.ColumnNumber
      End If
    Next
    xData.AllParts.Sort(New Comparer)
  End Sub
  Sub New()
    Items = New List(Of ItemData)
    RefDwgs = New List(Of RefDwgData)
    AllParts = New List(Of PartData)
    AllColumns = New List(Of Columns)
  End Sub
#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
        Items = Nothing
        RefDwgs = Nothing
        AllParts = Nothing
        AllColumns = Nothing
      End If

      ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' TODO: set large fields to null.
    End If
    disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    ' TODO: uncomment the following line if Finalize() is overridden above.
    ' GC.SuppressFinalize(Me)
  End Sub

#End Region

End Class
Class Comparer
  Implements IComparer(Of PartData)

  Public Function Compare(x As PartData, y As PartData) As Integer Implements IComparer(Of PartData).Compare
    Dim result As Integer = y.ColumnNumber.CompareTo(x.ColumnNumber)
    If result = 0 Then
      result = x.posMinY.CompareTo(y.posMinY)
    End If
    Return result
  End Function
End Class
<Serializable>
Public Class ItemData
  Implements IDisposable
  Implements ICloneable
  Public Function Clone() As Object Implements ICloneable.Clone
    Return MyBase.MemberwiseClone()
  End Function

  Public Property ItemType As String = ""
  Public Property ItemCode As String = ""
  Public Property ItemDescription As String = ""
  Public Property ItemQuantity As String = ""
  Public Property ItemWeight As String = ""
  Public Property ItemGroup As String = ""
  Public Property ItemRemarks As String = ""
  Public Property posMinX As Decimal = 0
  Public Property posMinY As Decimal = 0
  Public Property posMaxX As Decimal = 0
  Public Property posMaxY As Decimal = 0
  Public Property PartItems As List(Of PartData)
  Sub New()
    PartItems = New List(Of PartData)
  End Sub
#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
        PartItems = Nothing
      End If

      ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' TODO: set large fields to null.
    End If
    disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    ' TODO: uncomment the following line if Finalize() is overridden above.
    ' GC.SuppressFinalize(Me)
  End Sub
#End Region
End Class
<Serializable>
Public Class PartData
  Implements ICloneable
  Public Function Clone() As Object Implements ICloneable.Clone
    Return MyBase.MemberwiseClone()
  End Function
  Public Property LineType As String = ""
  Public Property ColumnNumber As Integer = 0
  Public Property ItemType As String = ""
  Public Property ItemCode As String = ""
  Public Property ItemDescription As String = ""
  Public Property ItemQuantity As String = ""
  Public Property ItemWeight As String = ""
  Public Property ItemGroup As String = ""
  Public Property ItemRemarks As String = ""
  Public Property PartNumber As String = ""
  Public Property PartDescription As String = ""
  Public Property PartSpecification As String = ""
  Public Property PartSize As String = ""
  Public Property PartQuantity As String = ""
  Public Property PartWeight As String = ""
  Public Property PartRemarks As String = ""
  Public Property PartRemarks1 As String = ""
  Public Property PartRemarks2 As String = ""
  Public Property posMinX As Decimal = 0
  Public Property posMinY As Decimal = 0
  Public Property posMaxX As Decimal = 0
  Public Property posMaxY As Decimal = 0
  Sub New()
    'dummy
  End Sub

End Class
<Serializable>
Public Class RefDwgData
  Implements ICloneable
  Public Function Clone() As Object Implements ICloneable.Clone
    Return MyBase.MemberwiseClone()
  End Function
  Public Property DrawingID As String = ""
  Public Property Revision As String = ""
  Public Property DrawingName As String = ""
  Public Property FileName As String = ""
  Sub New()
    'dummy
  End Sub

End Class
<Serializable>
Public Class Columns
  Implements ICloneable
  Public Function Clone() As Object Implements ICloneable.Clone
    Return MyBase.MemberwiseClone()
  End Function
  Public Property ColumnNumber As Integer = 0
  Public Property posMinX As Decimal = 0
  Public Property posMinY As Decimal = 0
  Public Property posMaxX As Decimal = 0
  Public Property posMaxY As Decimal = 0
  Public ReadOnly Property Height As Decimal
    Get
      Return posMaxY - posMinY
    End Get
  End Property
  Public ReadOnly Property Width As Decimal
    Get
      Return posMaxX - posMinX
    End Get
  End Property
  Public Function IsSameColumn(ByVal posX As Decimal) As Boolean
    Dim mRet As Boolean = False
    If posX = posMinX Then
      mRet = True
    ElseIf posX < posMinX Then
      Dim tmpDistance As Decimal = posMinX - posX
      If tmpDistance < (Width * 0.4) Then
        mRet = True
      End If
    ElseIf posX > posMinX Then
      Dim tmpDistance As Decimal = posX - posMinX
      If tmpDistance < (Width * 0.4) Then
        mRet = True
      End If
    End If
    Return mRet
  End Function
  Sub New()
    'dummy
  End Sub

End Class
