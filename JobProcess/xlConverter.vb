Imports Microsoft.Office.Interop.Excel
Imports EDICommon
Public Class xlConverter
  Public Enum whichXL
    Component = 1
    MainFile = 2
    MCDFile = 3
  End Enum
  Public Shared Function GetDrawingData(ByVal FilePath As String) As List(Of DrawingData)
    Dim mRet As List(Of DrawingData) = New List(Of DrawingData)
    If IO.File.Exists(FilePath) Then
      Dim cover As New DrawingData
      Dim xlAp As Microsoft.Office.Interop.Excel.ApplicationClass = Nothing
      Dim xlWb As Workbook = Nothing
      Dim xlWs As Worksheet = Nothing
      Try
        xlAp = New Microsoft.Office.Interop.Excel.ApplicationClass
        xlAp.DisplayAlerts = False
        xlWb = xlAp.Workbooks.Open(FilePath)
        xlWs = xlWb.Worksheets("BOM")
        Dim rng As Range = xlWs.UsedRange
        Dim rows As Integer = rng.Rows.Count + 6
        Dim cols As Integer = rng.Columns.Count
        xlWs = xlWb.Worksheets("CoverPage")
        With cover
          .ProjectID = xlWs.Cells(2, 3).Value
          .IWT = xlWs.Cells(3, 3).Value
          .ProjectYear = xlWs.Cells(4, 3).Value
          .ClientName = xlWs.Cells(5, 3).Value
          .Consultant = xlWs.Cells(6, 3).Value
          .Service1 = xlWs.Cells(7, 3).Value
          .Service2 = xlWs.Cells(8, 3).Value
          .ElementID = xlWs.Cells(9, 3).Value
          .Division = xlWs.Cells(10, 3).Value
          .Department = xlWs.Cells(11, 3).Value
        End With
        Dim tmp As RefDwgData = Nothing
        For I As Integer = 15 To 999
          If xlWs.Cells(I, 3).Value <> String.Empty Then
            tmp = New RefDwgData
            With tmp
              .DrawingID = xlWs.Cells(I, 3).Value
              .DrawingName = xlWs.Cells(I, 4).Value
              .Revision = xlWs.Cells(I, 5).Value
              .FileName = xlWs.Cells(I, 6).Value
            End With
            cover.RefDwgs.Add(tmp)
          Else
            Exit For
          End If
        Next
        xlWs = xlWb.Worksheets("BOM")
        Dim l_doc As String = ""
        Dim l_itm As String = ""
        Dim l_prt As String = ""
        Dim c_doc As String = ""
        Dim c_itm As String = ""
        Dim c_prt As String = ""
        Dim NoRowFound As Boolean = True
        Dim tmpDoc As DrawingData = Nothing
        Dim tillRow As Integer = 12
        'Find Till Row
        For I As Integer = rows To 1 Step -1
          If xlWs.Range("M" & I).Value <> "" AndAlso xlWs.Range("M" & I).Value.ToString.ToLower = "drawing number" Then
            tillRow = I + 1
            Exit For
          End If
        Next
        'TillRow Found
        For I As Integer = rows To tillRow Step -1
          If l_doc = "" Then
            l_doc = xlWs.Range("M" & I).Value
            l_itm = xlWs.Range("B" & I).Value
            l_prt = xlWs.Range("C" & I).Value
            If l_doc Is Nothing And l_itm Is Nothing And l_prt Is Nothing Then
              If NoRowFound Then
                Continue For
              End If
            Else
              NoRowFound = False
            End If
            If l_doc <> String.Empty Then
              tmpDoc = cover.Clone
              FillDoc(xlWs, I, tmpDoc)
              If l_itm <> String.Empty Then
                FillItm(xlWs, I, tmpDoc)
              End If
              mRet.Add(tmpDoc)
            Else
              Exit For
            End If
            Continue For
          End If
          c_doc = xlWs.Range("M" & I).Value
          c_itm = xlWs.Range("B" & I).Value
          c_prt = xlWs.Range("C" & I).Value
          If c_doc = String.Empty Then
            If c_itm = String.Empty Then
              If c_prt <> String.Empty Then
                FillPrt(xlWs, I, tmpDoc)
              End If
            Else
              FillItm(xlWs, I, tmpDoc)
            End If
          Else
            tmpDoc = Nothing
            tmpDoc = cover.Clone
            FillDoc(xlWs, I, tmpDoc)
            If c_itm <> String.Empty Then
              FillItm(xlWs, I, tmpDoc)
            End If
            mRet.Add(tmpDoc)
          End If
        Next
        xlWs = Nothing
        xlWb.Close()
        xlAp.Quit()
      Catch ex As Exception
        xlWs = Nothing
        If xlWs IsNot Nothing Then
          xlWb.Close()
        End If
        If xlAp IsNot Nothing Then
          xlAp.Quit()
        End If
      End Try
    End If
    Return mRet
  End Function

  Public Shared Function GetDrawingData(ByVal Job As jobFile, ByVal jpConfig As ConfigFile, ByVal whichOne As whichXL) As List(Of DrawingData)
    Dim mRet As List(Of DrawingData) = Nothing
    Dim FilePath As String = ""
    Select Case whichOne
      Case whichXL.Component
        FilePath = jpConfig.TempFolderPath & "\" & Job.ComponentXLFileName
      Case whichXL.MainFile
        FilePath = jpConfig.TempFolderPath & "\" & Job.FileName
    End Select
    Return GetDrawingData(FilePath)
  End Function
  Private Shared Sub FillPrt(ByVal xlWs As Worksheet, ByVal I As Integer, ByVal tmpDoc As DrawingData)
    Dim itm As New PartData
    With itm
      .PartWeight = xlWs.Range("K" & I).Value
      .PartQuantity = xlWs.Range("J" & I).Value
      .PartSize = xlWs.Range("H" & I).Value
      .PartSpecification = xlWs.Range("G" & I).Value
      .PartDescription = xlWs.Range("D" & I).Value
      .PartNumber = xlWs.Range("C" & I).Value
    End With
    tmpDoc.Items(tmpDoc.Items.Count - 1).PartItems.Add(itm)
  End Sub
  Private Shared Sub FillItm(ByVal xlWs As Worksheet, ByVal I As Integer, ByVal tmpDoc As DrawingData)
    Dim itm As New ItemData
    With itm
      .ItemWeight = xlWs.Range("K" & I).Value
      .ItemQuantity = xlWs.Range("J" & I).Value
      .ItemGroup = xlWs.Range("F" & I).Value
      .ItemType = xlWs.Range("E" & I).Value
      .ItemDescription = xlWs.Range("D" & I).Value
      .ItemCode = xlWs.Range("B" & I).Value
      .ItemRemarks = xlWs.Range("A" & I).Value
    End With
    tmpDoc.Items.Add(itm)
  End Sub

  Private Shared Sub FillDoc(ByVal xlWs As Worksheet, ByVal I As Integer, ByVal tmpDoc As DrawingData)
    With tmpDoc
      .SoftwareName = xlWs.Range("AB" & I).Value
      .Scale = xlWs.Range("AA" & I).Value
      .SheetSize = xlWs.Range("Z" & I).Value
      .Weight = xlWs.Range("Y" & I).Value
      .ForErection = xlWs.Range("X" & I).Value
      .ForProduction = xlWs.Range("W" & I).Value
      .ForApproval = xlWs.Range("V" & I).Value
      .ForInformation = xlWs.Range("U" & I).Value
      .CreationDate = xlWs.Range("T" & I).Value
      .ApprovedBy = xlWs.Range("S" & I).Value
      .CheckedBy = xlWs.Range("R" & I).Value
      .DrawnBy = xlWs.Range("Q" & I).Value
      .DrawingFileName = xlWs.Range("P" & I).Value
      .Title = xlWs.Range("O" & I).Value
      .Revision = xlWs.Range("N" & I).Value
      .DrawingID = xlWs.Range("M" & I).Value
      .DrawingNumber = .DrawingID
    End With
  End Sub
End Class
