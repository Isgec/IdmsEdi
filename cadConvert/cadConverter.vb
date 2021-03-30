Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
Imports System.Reflection
Imports EDICommon
Public Class cadConverter
  Private AcadSheetSizes As Dictionary(Of String, String)
  Private jpConfig As ConfigFile = Nothing
  Private Job As jobFile = Nothing
  Private TitleBlockName As String = ""

  Public Sub ReadJobFile(ByVal jFile As String)
    Try
      Dim FileName As String = jpConfig.JobPathWorking & "\" & jFile
      Job = jobFile.DeSerialize(Job, FileName)
    Catch ex As System.Exception
      Throw New System.Exception("Error while reading job file.")
    End Try

  End Sub


  <CommandMethod("ReadConfigFile")>
  Public Sub ReadConfigFile()
    Try
      Dim FileName As String = IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly.Location) & "\jpConfig.xml"
      jpConfig = ConfigFile.DeSerialize(jpConfig, FileName)
    Catch ex As System.Exception
      Throw New System.Exception("Error while reading config file.")
    End Try

  End Sub
  <CommandMethod("ExtractBOMPDF")>
  Public Sub ExtractBOMPDF()
    ReadConfigFile()
    doExtractBOM()
    doExtractPDF()
  End Sub

  <CommandMethod("ExtractBOM")>
  Public Sub ExtractBOM()
    ReadConfigFile()
    doExtractBOM()
  End Sub
  Private Sub doExtractBOM()
    Dim acDoc As Document = Nothing
    Try
      acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

      ReadJobFile(IO.Path.GetFileNameWithoutExtension(acDoc.Name) & ".slz")

      Dim lock As DocumentLock = acDoc.LockDocument
      Dim database As Database = acDoc.Database
      Dim xData As New DrawingData
      With xData
        .DrawingFileName = IO.Path.GetFileName(acDoc.Name)
      End With

      Using Transaction As Transaction = database.TransactionManager.StartTransaction
        Dim Table As BlockTable = TryCast(Transaction.GetObject(database.BlockTableId, 0), BlockTable)
        Dim Record As BlockTableRecord = TryCast(Transaction.GetObject(Table.Item(BlockTableRecord.ModelSpace), 0), BlockTableRecord)
        For Each ID As ObjectId In Record
          If (ID.ObjectClass.Name = "AcDbBlockReference") Then
            Dim reference As BlockReference = CType(Transaction.GetObject(ID, 0), BlockReference)
            '0. Title
            Dim TitleBlockFound As Boolean = False
            For Each tmp As String In jpConfig.TitleBlockNames.Split(",".ToCharArray)
              If (reference.Name.ToUpper = tmp.ToUpper) Then
                TitleBlockFound = True
                TitleBlockName = tmp
                Job.IsTitleBlockFound = True
                Exit For
              End If
            Next
            If TitleBlockFound Then
              Dim attributes As AttributeCollection = reference.AttributeCollection
              If (Not attributes Is Nothing) Then
                For Each tmpID As ObjectId In attributes
                  Dim tmpRef As AttributeReference = TryCast(Transaction.GetObject(tmpID, 0), AttributeReference)
                  If (tmpRef.Tag = "NUMBER") Then
                    xData.DrawingNumber = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "00") Then
                    xData.Revision = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "TITLE") Then
                    xData.Title = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "CONTRACT") Then
                    xData.ProjectID = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "EL.ID.") Then
                    xData.ElementID = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "IWT") Then
                    xData.IWT = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "SERVICE1") Then
                    xData.Service1 = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "SERVICE2") Then
                    xData.Service2 = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "CLIENT") Then
                    xData.ClientName = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "CONSULTANT") Then
                    xData.Consultant = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "GROUP") Then
                    xData.Division = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "RESP_DEPT") Then
                    xData.Department = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "YEAR") Then
                    xData.ProjectYear = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "SHEETSIZE") Then
                    xData.SheetSize = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "SCALE") Then
                    xData.Scale = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "WEIGHT") Then
                    xData.Weight = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "DRAWN") Then
                    xData.DrawnBy = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "CHQD") Then
                    xData.CheckedBy = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "APPD") Then
                    xData.ApprovedBy = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "DATE") Then
                    xData.CreationDate = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "APR") Then
                    xData.ForApproval = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "INF") Then
                    xData.ForInformation = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "PRO") Then
                    xData.ForProduction = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "ERE") Then
                    xData.ForErection = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "NAME_SOFTWARE") Then
                    xData.SoftwareName = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "DRGID") Then
                    xData.DrawingID = tmpRef.TextString.Trim
                  End If
                Next
              End If
            End If 'End 0. TitleBlockFound
            '1. Item
            Dim ItemBlockFound As Boolean = False


            For Each tmp As String In jpConfig.BAANBlockNames.Split(",".ToCharArray)
              If (reference.Name.ToUpper = tmp.ToUpper) Then
                ItemBlockFound = True
                Exit For
              End If
            Next
            If ItemBlockFound Then
              Dim attributes As AttributeCollection = reference.AttributeCollection
              If (Not attributes Is Nothing) Then
                Dim tmpItem As New PartData
                For Each tmpID As ObjectId In attributes
                  Dim tmpRef As AttributeReference = TryCast(Transaction.GetObject(tmpID, 0), AttributeReference)
                  If (tmpRef.Tag = "REMARK") Then
                    tmpItem.ItemRemarks = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "IT.QTY") Then
                    tmpItem.ItemQuantity = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "IT.WT") Then
                    tmpItem.ItemWeight = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "ITEM_G") Then
                    tmpItem.ItemGroup = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "T") Then
                    tmpItem.ItemType = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "ITEM_CODE") Then
                    tmpItem.ItemCode = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "ITEM_DESCN") Then
                    tmpItem.ItemDescription = tmpRef.TextString.Trim
                  End If
                Next
                tmpItem.posMinX = reference.Bounds.Value.MinPoint.X
                tmpItem.posMinY = reference.Bounds.Value.MinPoint.Y
                tmpItem.posMaxX = reference.Bounds.Value.MaxPoint.X
                tmpItem.posMaxY = reference.Bounds.Value.MaxPoint.Y
                tmpItem.LineType = "ITEM"
                xData.AllParts.Add(tmpItem)
              End If
            End If
            '2. Part
            Dim PartBlockFound As Boolean = False
            For Each tmp As String In jpConfig.PartItemBlockNames.Split(",".ToCharArray)
              If (reference.Name.ToUpper = tmp.ToUpper) Then
                PartBlockFound = True
                Exit For
              End If
            Next
            If PartBlockFound Then
              Dim attributes As AttributeCollection = reference.AttributeCollection
              If (Not attributes Is Nothing) Then
                Dim tmpPart As New PartData
                For Each tmpID As ObjectId In attributes
                  Dim tmpRef As AttributeReference = TryCast(Transaction.GetObject(tmpID, 0), AttributeReference)
                  If (tmpRef.Tag = "P.NO") Then
                    tmpPart.PartNumber = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "P_DESC") Then
                    tmpPart.PartDescription = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "SPEC") Then
                    tmpPart.PartSpecification = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "SIZE") Then
                    tmpPart.PartSize = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "P_QTY") Then
                    tmpPart.PartQuantity = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "P_WT") Then
                    tmpPart.PartWeight = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "REMARK") Then
                    tmpPart.PartRemarks = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "REMARK1") Then
                    tmpPart.PartRemarks1 = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "REMARK2") Then
                    tmpPart.PartRemarks2 = tmpRef.TextString.Trim
                  End If
                Next
                tmpPart.posMinX = reference.Bounds.Value.MinPoint.X
                tmpPart.posMinY = reference.Bounds.Value.MinPoint.Y
                tmpPart.posMaxX = reference.Bounds.Value.MaxPoint.X
                tmpPart.posMaxY = reference.Bounds.Value.MaxPoint.Y
                tmpPart.LineType = "PART"
                xData.AllParts.Add(tmpPart)
              End If
            End If
            '3. Ref Dwg
            Dim RefDwgBlockFound As Boolean = False
            For Each tmp As String In jpConfig.ReferenceDrawingBlockNames.Split(",".ToCharArray)
              If (reference.Name.ToUpper = tmp.ToUpper) Then
                RefDwgBlockFound = True
                Exit For
              End If
            Next
            If RefDwgBlockFound Then
              Dim key As String = ""
              Dim str5 As String = ""
              Dim attributes As AttributeCollection = reference.AttributeCollection
              If (Not attributes Is Nothing) Then
                Dim tmpDwg As New RefDwgData
                For Each tmpId As ObjectId In attributes
                  Dim tmpRef As AttributeReference = TryCast(Transaction.GetObject(tmpId, 0), AttributeReference)
                  If (tmpRef.Tag = "DRGNO") Then
                    tmpDwg.DrawingID = tmpRef.TextString.Trim
                  ElseIf (tmpRef.Tag = "DRG_DESCN") Then
                    tmpDwg.DrawingName = tmpRef.TextString.Trim
                  End If
                Next
                xData.RefDwgs.Add(tmpDwg)
              End If
            End If 'End 3.
          End If
        Next
      End Using
      lock.Dispose()

      'Try
      '  Dim FileName As String = jpConfig.JobPathWorking & "\" & IO.Path.GetFileNameWithoutExtension(acDoc.Name) & ".xml"
      '  Dim oSrz As XmlSerializer = New XmlSerializer(xData.GetType)
      '  Dim oSW As IO.StreamWriter = New IO.StreamWriter(FileName)
      '  oSrz.Serialize(oSW, xData)
      '  oSW.Close()
      'Catch ex1 As System.Exception
      'End Try
      Try
        DrawingData.CalculateColumns(xData)
        DrawingData.WriteXMLDocument(xData, Job, jpConfig.XMLFolderPath)
        Job.IsXMLGenerated = True
      Catch ex1 As System.Exception
        Job.IsError = True
        Job.ErrorList.Add(New lgErrors(ex1.Message))
      End Try
    Catch ex As System.Exception
      Job.IsError = True
      Job.ErrorList.Add(New lgErrors(ex.Message))
    End Try
    Job = jobFile.Serialize(jpConfig.JobPathWorking, Job)
  End Sub

  <CommandMethod("ExtractPDF")>
  Public Sub ExtractPDF()
    ReadConfigFile()
    doExtractPDF()
  End Sub
  Private Sub doExtractPDF()
    Try
      Dim acDoc As Document = Nothing
      acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

      ReadJobFile(IO.Path.GetFileNameWithoutExtension(acDoc.Name) & ".slz")

      Dim database As Database = acDoc.Database
      Me.AcadSheetSizes = New Dictionary(Of String, String)

      Dim CADFileName As String = IO.Path.GetFileName(acDoc.Name)

      Me.populateAcadSheetSizes()
      If TitleBlockName = "" Then
        Job.IsTitleBlockFound = getTitleBlockName(acDoc)
      End If
      Try
        CreateOrEditPageSetup()
      Catch obj2 As System.Exception
      End Try
      Using transaction As Transaction = database.TransactionManager.StartTransaction
        Dim manager As LayoutManager = Nothing
        manager = LayoutManager.Current
        Dim layout As Layout = Nothing
        layout = TryCast(transaction.GetObject(manager.GetLayoutId(manager.CurrentLayout), OpenMode.ForRead), Layout)
        Dim dictionary As DBDictionary = TryCast(transaction.GetObject(database.PlotSettingsDictionaryId, OpenMode.ForRead), DBDictionary)
        Dim settings As PlotSettings = Nothing
        If dictionary.Contains("MyPageSetup") Then
          settings = TryCast(dictionary.GetAt("MyPageSetup").GetObject(OpenMode.ForRead), PlotSettings)
          layout.UpgradeOpen()
          layout.CopyFrom(settings)
          'transaction.Commit()
        Else
          'transaction.Abort()
        End If
        acDoc.Editor.Regen()
        Dim info As New PlotInfo
        info.Layout = layout.ObjectId
        'Dim PlotSettingValidator As PlotSettingsValidator = PlotSettingsValidator.Current
        Dim validator As New PlotInfoValidator
        validator.MediaMatchingPolicy = MatchingPolicy.MatchEnabled
        validator.Validate(info)
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BackGroundPlot", 0)
        Using engine As PlotEngine = PlotFactory.CreatePublishEngine
          Dim dialog As New PlotProgressDialog(False, 1, True)
          Using dialog
            dialog.OnBeginPlot()
            dialog.IsVisible = False
            engine.BeginPlot(dialog, Nothing)
            Dim PDFFilePath As String = (jpConfig.PDFFolderPath & "\" & IO.Path.GetFileNameWithoutExtension(CADFileName) & ".PDF")
            engine.BeginDocument(info, acDoc.Name, Nothing, 1, True, PDFFilePath)
            dialog.OnBeginSheet()
            dialog.LowerSheetProgressRange = 0
            dialog.UpperSheetProgressRange = 100
            dialog.SheetProgressPos = 0
            Dim info2 As New PlotPageInfo
            engine.BeginPage(info2, info, True, Nothing)
            engine.BeginGenerateGraphics(Nothing)
            engine.EndGenerateGraphics(Nothing)
            engine.EndPage(Nothing)
            dialog.SheetProgressPos = 100
            dialog.OnEndSheet()
            engine.EndDocument(Nothing)
            dialog.PlotProgressPos = 100
            dialog.OnEndPlot()
            engine.EndPlot(Nothing)
          End Using
        End Using
      End Using
      Job.IsPDFCreated = True
    Catch exception As System.Exception
      Job.IsError = True
      Job.ErrorList.Add(New lgErrors(exception.Message))
    End Try
  End Sub

  Private Sub populateAcadSheetSizes()
    Try
      Me.AcadSheetSizes.Add("B5", "ISO_full_bleed_B5_(250.00_x_176.00_MM)")
      Me.AcadSheetSizes.Add("B4", "ISO_full_bleed_B4_(353.00_x_250.00_MM)")
      Me.AcadSheetSizes.Add("B3", "ISO_full_bleed_B3_(500.00_x_353.00_MM)")
      Me.AcadSheetSizes.Add("B2", "ISO_full_bleed_B2_(707.00_x_500.00_MM)")
      Me.AcadSheetSizes.Add("B1", "ISO_full_bleed_B1_(1000.00_x_707.00_MM)")
      Me.AcadSheetSizes.Add("B0", "ISO_full_bleed_B0_(1414.00_x_1000.00_MM)")
      Me.AcadSheetSizes.Add("A5", "ISO_full_bleed_A5_(210.00_x_148.00_MM)")
      Me.AcadSheetSizes.Add("4A0", "ISO_full_bleed_4A0_(1682.00_x_2378.00_MM)")
      Me.AcadSheetSizes.Add("2A0", "ISO_full_bleed_2A0_(1189.00_x_1682.00_MM)")
      Me.AcadSheetSizes.Add("A4", "ISO_full_bleed_A4_(297.00_x_210.00_MM)")
      Me.AcadSheetSizes.Add("A3", "ISO_full_bleed_A3_(420.00_x_297.00_MM)")
      Me.AcadSheetSizes.Add("A2", "ISO_full_bleed_A2_(594.00_x_420.00_MM)")
      Me.AcadSheetSizes.Add("A1", "ISO_full_bleed_A1_(841.00_x_594.00_MM)")
      Me.AcadSheetSizes.Add("A0", "ISO_full_bleed_A0_(841.00_x_1189.00_MM)")
    Catch exception As System.Exception
    End Try
  End Sub
  Public Sub CreateOrEditPageSetup()
    Try
      Dim database As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
      Using transaction As Transaction = database.TransactionManager.StartTransaction
        Dim dictionary As DBDictionary = TryCast(transaction.GetObject(database.PlotSettingsDictionaryId, OpenMode.ForRead), DBDictionary)
        transaction.GetObject(database.VisualStyleDictionaryId, OpenMode.ForRead)
        Dim settings As PlotSettings = Nothing
        Dim flag As Boolean = False
        Dim manager As LayoutManager = LayoutManager.Current
        Dim layout As Layout = TryCast(transaction.GetObject(manager.GetLayoutId(manager.CurrentLayout), OpenMode.ForRead), Layout)
        If Not dictionary.Contains("MyPageSetup") Then
          flag = True
          settings = New PlotSettings(layout.ModelType)
          settings.CopyFrom(layout)
          settings.PlotSettingsName = "MyPageSetup"
          settings.AddToPlotSettingsDictionary(database)
          transaction.AddNewlyCreatedDBObject(settings, True)
        Else
          settings = TryCast(dictionary.GetAt("MyPageSetup").GetObject(OpenMode.ForWrite), PlotSettings)
        End If
        Try
          Dim validator As PlotSettingsValidator = PlotSettingsValidator.Current
          Dim str As String = ""
          Try
            Dim key As String = ""
            Dim str3 As String = ""
            If (TitleBlockName <> "") Then
              key = TitleBlockName.Substring((TitleBlockName.LastIndexOf("_") + 1))
              If (key <> "") Then
                If AcadSheetSizes.TryGetValue(key, str3) Then
                  str = str3
                Else
                  str = "ISO_full_bleed_A0_(841.00_x_1189.00_MM)"
                End If
              Else
                str = "ISO_full_bleed_A0_(841.00_x_1189.00_MM)"
              End If
            Else
              str = "ISO_full_bleed_A0_(841.00_x_1189.00_MM)"
            End If
          Catch obj1 As System.Exception
          End Try
          validator.SetPlotConfigurationName(settings, "DWG To PDF.pc3", str)
          If Not layout.ModelType Then
            validator.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout)
          Else
            validator.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents)
            validator.SetPlotCentered(settings, True)
          End If
          validator.SetPlotOrigin(settings, New Point2d(0, 0))
          validator.SetUseStandardScale(settings, True)
          validator.SetStdScaleType(settings, 0)
          validator.SetPlotPaperUnits(settings, 0)
          settings.ScaleLineweights = True
          settings.ShowPlotStyles = True
          validator.RefreshLists(settings)
          settings.ShadePlot = 0
          settings.ShadePlotResLevel = ShadePlotResLevel.Normal
          settings.PrintLineweights = True
          settings.PlotTransparency = False
          settings.PlotPlotStyles = True
          settings.DrawViewportsFirst = True
          validator.SetPlotRotation(settings, 0)
          If database.PlotStyleMode Then
            validator.SetCurrentStyleSheet(settings, "monochrome.ctb")
          Else
            validator.SetCurrentStyleSheet(settings, "monochrome.stb")
          End If
          validator.SetZoomToPaperOnUpdate(settings, True)
        Catch exception As System.Exception
        End Try
        transaction.Commit()
        If flag Then
          settings.Dispose()
        End If
      End Using
    Catch exception2 As System.Exception
    End Try
  End Sub
  Private Function getTitleBlockName(ByVal acDoc As Autodesk.AutoCAD.ApplicationServices.Document) As Boolean
    Dim flag As Boolean
    Try
      Dim database As Database = acDoc.Database
      Dim reference As BlockReference = Nothing
      Dim lock As DocumentLock = acDoc.LockDocument
      Dim num As Integer = 0
      Using transaction As Transaction = database.TransactionManager.StartTransaction
        Dim table As BlockTable = TryCast(transaction.GetObject(database.BlockTableId, 0), BlockTable)
        Dim record As BlockTableRecord = TryCast(transaction.GetObject(table.Item(BlockTableRecord.ModelSpace), 1), BlockTableRecord)
        Dim id As ObjectId
        For Each id In record
          If (id.ObjectClass.Name = "AcDbBlockReference") Then
            reference = transaction.GetObject(id, 0)
            Dim str As String
            For Each str In jpConfig.TitleBlockNames.Split(New Char() {","c})
              If (reference.Name.ToUpper = str) Then
                num += 1
                TitleBlockName = reference.Name.ToUpper
              End If
            Next
          End If
        Next
        If (num = 0) Then
          Return False
        End If
        transaction.Commit()
        lock.Dispose()
        flag = True
      End Using
    Catch obj1 As System.Exception
      flag = False
    End Try
    Return flag
  End Function
End Class
