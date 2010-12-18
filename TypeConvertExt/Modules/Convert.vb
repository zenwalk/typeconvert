Option Strict Off
Option Explicit On
Module Convert

    Public Enum ConvertType
        TO_POLYGON = 1
        TO_POLYLINE = 2
        TO_POINT = 3
        CONVEX_HULL = 4
        ENVELOPE_RECTANGLE = 5
        STRATIFICATION = 6
        F_CENTROID = 7
        REMOVE_DUPLICATES = 8
        SEGMENTS = 9
    End Enum


    Sub StartConvert(ByRef pMap As IMap, ByRef pLayer As ILayer, ByRef iConvertType As Short)
        Dim I As Object
        Dim Y_Max As Object
        Dim Y_Min As Object
        Dim X_Max As Object
        Dim X_Min As Object
        Dim pExpOp As IExportOperation
        Dim pDSName As IDatasetName
        pExpOp = New ExportOperation
        Dim pWorkspace As IWorkspace
        Dim pName As IName
        Dim pNewFeatureClass As IFeatureClass
        Dim NewFeatureClasses As FeatureClassArray
        Dim NewFeatureLayers As FeatureLayerArray
        Dim hasSelection As Boolean
        Dim supportMapProjection As Boolean
        Dim saveProjection As Boolean
        Dim Exp_option As esriExportTableOptions
        Exp_option = esriExportTableOptions.esriExportAllRecords
        supportMapProjection = True

        Dim pfLayer As IFeatureLayer
        pfLayer = pLayer

        Dim pFeatureSel As IFeatureSelection
        pFeatureSel = pLayer
        Dim pSelSet As ISelectionSet
        pSelSet = pFeatureSel.SelectionSet
        If pSelSet.Count <> 0 Then
            hasSelection = True
        Else
            hasSelection = False
        End If

        pDSName = pExpOp.GetOptions(pfLayer.FeatureClass, pLayer.Name, hasSelection, supportMapProjection, 0, saveProjection, Exp_option)
        pExpOp = Nothing

        Dim pActiveView As IActiveView
        Dim pEnvelope As IEnvelope
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeature As IFeature
        Dim pFeatureTypeConverter As New FeatureTypeConverter
        Dim pSR As ISpatialReference
        Dim pGeoDataset As IGeoDataset
        Dim pNewFeatureLayer As IFeatureLayer
        If pDSName.Name <> "" Then

            If Exp_option = esriExportTableOptions.esriExportAllRecords Then
                pSelSet = Nothing
            End If

            If Exp_option = esriExportTableOptions.esriExportFeaturesWithinExtent Then
                pActiveView = pMap
                pEnvelope = pActiveView.Extent
                X_Min = pEnvelope.XMin
                X_Max = pEnvelope.XMax
                Y_Min = pEnvelope.YMin
                Y_Max = pEnvelope.YMax

                pFeatureCursor = pfLayer.FeatureClass.Search(Nothing, False)
                pFeature = pFeatureCursor.NextFeature
                Do While Not pFeature Is Nothing
                    If pFeature.Extent.XMax <= X_Max And pFeature.Extent.XMin >= X_Min And pFeature.Extent.YMax <= Y_Max And pFeature.Extent.YMin >= Y_Min Then
                        pFeatureSel.Add(pFeature)
                    End If
                    pFeature = pFeatureCursor.NextFeature
                Loop
                pSelSet = pFeatureSel.SelectionSet
                pFeatureSel.Clear()
            End If

            pName = pDSName.WorkspaceName
            pWorkspace = pName.Open
            pFeatureTypeConverter.DatasetName = pDSName


            If saveProjection Then
                pSR = pMap.SpatialReference
            Else
                pGeoDataset = pLayer
                pSR = pGeoDataset.SpatialReference
            End If
            If Not (pSR Is Nothing) Then
                pFeatureTypeConverter.SpatialReference = pSR
            Else
                pFeatureTypeConverter.SpatialReference = New UnknownCoordinateSystem
            End If

            Select Case iConvertType

                Case ConvertType.TO_POLYLINE

                    pNewFeatureClass = pFeatureTypeConverter.ToPolyline((pfLayer.FeatureClass), pSelSet)

                Case ConvertType.TO_POLYGON

                    pNewFeatureClass = pFeatureTypeConverter.ToPolygon((pfLayer.FeatureClass), pSelSet)

                Case ConvertType.TO_POINT

                    pNewFeatureClass = pFeatureTypeConverter.ToPoint((pfLayer.FeatureClass), pSelSet)

                Case ConvertType.SEGMENTS

                    pNewFeatureClass = pFeatureTypeConverter.ToSegments((pfLayer.FeatureClass), pSelSet)

                Case ConvertType.CONVEX_HULL

                    pNewFeatureClass = pFeatureTypeConverter.ToConvexHull((pfLayer.FeatureClass), pSelSet)

                Case ConvertType.ENVELOPE_RECTANGLE

                    pNewFeatureClass = pFeatureTypeConverter.ToEnvelope((pfLayer.FeatureClass), pSelSet)

                Case ConvertType.F_CENTROID

                    pNewFeatureClass = pFeatureTypeConverter.ToCentroid((pfLayer.FeatureClass), pSelSet)

                Case ConvertType.REMOVE_DUPLICATES

                    pNewFeatureClass = Nothing
                    NewFeatureClasses = pFeatureTypeConverter.RemoveDuplicates((pfLayer.FeatureClass), pSelSet)

                Case ConvertType.STRATIFICATION

                    pNewFeatureClass = Nothing
                    NewFeatureLayers = pFeatureTypeConverter.Stratify(pfLayer, pSelSet)
            End Select

            On Error GoTo err_h

            If Not pNewFeatureClass Is Nothing Then
                If MsgBox("Do you want to add converted data to the map as new layer ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Type converter") = MsgBoxResult.Yes Then

                    pNewFeatureLayer = New FeatureLayer
                    pNewFeatureLayer.FeatureClass = pNewFeatureClass
                    pNewFeatureLayer.Name = pNewFeatureClass.AliasName
                    pMap.AddLayer(pNewFeatureLayer)
                End If
            End If

            If (iConvertType = ConvertType.REMOVE_DUPLICATES) Then
                If (UBound(NewFeatureClasses.FeatureClass) <> 0) Then
                    If MsgBox("Do you want to add converted data to the map as new layers ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Type converter") = MsgBoxResult.Yes Then
                        For I = 0 To UBound(NewFeatureClasses.FeatureClass)
                            If Not NewFeatureClasses.FeatureClass(I) Is Nothing Then
                                pNewFeatureLayer = New FeatureLayer
                                pNewFeatureLayer.FeatureClass = NewFeatureClasses.FeatureClass(I)
                                pNewFeatureLayer.Name = NewFeatureClasses.FeatureClass(I).AliasName
                                pMap.AddLayer(pNewFeatureLayer)
                            End If
                        Next I
                    End If
                End If
            End If

            If (iConvertType = ConvertType.STRATIFICATION) Then
                If MsgBox("Do you want to add converted data to the map as new layers ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Type converter") = MsgBoxResult.Yes Then
                    With NewFeatureLayers
                        For I = 0 To UBound(.FeatureLayer)
                            If Not .FeatureLayer(I) Is Nothing Then
                                pMap.AddLayer(.FeatureLayer(I))
                            End If
                        Next I
                    End With
                End If
            End If
        End If

        pFeatureTypeConverter = Nothing
        pDSName = Nothing

        Exit Sub

err_h:
        MsgBox(Err.Description)

    End Sub
    Public Sub DivideSegments(ByVal pLayer As IFeatureLayer, ByVal pMap As IMap)
        Dim pFeatureTypeConverter As New FeatureTypeConverter
        Dim pFeatureSel As IFeatureSelection
        pFeatureSel = pLayer
        Dim pSelSet As ISelectionSet
        pSelSet = pFeatureSel.SelectionSet
        If pSelSet.Count = 0 Then
            pSelSet = Nothing
        End If

        Dim Lp As String
        Lp = InputBox("Input interval length (in display units)", "TypeConvert")
        If Not IsNumeric(Lp) OrElse CDbl(Lp) < 0 Then
            'MsgBox("Length value is not valid", MsgBoxStyle.Critical, "TypeConvert")
            Exit Sub
        End If
        If pFeatureTypeConverter.DivideSegments(pLayer.FeatureClass, pSelSet, pMap, CDbl(Lp)) Then
            MsgBox("Segments is divided successfully", MsgBoxStyle.Information, "TypeConvert")
        End If
    End Sub
    Public Sub ConvertGraphics(ByRef pMap As IMap)
        Dim pFeatureTypeConverter As New FeatureTypeConverter
        Dim I As Object
        Dim Y_Max As Object
        Dim Y_Min As Object
        Dim X_Max As Object
        Dim X_Min As Object
        Dim pActiveView As IActiveView
        Dim pEnvelope As IEnvelope
        Dim pExpOp As IExportOperation
        Dim pDSName As IDatasetName
        pExpOp = New ExportOperation
        Dim pSR As ISpatialReference
        Dim pGraphicsContainer As IGraphicsContainer
        Dim pGraphContSel As IGraphicsContainerSelect
        Dim pNewFeatureLayer As IFeatureLayer
        Dim hasSelection As Boolean
        Dim Exp_option As esriExportTableOptions
        Dim NewFeatureClasses As FeatureClassArray
        Exp_option = esriExportTableOptions.esriExportAllRecords

        Dim pElement As IElement
        pGraphicsContainer = pMap
        pGraphContSel = pGraphicsContainer
        Dim pGraphicElementEnum As IEnumElement
        pGraphicElementEnum = pGraphContSel.SelectedElements
        pGraphicElementEnum.Reset()
        pElement = pGraphicElementEnum.Next

        If pElement Is Nothing Then
            hasSelection = False
        Else
            hasSelection = True
        End If

        Dim pFeatureClass As IFeatureClass
        Dim pWSFactory As IWorkspaceFactory
        Dim pWS As IWorkspace
        pWSFactory = New ShapefileWorkspaceFactory
        pWS = pWSFactory.OpenFromFile(Environ("Temp"), 0)
        pSR = New UnknownCoordinateSystem
        pFeatureClass = CreateFC(pWS, "temp", esriGeometryType.esriGeometryPolygon, pSR)
        pDSName = pExpOp.GetOptions(pFeatureClass, "", hasSelection, False, 0, False, Exp_option)
        Dim pDataset As IDataset
        pDataset = pFeatureClass
        pDataset.Delete()
        pExpOp = Nothing

        If Not pDSName Is Nothing Then

            If Exp_option = esriExportTableOptions.esriExportAllRecords Then
                pGraphicElementEnum = Nothing
            End If

            If Exp_option = esriExportTableOptions.esriExportFeaturesWithinExtent Then
                pActiveView = pMap
                pEnvelope = pActiveView.Extent
                X_Min = pEnvelope.XMin
                X_Max = pEnvelope.XMax
                Y_Min = pEnvelope.YMin
                Y_Max = pEnvelope.YMax
                pGraphicsContainer.Reset()
                pElement = pGraphicsContainer.Next
                Do While Not pElement Is Nothing
                    pEnvelope = pElement.Geometry.Envelope
                    If pEnvelope.XMax <= X_Max And pEnvelope.XMin >= X_Min And pEnvelope.YMax <= Y_Max And pEnvelope.YMin >= Y_Min Then
                        pGraphContSel.SelectElement(pElement)
                    End If
                    pElement = pGraphicsContainer.Next
                Loop
                pGraphicElementEnum = pGraphContSel.SelectedElements
                pGraphicElementEnum.Reset()

            End If


            pSR = pMap.SpatialReference
            If pSR Is Nothing Then pSR = New UnknownCoordinateSystem

            pFeatureTypeConverter.DatasetName = pDSName
            pFeatureTypeConverter.SpatialReference = pSR
            NewFeatureClasses = pFeatureTypeConverter.ConvertGraphics(pGraphicsContainer, pGraphicElementEnum)

            If MsgBox("Do you want to add converted data to the map as new layer(s) ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Type converter") = MsgBoxResult.Yes Then
                With NewFeatureClasses
                    For I = 0 To UBound(.FeatureClass)
                        If Not .FeatureClass(I) Is Nothing Then
                            pNewFeatureLayer = New FeatureLayer
                            pNewFeatureLayer.FeatureClass = .FeatureClass(I)
                            pNewFeatureLayer.Name = .FeatureClass(I).AliasName
                            pMap.AddLayer(pNewFeatureLayer)
                        End If
                    Next I
                End With
            End If

            pActiveView = pMap
            pActiveView.Refresh()

        End If
    End Sub

    Private Function CreateFC(ByRef OutWS As IFeatureWorkspace, ByRef FCName As String, ByRef GeomType As esriGeometryType, ByRef pNewSR As ISpatialReference) As IFeatureClass
        Dim I As Object

        Dim sName As String
        Dim pFields As IFields
        Dim pFieldsEdit As IFieldsEdit
        pFields = New Fields
        pFieldsEdit = pFields
        pFieldsEdit.FieldCount_2 = 3

        Dim pField As IField
        Dim pFieldEdit As IFieldEdit

        pField = New Field
        pFieldEdit = pField

        With pFieldEdit
            .Name_2 = "OBJECTID"
            .AliasName_2 = "FID"
            .Type_2 = esriFieldType.esriFieldTypeOID
        End With

        pFieldsEdit.Field_2(0) = pField

        pField = New Field
        pFieldEdit = pField
        pFieldEdit.Name_2 = "Shape"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeGeometry
        Dim pGeomDef As IGeometryDef
        Dim pGeomDefEdit As IGeometryDefEdit
        pGeomDef = New GeometryDef
        pGeomDefEdit = pGeomDef
        pGeomDefEdit.GeometryType_2 = GeomType
        pGeomDefEdit.SpatialReference_2 = pNewSR

        pFieldEdit.GeometryDef_2 = pGeomDef

        pFieldsEdit.Field_2(1) = pField

        pField = New Field
        pFieldEdit = pField
        pFieldEdit.Name_2 = "Text"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldsEdit.Field_2(2) = pField

        Dim pFC As IFeatureClass
        On Error Resume Next

        pFC = OutWS.OpenFeatureClass(FCName)
        sName = FCName
        Do While (Err.Number = 0)
            I = I + 1
            sName = FCName & "_" & CStr(I)
            pFC = OutWS.OpenFeatureClass(sName)
        Loop

        pFC = OutWS.CreateFeatureClass(sName, pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, "Shape", "")

        CreateFC = pFC

    End Function

    Public Sub ExportToBLN(ByVal pFeatureLayer As IFeatureLayer, ByVal FileName As String, ByVal Outside As Boolean)
        Dim pFeatureTypeConverter As New FeatureTypeConverter
        pFeatureTypeConverter.ExportToBLN(pFeatureLayer.FeatureClass, FileName, Outside)
    End Sub

    Public Sub ConvertToKML(ByVal pFeatureLayer As IFeatureLayer, ByVal FileName As String, ByVal LabelFieldName As String, ByVal PMFlag As Boolean, Optional ByVal blnExtrude As Integer = 0, Optional ByVal AltitudeMode As Integer = 0, Optional ByVal AttributeFieldName As String = "", Optional ByVal SetValue As Double = 0, Optional ByVal Distribute As Boolean = False)
        Dim pFeatureTypeConverter As New FeatureTypeConverter
        pFeatureTypeConverter.ConvertToKML(pFeatureLayer, FileName, LabelFieldName, PMFlag, blnExtrude, AltitudeMode, AttributeFieldName, SetValue, Distribute)
    End Sub
End Module