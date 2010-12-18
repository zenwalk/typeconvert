Option Strict Off
Option Explicit On 
Namespace TypeConvert
    Public Interface IFeatureTypeConverter
        WriteOnly Property SpatialReference() As ISpatialReference
        WriteOnly Property DatasetName() As IDatasetName

        Function ToPolygon(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass
        Function ToPolyline(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass
        Function ToPoint(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass
        Function ToSegments(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass
        Function DivideSegments(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet, ByVal pMap As IMap, ByVal Lp As Double) As Boolean
        Function ToConvexHull(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass
        Function ToEnvelope(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass
        Function ToCentroid(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass
        Function Stratify(ByVal pFeatureLayer As IFeatureLayer, ByVal pSelSet As ISelectionSet) As FeatureLayerArray
        Function RemoveDuplicates(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As FeatureClassArray
        Function ConvertGraphics(ByVal pGraphicsContainer As IGraphicsContainer, ByVal pGraphicElementEnum As IEnumElement) As FeatureClassArray
        Sub ExportToBLN(ByVal pFeatureClass As IFeatureClass, ByVal FileName As String, ByVal Outside As Boolean)
        Sub ConvertToKML(ByVal pFeatureLayer As IFeatureLayer, ByVal FileName As String, ByVal LabelFieldName As String, ByVal PMFlag As Boolean, Optional ByVal blnExtrude As Integer = 0, Optional ByVal AltitudeMode As Integer = 0, Optional ByVal AttributeFieldName As String = "", Optional ByVal SetValue As Double = 0, Optional ByVal Distribute As Boolean = False)
    End Interface
    Public Structure FeatureClassArray
        Dim FeatureClass() As IFeatureClass
    End Structure
    Public Structure FeatureLayerArray
        Dim FeatureLayer() As IFeatureLayer
    End Structure
    <Guid("43A92D5A-DC09-4161-B911-5F9D5ECD15DA")> _
    Public Class FeatureTypeConverter
        Implements IFeatureTypeConverter

        Dim pTrackCancel As ITrackCancel ' Прогрессор
        Dim pPDlgFact As IProgressDialogFactory
        Dim pStepPro As IStepProgressor
        Dim bContinue As Boolean
        Dim iPosition As Integer
        Dim iFeatureCount As Integer
        Dim pPDlg As IProgressDialog2

        Dim pOutWorkspace As IFeatureWorkspace
        Dim pNewSR As ISpatialReference = New UnknownCoordinateSystem
        Dim pDSName As IDatasetName
        Public WriteOnly Property DatasetName() As IDatasetName Implements IFeatureTypeConverter.DatasetName
            Set(ByVal Value As IDatasetName)
                Dim pWorkspace As IWorkspace
                Dim pName As IName
                pDSName = Value
                pName = pDSName.WorkspaceName
                pWorkspace = pName.Open
                pOutWorkspace = pWorkspace
            End Set
        End Property
        Public WriteOnly Property SpatialReference() As ISpatialReference Implements IFeatureTypeConverter.SpatialReference
            Set(ByVal Value As ISpatialReference)
                pNewSR = Value
            End Set
        End Property

        Private Sub StartProgress(ByVal MaxRange As Long)
            iPosition = 0
            pTrackCancel = New ESRI.ArcGIS.Display.CancelTracker
            pPDlgFact = New ProgressDialogFactory
            pPDlg = pPDlgFact.Create(pTrackCancel, 0)
            pPDlg.CancelEnabled = True
            pPDlg.Animation = esriProgressAnimationTypes.esriDownloadFile
            pPDlg.Title = "Converting..."
            pStepPro = pPDlg
            pStepPro.MinRange = 0
            pStepPro.MaxRange = MaxRange
            pStepPro.StepValue = 1

        End Sub
        Public Function ToPolygon(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass Implements IFeatureTypeConverter.ToPolygon
            Select Case pFeatureClass.ShapeType

                Case esriGeometryType.esriGeometryPolygon, esriGeometryType.esriGeometryPolyline

                    Return PolyToPoly(pFeatureClass, pSelSet, esriGeometryType.esriGeometryPolygon)

                Case esriGeometryType.esriGeometryPoint

                    Return PointToPoly(pFeatureClass, pSelSet, True)
                Case Else
                    Return Nothing
            End Select
        End Function
        Public Function ToPolyline(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass Implements IFeatureTypeConverter.ToPolyline

            Select Case pFeatureClass.ShapeType

                Case esriGeometryType.esriGeometryPolygon, esriGeometryType.esriGeometryPolyline

                    Return PolyToPoly(pFeatureClass, pSelSet, esriGeometryType.esriGeometryPolyline)

                Case esriGeometryType.esriGeometryPoint

                    Return PointToPoly(pFeatureClass, pSelSet, False)
                Case Else
                    Return Nothing
            End Select
        End Function
        Public Function ToPoint(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass Implements IFeatureTypeConverter.ToPoint
            Select Case pFeatureClass.ShapeType

                Case esriGeometryType.esriGeometryPolygon, esriGeometryType.esriGeometryPolyline

                    Return PolyToPoint(pFeatureClass, pSelSet)

                Case esriGeometryType.esriGeometryPoint

                    Return PointToPoint(pFeatureClass, pSelSet)

                Case Else
                    Return Nothing
            End Select
        End Function

        Private Function PolyToPoly(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet, ByVal GeomType As esriGeometryType) As IFeatureClass
            Dim I As Short
            Dim J As Short
            Dim ShpFieldIndex As Short

            Dim pNewFeatureClass As IFeatureClass
            Dim pDataset As IDataset
            Dim pGeometry As IGeometry
            Dim pPointColl As IPointCollection

            pDataset = pFeatureClass

            Dim pNewFields As IFields
            pNewFields = fnMakeFields(pFeatureClass.Fields, (pFeatureClass.ShapeFieldName), GeomType)
            If pNewFields Is Nothing Then Return Nothing
            If Not pNewFields.FieldCount = pFeatureClass.Fields.FieldCount Then Return Nothing

            On Error Resume Next
            Err.Clear()

            pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name(), pNewFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Do While (Err.Number = fdoError.FDO_E_NO_PERMISSION) Or (Err.Number = fdoError.FDO_E_TABLE_ALREADY_EXISTS)
                Err.Clear()
                I = I + 1
                pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name() & "_" & CStr(I), pNewFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Loop

            ' Start editing the FeatureClass we have created.

            Dim pFeatureCursor As IFeatureCursor
            Dim pFeature As IFeature
            If Not pSelSet Is Nothing Then
                pSelSet.Search(Nothing, False, pFeatureCursor)
                iFeatureCount = pSelSet.Count
            Else
                pFeatureCursor = pFeatureClass.Search(Nothing, False)
                iFeatureCount = pFeatureClass.FeatureCount(Nothing)
            End If
            pFeature = pFeatureCursor.NextFeature

            Dim pCursor As IFeatureCursor
            pCursor = pNewFeatureClass.Insert(True)
            Dim pBuffer As IFeatureBuffer
            pBuffer = pNewFeatureClass.CreateFeatureBuffer
            Dim pVertices As IPointCollection

            StartProgress((iFeatureCount))

            Dim pZAware As IZAware
            Dim pMAware As IMAware
            Do While Not pFeature Is Nothing

                For I = 0 To pNewFields.FieldCount - 1
                    If (Not pNewFields.Field(I).Type = esriFieldType.esriFieldTypeGeometry) And (Not pNewFields.Field(I).Type = esriFieldType.esriFieldTypeOID) Then
                        If pNewFields.Field(I).Editable Then
                            J = pFeature.Fields.FindField(pNewFields.Field(I).Name)
                            If J <> -1 Then
                                pBuffer.Value(I) = pFeature.Value(J)
                            End If
                        End If
                    End If
                    If pNewFields.Field(I).Type = esriFieldType.esriFieldTypeGeometry Then
                        ShpFieldIndex = I
                    End If
                Next I


                pPointColl = pFeature.ShapeCopy

                If GeomType = esriGeometryType.esriGeometryPolygon Then
                    pVertices = New Polygon
                Else
                    pVertices = New Polyline
                End If

                For I = 0 To pPointColl.PointCount - 1
                    pVertices.AddPoint(pPointColl.Point(I))
                Next I

                If GeomType = esriGeometryType.esriGeometryPolygon And pFeatureClass.ShapeType <> esriGeometryType.esriGeometryPolygon Then
                    pVertices.AddPoint(pVertices.Point(0))
                End If

                If pFeature.Fields.Field(ShpFieldIndex).GeometryDef.HasZ Then
                    pZAware = pVertices
                    pZAware.ZAware = True
                End If
                If pFeature.Fields.Field(ShpFieldIndex).GeometryDef.HasM Then
                    pMAware = pVertices
                    pZAware.ZAware = True
                End If

                pBuffer.Shape = pVertices
                pGeometry = pBuffer.Shape
                pGeometry.Project(pNewSR)
                pBuffer.Shape = pGeometry
                pCursor.InsertFeature(pBuffer)

                pFeature = pFeatureCursor.NextFeature
                iPosition = iPosition + 1
                pStepPro.Position = iPosition
                pStepPro.Message = "Converting feature: " & iPosition & " of " & iFeatureCount
                bContinue = pTrackCancel.[Continue]
                If Not bContinue Then GoTo exit_Sub
            Loop

exit_Sub:

            pCursor.Dispose()
            pPDlg.HideDialog()

            Return pNewFeatureClass

        End Function


        Private Function PolyToPoint(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass

            Dim I As Short
            'On Error GoTo errHand
            On Error Resume Next
            Dim pNewFeatureClass As IFeatureClass
            Dim pDataset As IDataset
            pDataset = pFeatureClass

            Dim pGeoDataset As IGeoDataset
            pGeoDataset = pFeatureClass

            Dim pFCursor As IFeatureCursor

            If Not pSelSet Is Nothing Then
                pSelSet.Search(Nothing, False, pFCursor)
                iFeatureCount = pSelSet.Count
            Else
                pFCursor = pFeatureClass.Search(Nothing, False)
                iFeatureCount = pFeatureClass.FeatureCount(Nothing)
            End If

            'Check if the line is MAware or ZAware
            Dim lGeomIndex As Integer
            Dim sShpName As String
            Dim pFieldsTest As IFields
            Dim pFieldTest As IField
            Dim pGeometryDefTest As IGeometryDef
            sShpName = pFeatureClass.ShapeFieldName
            pFieldsTest = pFeatureClass.Fields
            lGeomIndex = pFieldsTest.FindField(sShpName)
            pFieldTest = pFieldsTest.Field(lGeomIndex)
            pGeometryDefTest = pFieldTest.GeometryDef
            Dim bZAware As Boolean
            Dim bMAware As Boolean
            'Determine if M or Z aware
            bZAware = pGeometryDefTest.HasZ
            bMAware = pGeometryDefTest.HasM

            Dim pFields As IFields
            Dim pFieldsEdit As IFieldsEdit
            pFields = New Fields
            pFieldsEdit = pFields
            pFieldsEdit.FieldCount_2 = 2

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
            'Define the geometry field according to the input attribut awareness

            With pGeomDefEdit
                .GeometryType_2 = esriGeometryType.esriGeometryPoint
                .HasM_2 = bMAware
                .HasZ_2 = bZAware
                .SpatialReference_2 = pNewSR 'New UnknownCoordinateSystem
            End With

            pFieldEdit.GeometryDef_2 = pGeomDef

            pFieldsEdit.Field_2(1) = pField


            On Error Resume Next
            Err.Clear()
            pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name(), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Do While (Err.Number = fdoError.FDO_E_NO_PERMISSION) Or (Err.Number = fdoError.FDO_E_TABLE_ALREADY_EXISTS)
                Err.Clear()
                I = I + 1
                pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name() & "_" & CStr(I), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Loop

            'On Error GoTo errHand
            Dim pExistingFields As IFields
            'Variable for storing
            pExistingFields = pNewFeatureClass.Fields
            Dim lFieldCount As Integer
            lFieldCount = pFeatureClass.Fields.FieldCount
            Dim pFieldsIn As IFields
            pFieldsIn = pFeatureClass.Fields

            Dim J As Integer
            Dim bFlag As Boolean

            'Verify if the necessary fields are in the output feature class if not it add those
            For I = 0 To lFieldCount - 1
                bFlag = False
                If pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeOID And pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeGeometry And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_LENGTH" And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_AREA" Then
                    For J = 0 To pExistingFields.FieldCount - 1
                        If pExistingFields.Field(J).Name = pFieldsIn.Field(I).Name Then
                            bFlag = True
                            Exit For
                        End If
                    Next
                    If bFlag = False Then
                        If I = 14 Then
                            System.Diagnostics.Debug.WriteLine("err")
                        End If
                        pNewFeatureClass.AddField(pFieldsIn.Field(I))
                    End If
                End If
            Next

            If pNewFeatureClass.Fields.FindField("gsGROUP") = -1 Then
                pField = New Field
                pFieldEdit = pField

                With pFieldEdit
                    .Name_2 = "gsGROUP"
                    .AliasName_2 = "Group"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With

                pNewFeatureClass.AddField(pField)
            End If

            If pNewFeatureClass.Fields.FindField("gsORDER") = -1 Then
                pField = New Field
                pFieldEdit = pField

                With pFieldEdit
                    .Name_2 = "gsORDER"
                    .AliasName_2 = "Order"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With

                pNewFeatureClass.AddField(pField)
            End If

            Dim pGeometry As IGeometry
            Dim pFeatureBuffer As IFeatureBuffer
            Dim pNewFCursor As IFeatureCursor
            pNewFCursor = pNewFeatureClass.Insert(True)
            pFeatureBuffer = pNewFeatureClass.CreateFeatureBuffer
            Dim pFeatureLine As IFeature
            pFeatureLine = pFCursor.NextFeature
            Dim pPtColl As IPointCollection
            Dim pEnunVertices As IEnumVertex
            Dim pPt As IPoint
            pPt = New ESRI.ArcGIS.Geometry.Point
            Dim lPartIndex As Integer
            Dim lVertexIndex As Integer
            Dim OIDIndex As Short
            Dim GeomCount As Short
            Dim GroupCount As Short

            StartProgress((iFeatureCount))
            GroupCount = 0
            iPosition = 1
            Dim pGeomColl As IGeometryCollection
            While Not pFeatureLine Is Nothing

                pPDlg.Description = "Converting feature: " & iPosition & " of " & iFeatureCount
                iPosition = iPosition + 1
                pStepPro.Position = iPosition

                pGeomColl = pFeatureLine.ShapeCopy
                GroupCount = GroupCount + 1
                For GeomCount = 0 To pGeomColl.GeometryCount - 1

                    pPtColl = pGeomColl.Geometry(GeomCount)
                    pGeometry = pGeomColl.Geometry(GeomCount)

                    pGeometry.Project(pNewSR)
                    pPtColl = pGeometry
                    pEnunVertices = pPtColl.EnumVertices
                    pEnunVertices.QueryNext(pPt, lPartIndex, lVertexIndex)
                    While Not pPt.IsEmpty
                        pFeatureBuffer.Shape = pPt
                        For I = 0 To lFieldCount - 1
                            If pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeOID And pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeGeometry And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_LENGTH" And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_AREA" Then
                                J = pFeatureBuffer.Fields.FindField(pFieldsIn.Field(I).Name)
                                If J <> -1 Then
                                    pFeatureBuffer.Value(J) = pFeatureLine.Value(I)
                                End If
                            End If
                            If pFieldsIn.Field(I).Type = esriFieldType.esriFieldTypeOID Then
                                OIDIndex = I
                            End If
                        Next


                        pFeatureBuffer.Value(pFeatureBuffer.Fields.FindField("gsGROUP")) = GroupCount
                        pFeatureBuffer.Value(pFeatureBuffer.Fields.FindField("gsORDER")) = lVertexIndex + 1

                        pNewFCursor.InsertFeature(pFeatureBuffer)

                        pStepPro.Message = "Vertex: " & (lVertexIndex + 1) & " of " & pPtColl.PointCount
                        pEnunVertices.QueryNext(pPt, lPartIndex, lVertexIndex)

                    End While
                    If GeomCount < (pGeomColl.GeometryCount - 1) Then GroupCount = GroupCount + 1
                Next GeomCount

                pFeatureLine = pFCursor.NextFeature

                bContinue = pTrackCancel.[Continue]
                If Not bContinue Then GoTo exit_Sub
            End While

exit_Sub:


            pPDlg.HideDialog()
            pNewFCursor.Dispose()

            Return pNewFeatureClass

            Exit Function
errHand:
            MsgBox(Err.Description)
        End Function


        Private Function PointToPoly(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet, ByVal ToPolygon As Boolean) As IFeatureClass
            Dim I As Short

            Dim pField As IField
            Dim pFieldEdit As IFieldEdit
            Dim fGroupField As Short
            Dim fOrderField As Short


            fGroupField = pFeatureClass.Fields.FindField("gsGroup")
            fOrderField = pFeatureClass.Fields.FindField("gsOrder")
            If fGroupField = -1 Then
                If MsgBox("Group field not found. Create it ?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Group field not found") = MsgBoxResult.Yes Then
                    pField = New Field
                    pFieldEdit = pField
                    With pFieldEdit
                        .Name_2 = "gsGroup"
                        .AliasName_2 = "Group"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureClass.AddField. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    pFeatureClass.AddField(pField)
                End If
            End If
            If fOrderField = -1 Then
                If MsgBox("Order field not found. Create it ?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Order field not found") = MsgBoxResult.Yes Then
                    pField = New Field
                    pFieldEdit = pField
                    With pFieldEdit
                        .Name_2 = "gsOrder"
                        .AliasName_2 = "Order"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureClass.AddField. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                    pFeatureClass.AddField(pField)
                End If
            End If

            If (fOrderField = -1) Or (fGroupField = -1) Then Return Nothing

            Dim pFields As IFields
            Dim pFieldsEdit As IFieldsEdit
            pFields = New Fields
            pFieldsEdit = pFields
            pFieldsEdit.FieldCount_2 = 3

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
            If ToPolygon Then
                pGeomDefEdit.GeometryType_2 = esriGeometryType.esriGeometryPolygon
            Else
                pGeomDefEdit.GeometryType_2 = esriGeometryType.esriGeometryPolyline
            End If
            pGeomDefEdit.SpatialReference_2 = pNewSR
            'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureClass.Fields. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            pGeomDefEdit.HasM_2 = pFeatureClass.Fields.Field(pFeatureClass.Fields.FindField(pFeatureClass.ShapeFieldName)).GeometryDef.HasM
            'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureClass.Fields. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            pGeomDefEdit.HasZ_2 = pFeatureClass.Fields.Field(pFeatureClass.Fields.FindField(pFeatureClass.ShapeFieldName)).GeometryDef.HasZ

            pFieldEdit.GeometryDef_2 = pGeomDef
            pFieldsEdit.Field_2(1) = pField

            pField = New Field
            pFieldEdit = pField

            With pFieldEdit
                .Name_2 = "gsGroup"
                .AliasName_2 = "Group"
                .Type_2 = esriFieldType.esriFieldTypeInteger
            End With
            pFieldsEdit.Field_2(2) = pField

            Dim pNewFeatureClass As IFeatureClass

            On Error Resume Next
            Err.Clear()
            pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name(), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Do While (Err.Number = fdoError.FDO_E_NO_PERMISSION) Or (Err.Number = fdoError.FDO_E_TABLE_ALREADY_EXISTS)
                Err.Clear()
                I = I + 1
                pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name() & "_" & CStr(I), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Loop

            Dim pGeometry As IGeometry
            Dim pPoint As IPoint
            Dim pVertices As IPointCollection
            Dim pFeature As IFeature
            Dim pTable As ITable
            Dim pRow As IRow
            Dim pPointRow As IRow
            Dim pPointCursor As ICursor
            Dim pFeatureCursor As IFeatureCursor
            Dim pQF As IQueryFilter
            Dim UniqueValues() As Object
            Dim pUniqueCursor As ICursor

            If pSelSet Is Nothing Then
                pUniqueCursor = pFeatureClass.Search(Nothing, False)
            Else
                pSelSet.Search(Nothing, False, pUniqueCursor)
            End If

            StartProgress((0))
            pPDlg.Description = "Processing groups. Please wait..."
            'UPGRADE_WARNING: Couldn't resolve default property of object UniqueValuesArray(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'

            UniqueValues = UniqueValuesArray(pUniqueCursor, "gsGroup")
            pPDlg.HideDialog()

            Dim pTableSort As ITableSort
            pTableSort = New TableSort
            pTable = pFeatureClass
            StartProgress((UBound(UniqueValues)))

            Dim pCursor As IFeatureCursor
            Dim pBuffer As IFeatureBuffer
            Dim pFirstPoint As IPoint
            Dim VertNumber As Integer
            Dim pZAware As IZAware
            Dim pMAware As IMAware
            For I = 0 To UniqueValues.Length - 1

                pPDlg.Description = "Converting point's group: " & I + 1 & " of " & UBound(UniqueValues)
                iPosition = iPosition + 1
                pStepPro.Position = iPosition
                bContinue = pTrackCancel.[Continue]
                If Not bContinue Then GoTo exit_Sub

                If Not IsDBNull(UniqueValues(I)) Then
                    pQF = New QueryFilter
                    If Not UniqueValues(I) Is Nothing Then
                        pQF.WhereClause = "gsGroup=" & CStr(UniqueValues(I))
                    Else
                        pQF.WhereClause = "gsGroup IS NULL"
                    End If

                    With pTableSort
                        .QueryFilter = pQF
                        If pSelSet Is Nothing Then
                            .Table = pTable
                        Else
                            .SelectionSet = pSelSet
                        End If
                        .Fields = "gsOrder"
                        .Ascending("gsOrder") = True
                    End With
                    pTableSort.Sort(Nothing)

                    pPointCursor = pTableSort.Rows
                    pPointRow = pPointCursor.NextRow
                    If Not pPointRow Is Nothing Then
                        pCursor = pNewFeatureClass.Insert(True)
                        pBuffer = pNewFeatureClass.CreateFeatureBuffer
                        If ToPolygon Then
                            pVertices = New Polygon
                        Else
                            pVertices = New Polyline
                        End If

                        pFeature = pPointRow
                        pFirstPoint = pFeature.ShapeCopy

                        VertNumber = 0
                        Do Until pPointRow Is Nothing

                            pFeature = pPointRow
                            pPoint = pFeature.ShapeCopy
                            pVertices.AddPoint(pPoint)
                            VertNumber = VertNumber + 1
                            'UPGRADE_WARNING: Couldn't resolve default property of object pStepPro.Message. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            pStepPro.Message = "Vertex: " & VertNumber
                            pPointRow = pPointCursor.NextRow

                        Loop

                        If ToPolygon Then
                            pVertices.AddPoint(pFirstPoint, , pVertices.PointCount)
                        End If

                        If pGeomDef.HasZ Then
                            pZAware = pVertices
                            pZAware.ZAware = True
                        End If
                        If pGeomDef.HasM Then
                            pMAware = pVertices
                            pMAware.MAware = True
                        End If

                        pBuffer.Shape = pVertices
                        pGeometry = pBuffer.Shape
                        pGeometry.Project(pNewSR)
                        pBuffer.Shape = pGeometry

                        'UPGRADE_WARNING: Couldn't resolve default property of object pBuffer.Fields. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object pBuffer.value. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                        pBuffer.Value(pBuffer.Fields.FindField("gsGroup")) = UniqueValues(I)

                        pCursor.InsertFeature(pBuffer)
                    End If
                End If

            Next I
exit_Sub:
            pPDlg.HideDialog()
            pCursor.Dispose()

            Return pNewFeatureClass
        End Function
        Private Function PointToPoint(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass
            PointToPoint = ExportFC(pDSName.Name, pFeatureClass, pSelSet)
        End Function
        Public Function ToSegments(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass Implements IFeatureTypeConverter.ToSegments

            Dim I As Short
            'On Error GoTo errHand
            On Error Resume Next
            Dim pNewFeatureClass As IFeatureClass
            Dim pDataset As IDataset
            pDataset = pFeatureClass

            Dim pGeoDataset As IGeoDataset
            pGeoDataset = pFeatureClass

            Dim pFCursor As IFeatureCursor

            If Not pSelSet Is Nothing Then
                pSelSet.Search(Nothing, False, pFCursor)
                iFeatureCount = pSelSet.Count
            Else
                pFCursor = pFeatureClass.Search(Nothing, False)
                iFeatureCount = pFeatureClass.FeatureCount(Nothing)
            End If

            'Check if the line is MAware or ZAware
            Dim lGeomIndex As Integer
            Dim sShpName As String
            Dim pFieldsTest As IFields
            Dim pFieldTest As IField
            Dim pGeometryDefTest As IGeometryDef
            sShpName = pFeatureClass.ShapeFieldName
            pFieldsTest = pFeatureClass.Fields
            lGeomIndex = pFieldsTest.FindField(sShpName)
            pFieldTest = pFieldsTest.Field(lGeomIndex)
            pGeometryDefTest = pFieldTest.GeometryDef
            Dim bZAware As Boolean
            Dim bMAware As Boolean
            'Determine if M or Z aware
            bZAware = pGeometryDefTest.HasZ
            bMAware = pGeometryDefTest.HasM
            'Create a new shapefile

            Dim pFields As IFields
            Dim pFieldsEdit As IFieldsEdit
            pFields = New Fields
            pFieldsEdit = pFields
            pFieldsEdit.FieldCount_2 = 2

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
            'Define the geometry field according to the input attribut awareness
            If (bMAware = False) And (bZAware = False) Then
                With pGeomDefEdit
                    .GeometryType_2 = esriGeometryType.esriGeometryPolyline
                    .SpatialReference_2 = pNewSR 'New UnknownCoordinateSystem
                End With
            ElseIf (bMAware = True) And (bZAware = False) Then
                With pGeomDefEdit
                    .GeometryType_2 = esriGeometryType.esriGeometryPolyline
                    .SpatialReference_2 = pNewSR 'New UnknownCoordinateSystem
                    .HasM_2 = True
                End With
            ElseIf (bMAware = False) And (bZAware = True) Then
                With pGeomDefEdit
                    .GeometryType_2 = esriGeometryType.esriGeometryPolyline
                    .SpatialReference_2 = pNewSR 'New UnknownCoordinateSystem
                    .HasZ_2 = True
                End With
            ElseIf (bMAware = True) And (bZAware = True) Then
                With pGeomDefEdit
                    .GeometryType_2 = esriGeometryType.esriGeometryPolyline
                    .SpatialReference_2 = pNewSR 'New UnknownCoordinateSystem
                    .HasZ_2 = True
                    .HasM_2 = True
                End With
            End If
            pFieldEdit.GeometryDef_2 = pGeomDef

            pFieldsEdit.Field_2(1) = pField


            On Error Resume Next
            Err.Clear()
            pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name(), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Do While (Err.Number = fdoError.FDO_E_NO_PERMISSION) Or (Err.Number = fdoError.FDO_E_TABLE_ALREADY_EXISTS)
                Err.Clear()
                I = I + 1
                pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name() & "_" & CStr(I), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Loop

            'On Error GoTo errHand
            Dim pExistingFields As IFields
            'Variable for storing
            pExistingFields = pNewFeatureClass.Fields
            Dim lFieldCount As Integer
            lFieldCount = pFeatureClass.Fields.FieldCount
            Dim pFieldsIn As IFields
            pFieldsIn = pFeatureClass.Fields

            Dim J As Integer
            Dim bFlag As Boolean

            'Verify if the necessary fields are in the output feature class if not it add those
            For I = 0 To lFieldCount - 1
                bFlag = False
                If pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeOID And pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeGeometry And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_LENGTH" And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_AREA" Then
                    For J = 0 To pExistingFields.FieldCount - 1
                        If pExistingFields.Field(J).Name = pFieldsIn.Field(I).Name Then
                            bFlag = True
                            Exit For
                        End If
                    Next
                    If bFlag = False Then
                        If I = 14 Then
                            System.Diagnostics.Debug.WriteLine("err")
                        End If
                        pNewFeatureClass.AddField(pFieldsIn.Field(I))
                    End If
                End If
            Next

            If pNewFeatureClass.Fields.FindField("gsGROUP") = -1 Then
                pField = New Field
                pFieldEdit = pField

                With pFieldEdit
                    .Name_2 = "gsGROUP"
                    .AliasName_2 = "Group"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With

                pNewFeatureClass.AddField(pField)
            End If

            If pNewFeatureClass.Fields.FindField("gsORDER") = -1 Then
                pField = New Field
                pFieldEdit = pField

                With pFieldEdit
                    .Name_2 = "gsORDER"
                    .AliasName_2 = "Order"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With

                pNewFeatureClass.AddField(pField)
            End If

            Dim pGeometry As IGeometry
            Dim pFeatureBuffer As IFeatureBuffer
            Dim pNewFCursor As IFeatureCursor
            pNewFCursor = pNewFeatureClass.Insert(True)
            pFeatureBuffer = pNewFeatureClass.CreateFeatureBuffer
            Dim pFeature As IFeature
            pFeature = pFCursor.NextFeature
            Dim pSegColl As ISegmentCollection
            Dim pEnumSegment As IEnumSegment
            Dim pSegment As ISegment

            Dim outPartIndex As Integer
            Dim SegmentIndex As Integer
            Dim OIDIndex As Short
            Dim GeomCount As Short
            Dim GroupCount As Short

            StartProgress((iFeatureCount))
            GroupCount = 0
            iPosition = 1
            Dim pGeomColl As IGeometryCollection
            Dim pPointColl As IPointCollection
            Dim pZAware As IZAware
            Dim pMAware As IMAware
            While Not pFeature Is Nothing

                pPDlg.Description = "Converting feature: " & iPosition & " of " & iFeatureCount
                iPosition = iPosition + 1
                pStepPro.Position = iPosition

                pGeomColl = pFeature.ShapeCopy
                GroupCount = GroupCount + 1
                For GeomCount = 0 To pGeomColl.GeometryCount - 1

                    pSegColl = pGeomColl.Geometry(GeomCount)
                    pGeometry = pGeomColl.Geometry(GeomCount)

                    pGeometry.Project(pNewSR)
                    pSegColl = pGeometry
                    pEnumSegment = pSegColl.EnumSegments
                    pEnumSegment.Next(pSegment, outPartIndex, SegmentIndex)

                    While Not pSegment Is Nothing

                        pPointColl = New Polyline
                        pPointColl.AddPoint(pSegment.FromPoint)
                        pPointColl.AddPoint(pSegment.ToPoint)
                        If pGeomDef.HasZ Then
                            pZAware = pPointColl
                            pZAware.ZAware = True
                        End If
                        If pGeomDef.HasM Then
                            pMAware = pPointColl
                            pMAware.MAware = True
                        End If

                        pFeatureBuffer.Shape = pPointColl
                        For I = 0 To lFieldCount - 1
                            If pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeOID And pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeGeometry And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_LENGTH" And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_AREA" Then
                                J = pFeatureBuffer.Fields.FindField(pFieldsIn.Field(I).Name)
                                If J <> -1 Then
                                    pFeatureBuffer.Value(J) = pFeature.Value(I)
                                End If
                            End If
                            If pFieldsIn.Field(I).Type = esriFieldType.esriFieldTypeOID Then
                                OIDIndex = I
                            End If
                        Next
                        pFeatureBuffer.Value(pFeatureBuffer.Fields.FindField("gsGROUP")) = GroupCount
                        pFeatureBuffer.Value(pFeatureBuffer.Fields.FindField("gsORDER")) = SegmentIndex + 1

                        pNewFCursor.InsertFeature(pFeatureBuffer)
                        pStepPro.Message = "Segment: " & (SegmentIndex + 1) & " of " & pSegColl.SegmentCount
                        pEnumSegment.Next(pSegment, outPartIndex, SegmentIndex)

                    End While
                    If GeomCount < (pGeomColl.GeometryCount - 1) Then GroupCount = GroupCount + 1
                Next GeomCount

                pFeature = pFCursor.NextFeature

                bContinue = pTrackCancel.[Continue]
                If Not bContinue Then GoTo exit_Sub
            End While

exit_Sub:


            pPDlg.HideDialog()
            pNewFCursor.Dispose()

            Return pNewFeatureClass

            Exit Function
errHand:
            MsgBox(Err.Description)
        End Function

        Public Function ToConvexHull(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass Implements IFeatureTypeConverter.ToConvexHull

            Dim I As Short
            Dim pFields As IFields
            Dim pFieldsEdit As IFieldsEdit
            Dim pField As IField
            Dim pFieldEdit As IFieldEdit
            pFields = New Fields
            pFieldsEdit = pFields
            pFieldsEdit.FieldCount_2 = 2

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
            'If ToPolygon Then
            pGeomDefEdit.GeometryType_2 = esriGeometryType.esriGeometryPolygon
            'Else
            'pGeomDefEdit.GeometryType = esriGeometryPolyline
            'End If
            pGeomDefEdit.SpatialReference_2 = pNewSR
            pFieldEdit.GeometryDef_2 = pGeomDef
            pFieldsEdit.Field_2(1) = pField

            Dim pPolygon As IPolygon
            Dim pPolygonPointColl As IPointCollection
            Dim pTopoOper As ITopologicalOperator
            Dim pFeatureCursor As IFeatureCursor
            Dim pFeature As IFeature
            If Not pSelSet Is Nothing Then
                pSelSet.Search(Nothing, False, pFeatureCursor)
            Else
                pFeatureCursor = pFeatureClass.Search(Nothing, False)
            End If
            pFeature = pFeatureCursor.NextFeature

            pTopoOper = pFeature.Shape
            pFeature = pFeatureCursor.NextFeature

            Do While Not pFeature Is Nothing

                pTopoOper = pTopoOper.Union(pFeature.Shape)
                pFeature = pFeatureCursor.NextFeature

            Loop

            pTopoOper = pTopoOper.ConvexHull

            'Create a new polygon and call Close to ensure a closed polygon
            pPolygonPointColl = New Polygon
            pPolygonPointColl.AddPointCollection(pTopoOper)
            pPolygon = pPolygonPointColl 'QI
            pPolygon.Close()

            Dim pNewFeatureClass As IFeatureClass

            On Error Resume Next
            Err.Clear()
            pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name(), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Do While (Err.Number = fdoError.FDO_E_NO_PERMISSION) Or (Err.Number = fdoError.FDO_E_TABLE_ALREADY_EXISTS)
                Err.Clear()
                I = I + 1
                pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name() & "_" & CStr(I), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Loop

            Dim pNewDataset As IDataset
            Dim pNewWorkspaceEdit As IWorkspaceEdit
            pNewDataset = pNewFeatureClass
            pNewWorkspaceEdit = pNewDataset.Workspace
            pNewWorkspaceEdit.StartEditing(False)
            pNewWorkspaceEdit.StartEditOperation()

            Dim pNewFeature As IFeature
            Dim pGeometry As IGeometry
            pNewFeature = pNewFeatureClass.CreateFeature
            pNewFeature.Shape = pPolygon
            pGeometry = pNewFeature.ShapeCopy
            pGeometry.Project(pNewSR)
            pNewFeature.Shape = pGeometry
            pNewFeature.Store()

            pNewWorkspaceEdit.StopEditOperation()
            pNewWorkspaceEdit.StopEditing(True)

            ToConvexHull = pNewFeatureClass

        End Function
        Public Function ToEnvelope(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass Implements IFeatureTypeConverter.ToEnvelope

            Dim I As Short
            Dim pFields As IFields
            Dim pFieldsEdit As IFieldsEdit
            Dim pField As IField
            Dim pFieldEdit As IFieldEdit
            pFields = New Fields
            pFieldsEdit = pFields
            pFieldsEdit.FieldCount_2 = 2

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
            'If ToPolygon Then
            pGeomDefEdit.GeometryType_2 = esriGeometryType.esriGeometryPolygon
            'Else
            'pGeomDefEdit.GeometryType = esriGeometryPolyline
            'End If
            pGeomDefEdit.SpatialReference_2 = pNewSR
            pFieldEdit.GeometryDef_2 = pGeomDef
            pFieldsEdit.Field_2(1) = pField

            Dim pPolygon As IPolygon
            Dim pPolygonPointColl As IPointCollection
            Dim pTopoOper As ITopologicalOperator
            Dim pFeatureCursor As IFeatureCursor
            Dim pFeature As IFeature
            If Not pSelSet Is Nothing Then
                pSelSet.Search(Nothing, False, pFeatureCursor)
            Else
                pFeatureCursor = pFeatureClass.Search(Nothing, False)
            End If
            pFeature = pFeatureCursor.NextFeature

            pTopoOper = pFeature.Shape
            pFeature = pFeatureCursor.NextFeature

            Do While Not pFeature Is Nothing

                pTopoOper = pTopoOper.Union(pFeature.Shape)
                pFeature = pFeatureCursor.NextFeature

            Loop

            Dim pGeometry As IGeometry
            Dim pEnvelope As IEnvelope
            pGeometry = pTopoOper
            pEnvelope = pGeometry.Envelope

            'Create a new polygon and call Close to ensure a closed polygon
            pPolygonPointColl = New Polygon
            Dim pPoint As IPoint
            pPoint = New ESRI.ArcGIS.Geometry.Point
            pPoint.X = pEnvelope.XMin
            pPoint.Y = pEnvelope.YMin
            pPolygonPointColl.AddPoint(pPoint)
            pPoint.X = pEnvelope.XMax
            pPoint.Y = pEnvelope.YMin
            pPolygonPointColl.AddPoint(pPoint)
            pPoint.X = pEnvelope.XMax
            pPoint.Y = pEnvelope.YMax
            pPolygonPointColl.AddPoint(pPoint)
            pPoint.X = pEnvelope.XMin
            pPoint.Y = pEnvelope.YMax
            pPolygonPointColl.AddPoint(pPoint)

            pPolygon = pPolygonPointColl 'QI
            pPolygon.Close()

            Dim pNewFeatureClass As IFeatureClass

            On Error Resume Next
            Err.Clear()
            pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name(), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Do While (Err.Number = fdoError.FDO_E_NO_PERMISSION) Or (Err.Number = fdoError.FDO_E_TABLE_ALREADY_EXISTS)
                Err.Clear()
                I = I + 1
                pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name() & "_" & CStr(I), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Loop

            Dim pNewDataset As IDataset
            Dim pNewWorkspaceEdit As IWorkspaceEdit
            pNewDataset = pNewFeatureClass
            pNewWorkspaceEdit = pNewDataset.Workspace
            pNewWorkspaceEdit.StartEditing(False)
            pNewWorkspaceEdit.StartEditOperation()

            Dim pNewFeature As IFeature
            pNewFeature = pNewFeatureClass.CreateFeature
            pNewFeature.Shape = pPolygon
            pGeometry = pNewFeature.ShapeCopy
            pGeometry.Project(pNewSR)
            pNewFeature.Shape = pGeometry
            'UPGRADE_WARNING: Couldn't resolve default property of object pNewFeature.Store. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            pNewFeature.Store()

            pNewWorkspaceEdit.StopEditOperation()
            pNewWorkspaceEdit.StopEditing(True)

            ToEnvelope = pNewFeatureClass

        End Function
        Public Function ToCentroid(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass Implements IFeatureTypeConverter.ToCentroid

            Dim I As Short
            'On Error GoTo errHand
            On Error Resume Next
            Dim pNewFeatureClass As IFeatureClass
            Dim pDataset As IDataset
            pDataset = pFeatureClass

            Dim pGeoDataset As IGeoDataset
            pGeoDataset = pFeatureClass

            Dim pFCursor As IFeatureCursor

            If Not pSelSet Is Nothing Then
                pSelSet.Search(Nothing, False, pFCursor)
                iFeatureCount = pSelSet.Count
            Else
                pFCursor = pFeatureClass.Search(Nothing, False)
                iFeatureCount = pFeatureClass.FeatureCount(Nothing)
            End If

            'Check if the line is MAware or ZAware
            Dim lGeomIndex As Integer
            Dim sShpName As String
            Dim pFieldsTest As IFields
            Dim pFieldTest As IField
            Dim pGeometryDefTest As IGeometryDef
            sShpName = pFeatureClass.ShapeFieldName
            pFieldsTest = pFeatureClass.Fields
            lGeomIndex = pFieldsTest.FindField(sShpName)
            pFieldTest = pFieldsTest.Field(lGeomIndex)
            pGeometryDefTest = pFieldTest.GeometryDef
            Dim bZAware As Boolean
            Dim bMAware As Boolean
            'Determine if M or Z aware
            bZAware = pGeometryDefTest.HasZ
            bMAware = pGeometryDefTest.HasM
            'Create a new shapefile

            Dim pFields As IFields
            Dim pFieldsEdit As IFieldsEdit
            pFields = New Fields
            pFieldsEdit = pFields
            pFieldsEdit.FieldCount_2 = 2

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
            'Define the geometry field according to the input attribut awareness
            If (bMAware = False) And (bZAware = False) Then
                With pGeomDefEdit
                    .GeometryType_2 = esriGeometryType.esriGeometryPoint
                    .SpatialReference_2 = pNewSR 'New UnknownCoordinateSystem
                End With
            ElseIf (bMAware = True) And (bZAware = False) Then
                With pGeomDefEdit
                    .GeometryType_2 = esriGeometryType.esriGeometryPoint
                    .SpatialReference_2 = pNewSR 'New UnknownCoordinateSystem
                    .HasM_2 = True
                End With
            ElseIf (bMAware = False) And (bZAware = True) Then
                With pGeomDefEdit
                    .GeometryType_2 = esriGeometryType.esriGeometryPoint
                    .SpatialReference_2 = pNewSR 'New UnknownCoordinateSystem
                    .HasZ_2 = True
                End With
            ElseIf (bMAware = True) And (bZAware = True) Then
                With pGeomDefEdit
                    .GeometryType_2 = esriGeometryType.esriGeometryPoint
                    .SpatialReference_2 = pNewSR 'New UnknownCoordinateSystem
                    .HasZ_2 = True
                    .HasM_2 = True
                End With
            End If
            pFieldEdit.GeometryDef_2 = pGeomDef

            pFieldsEdit.Field_2(1) = pField


            On Error Resume Next
            Err.Clear()
            pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name(), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Do While (Err.Number = fdoError.FDO_E_NO_PERMISSION) Or (Err.Number = fdoError.FDO_E_TABLE_ALREADY_EXISTS)
                Err.Clear()
                I = I + 1
                pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name() & "_" & CStr(I), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            Loop

            'On Error GoTo errHand
            Dim pExistingFields As IFields
            'Variable for storing
            pExistingFields = pNewFeatureClass.Fields
            Dim lFieldCount As Integer
            lFieldCount = pFeatureClass.Fields.FieldCount
            Dim pFieldsIn As IFields
            pFieldsIn = pFeatureClass.Fields

            Dim J As Integer
            Dim bFlag As Boolean

            'Verify if the necessary fields are in the output feature class if not it add those
            For I = 0 To lFieldCount - 1
                bFlag = False
                If pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeOID And pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeGeometry And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_LENGTH" And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_AREA" Then
                    For J = 0 To pExistingFields.FieldCount - 1
                        If pExistingFields.Field(J).Name = pFieldsIn.Field(I).Name Then
                            bFlag = True
                            Exit For
                        End If
                    Next
                    If bFlag = False Then
                        If I = 14 Then
                            System.Diagnostics.Debug.WriteLine("err")
                        End If
                        pNewFeatureClass.AddField(pFieldsIn.Field(I))
                    End If
                End If
            Next

            If pGeomDef.HasZ Then

                If pNewFeatureClass.Fields.FindField("gsZ") = -1 Then
                    pField = New Field
                    pFieldEdit = pField
                    With pFieldEdit
                        .Name_2 = "gsZ"
                        .AliasName_2 = "Z"
                        .Type_2 = esriFieldType.esriFieldTypeDouble
                    End With

                    pNewFeatureClass.AddField(pField)
                End If

                If pNewFeatureClass.Fields.FindField("gsZmin") = -1 Then
                    pField = New Field
                    pFieldEdit = pField
                    With pFieldEdit
                        .Name_2 = "gsZmin"
                        .AliasName_2 = "Zmin"
                        .Type_2 = esriFieldType.esriFieldTypeDouble
                    End With

                    pNewFeatureClass.AddField(pField)
                End If

                If pNewFeatureClass.Fields.FindField("gsZmax") = -1 Then
                    pField = New Field
                    pFieldEdit = pField
                    With pFieldEdit
                        .Name_2 = "gsZmax"
                        .AliasName_2 = "Zmax"
                        .Type_2 = esriFieldType.esriFieldTypeDouble
                    End With

                    pNewFeatureClass.AddField(pField)
                End If

            End If

            Dim pGeometry As IGeometry
            Dim pFeatureBuffer As IFeatureBuffer
            Dim pNewFCursor As IFeatureCursor
            pNewFCursor = pNewFeatureClass.Insert(True)
            pFeatureBuffer = pNewFeatureClass.CreateFeatureBuffer
            Dim pFeature As IFeature
            pFeature = pFCursor.NextFeature
            Dim pPt As IPoint

            StartProgress(iFeatureCount)

            Dim pEnvelope As IEnvelope
            Dim pArea As IArea
            Dim pCenter As IPoint
            Dim GeomCount As Short

            Dim pGeomColl As IGeometryCollection
            Dim pZAware As IZAware
            Dim pMAware As IMAware
            While Not pFeature Is Nothing

                pStepPro.Message = "Converting feature: " & iPosition & " of " & iFeatureCount
                iPosition = iPosition + 1
                pStepPro.Position = iPosition

                'pGeomColl = pFeature.ShapeCopy

                'For GeomCount = 0 To pGeomColl.GeometryCount - 1

                'pEnvelope = pGeomColl.Geometry(GeomCount).Envelope
                'pArea = pEnvelope

                'If Not pGeomColl.Geometry(GeomCount) Is Nothing Then

                pGeometry = pFeature.ShapeCopy
                pArea = pGeometry.Envelope

                pCenter = New Point
                pArea.QueryCentroid(pCenter)

                pPt = New Point
                pPt.PutCoords(pCenter.X, pCenter.Y)

                If pGeomDef.HasZ Then
                    pZAware = pPt
                    pZAware.ZAware = True
                    pPt.Z = (pEnvelope.ZMax + pEnvelope.ZMin) / 2
                End If

                If pGeomDef.HasM Then
                    pMAware = pPt
                    pMAware.MAware = True
                    pPt.M = 0
                End If

                pGeometry = pPt
                pGeometry.Project(pNewSR)
                pFeatureBuffer.Shape = pGeometry

                For I = 0 To lFieldCount - 1
                    If pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeOID And pFieldsIn.Field(I).Type <> esriFieldType.esriFieldTypeGeometry And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_LENGTH" And UCase(pFieldsIn.Field(I).Name) <> "SHAPE_AREA" Then
                        J = pFeatureBuffer.Fields.FindField(pFieldsIn.Field(I).Name)
                        If J <> -1 Then
                            pFeatureBuffer.Value(J) = pFeature.Value(I)
                        End If
                    End If
                Next
                If pGeomDef.HasZ Then
                    pFeatureBuffer.Value(pFeatureBuffer.Fields.FindField("gsZ")) = pPt.Z
                    pFeatureBuffer.Value(pFeatureBuffer.Fields.FindField("gsZmin")) = pEnvelope.ZMin
                    pFeatureBuffer.Value(pFeatureBuffer.Fields.FindField("gsZmax")) = pEnvelope.ZMax
                End If

                pNewFCursor.InsertFeature(pFeatureBuffer)
                'End If

                'Next GeomCount

                pFeature = pFCursor.NextFeature

                bContinue = pTrackCancel.[Continue]
                If Not bContinue Then GoTo exit_Sub
            End While

exit_Sub:

            pPDlg.HideDialog()
            pNewFCursor.Dispose()

            Return pNewFeatureClass

            Exit Function
errHand:
            MsgBox(Err.Description)
        End Function
        Public Function Stratify(ByVal pFeatureLayer As IFeatureLayer, ByVal pSelSet As ISelectionSet) As FeatureLayerArray Implements IFeatureTypeConverter.Stratify

            Dim pGeoLayer As IGeoFeatureLayer

            pGeoLayer = pFeatureLayer

            If (TypeOf pGeoLayer.Renderer Is IClassBreaksRenderer) Then
                Stratify = StratifyGradColors(pFeatureLayer, pSelSet)
                Exit Function
            End If


            Dim pUVRend As IUniqueValueRenderer
            pUVRend = pGeoLayer.Renderer

            Dim sFieldName As String
            Dim pAttribField As IField
            Dim AttribStringType As Boolean
            Dim I, J As Short
            Dim varValue As Object
            Dim varLabel As Object
            Dim AllValues As String

            '----------------------------------------
            Dim QueryValue As String
            Dim QuerySymbol As Char = "'"
            Dim pQF() As IQueryFilter
            ReDim pQF(pUVRend.ValueCount - 1)
            sFieldName = pUVRend.Field(0)
            pAttribField = pFeatureLayer.FeatureClass.Fields.Field(pFeatureLayer.FeatureClass.Fields.FindField(sFieldName))
            AttribStringType = (pAttribField.Type = esriFieldType.esriFieldTypeString)

            If pUVRend.Value(0) <> "<Null>" Then
                If Not AttribStringType Then
                    AllValues = CStr(pUVRend.Value(0))
                Else
                    QueryValue = pUVRend.Value(0)
                    GetQuerySymbol(QueryValue)
                    AllValues = QuerySymbol & QueryValue & QuerySymbol
                End If
            Else
                AllValues = "NULL"
            End If
            For I = 0 To pUVRend.ValueCount - 1
                pQF(I) = New QueryFilter
                Try
                    For J = 0 To pUVRend.ValueCount - 1
                        If pUVRend.Value(J) = pUVRend.ReferenceValue(pUVRend.Value(I)) Then
                            If Not AttribStringType Then
                                pQF(J).WhereClause = pQF(J).WhereClause & " OR " & sFieldName & "=" & pUVRend.Value(I)
                            Else
                                QueryValue = pUVRend.Value(I)
                                GetQuerySymbol(QueryValue)
                                pQF(J).WhereClause = pQF(J).WhereClause & " OR " & sFieldName & "=" & QuerySymbol & QueryValue & QuerySymbol
                            End If
                            pQF(I).WhereClause = "Skip"
                            Exit For
                        End If
                    Next

                Catch ex As Exception

                    If pUVRend.Value(I) <> "<Null>" Then
                        If Not AttribStringType Then
                            pQF(I).WhereClause = sFieldName & "=" & pUVRend.Value(I)
                        Else
                            QueryValue = pUVRend.Value(I)
                            GetQuerySymbol(QueryValue)
                            pQF(I).WhereClause = sFieldName & "=" & QuerySymbol & QueryValue & QuerySymbol
                        End If
                    Else
                        pQF(I).WhereClause = sFieldName & " IS NULL"
                    End If
                End Try

                If I > 0 Then
                    If pUVRend.Value(I) <> "<Null>" Then
                        If Not AttribStringType Then
                            AllValues = AllValues & "," & CStr(pUVRend.Value(I))
                        Else
                            QueryValue = pUVRend.Value(I)
                            GetQuerySymbol(QueryValue)
                            AllValues = AllValues & "," & QuerySymbol & QueryValue & QuerySymbol
                        End If
                    End If
                Else
                    AllValues = AllValues & ",NULL"
                End If
            Next I

            If pUVRend.UseDefaultSymbol Then
                ReDim Preserve pQF(pUVRend.ValueCount)
                pQF(pUVRend.ValueCount) = New QueryFilter
                pQF(pUVRend.ValueCount).WhereClause = " NOT " & sFieldName & " IN (" & AllValues & ")"
            End If
            '----------------------------------------

            Dim pNewFeatureLayer As IFeatureLayer
            Dim pNewFeatureClass As IFeatureClass
            Dim pAttribSel As ISelectionSet
            Dim pOcSchemaEdit As IClassSchemaEdit
            Dim pOc As IObjectClass
            Dim pSimpleRend As ISimpleRenderer
            Dim pGeoFL As IGeoFeatureLayer
            Dim pAnnotateLayerProperties As IAnnotateLayerProperties
            Dim pDS As IDataset = pFeatureLayer.FeatureClass
            'pScratchWorkspaceFactory = New ESRI.ArcGIS.DataSourcesGDB.ScratchWorkspaceFactory
            'pScratchWorkspace = pScratchWorkspaceFactory.DefaultScratchWorkspace
            For I = 0 To pQF.Length - 1
                If pQF(I).WhereClause <> "Skip" Then
                    varValue = pUVRend.Value(I)
                    varLabel = pUVRend.Label(varValue)
                    If pSelSet Is Nothing Then
                        pAttribSel = pFeatureLayer.FeatureClass.Select(pQF(I), esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pDS.Workspace)
                    Else
                        pAttribSel = pSelSet.Select(pQF(I), esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pDS.Workspace)
                    End If

                    If I = pUVRend.ValueCount And pAttribSel.Count = 0 Then Return Stratify
                    pNewFeatureClass = ExportFC(pDSName.Name() & "_" & sFieldName & "_" & NameCheck(CStr(varValue)), pFeatureLayer.FeatureClass, pAttribSel)
                    Try
                        pOc = pNewFeatureClass
                        pOcSchemaEdit = pOc
                        pOcSchemaEdit.AlterAliasName(varLabel)
                    Catch ex As Exception
                    End Try
                    pNewFeatureLayer = New FeatureLayer
                    pNewFeatureLayer.FeatureClass = pNewFeatureClass
                    pNewFeatureLayer.Name = pNewFeatureClass.AliasName
                    pSimpleRend = New SimpleRenderer
                    If I < pUVRend.ValueCount Then
                        pSimpleRend.Symbol = pUVRend.Symbol(varValue)
                    Else
                        pSimpleRend.Symbol = pUVRend.DefaultSymbol
                    End If
                    pGeoFL = pNewFeatureLayer
                    pGeoFL.Renderer = pSimpleRend

                    pGeoFL.AnnotationProperties.Clear()
                    For J = 0 To pGeoLayer.AnnotationProperties.Count - 1
                        pGeoLayer.AnnotationProperties.QueryItem(J, pAnnotateLayerProperties)
                        pGeoFL.AnnotationProperties.Add(pAnnotateLayerProperties)
                    Next J

                    ReDim Preserve Stratify.FeatureLayer(I)
                    Stratify.FeatureLayer(I) = pNewFeatureLayer
                End If
            Next I

            Return Stratify
        End Function
        Private Function GetQuerySymbol(ByRef Value As String) As Boolean
            Value = Value.Replace("'", "''")
        End Function
        Private Function StratifyGradColors(ByVal pFeatureLayer As IFeatureLayer, ByVal pSelSet As ISelectionSet) As FeatureLayerArray
            On Error GoTo err_sub

            Dim pGeoLayer As IGeoFeatureLayer
            Dim pFeatureRender As IFeatureRenderer
            Dim pCBRender As IClassBreaksRenderer
            Dim pUIProperties As IClassBreaksUIProperties

            pGeoLayer = pFeatureLayer
            pFeatureRender = pGeoLayer.Renderer

            Dim sFieldName As String

            pCBRender = pFeatureRender
            pUIProperties = pCBRender

            sFieldName = pCBRender.Field

            Dim I As Short
            Dim varLowValue As Double
            Dim varHighValue As Double
            Dim varLabel As Object

            Dim pNewFeatureLayer As IFeatureLayer
            Dim pNewFeatureClass As IFeatureClass

            Dim pQF As IQueryFilter
            Dim pAttribSel As ISelectionSet
            Dim pScratchWorkspace As IWorkspace
            Dim pScratchWorkspaceFactory As IScratchWorkspaceFactory
            Dim pOcSchemaEdit As IClassSchemaEdit
            Dim pOc As IObjectClass


            Dim pSimpleRend As ISimpleRenderer
            Dim pLegendInfo As ILegendInfo
            Dim pGeoFL As IGeoFeatureLayer
            Dim pAnnotateLayerProperties As IAnnotateLayerProperties
            Dim J As Short
            For I = 0 To pCBRender.BreakCount - 1

                varLowValue = pUIProperties.LowBreak(I)
                varHighValue = pCBRender.Break(I)
                varLabel = pCBRender.Label(I)

                pScratchWorkspaceFactory = New ESRI.ArcGIS.DataSourcesGDB.ScratchWorkspaceFactory
                pScratchWorkspace = pScratchWorkspaceFactory.DefaultScratchWorkspace
                pQF = New QueryFilter
                pQF.WhereClause = sFieldName & " >= " & varLowValue & " AND " & sFieldName & " <= " & varHighValue

                If pSelSet Is Nothing Then
                    pAttribSel = pFeatureLayer.FeatureClass.Select(pQF, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pScratchWorkspace)
                Else
                    pAttribSel = pSelSet.Select(pQF, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pScratchWorkspace)
                End If

                pNewFeatureClass = ExportFC(pDSName.Name() & "_" & sFieldName & "_" & NameCheck(CStr(varLowValue) & "_to_" & CStr(varHighValue)), pFeatureLayer.FeatureClass, pAttribSel)


                On Error Resume Next
                pOc = pNewFeatureClass
                pOcSchemaEdit = pOc

                pOcSchemaEdit.AlterAliasName(varLowValue & " - " & varHighValue)


                pNewFeatureLayer = New FeatureLayer
                pNewFeatureLayer.FeatureClass = pNewFeatureClass
                pNewFeatureLayer.Name = pNewFeatureClass.AliasName

                ' set symbol.  we must use ISimpleRenderer interface
                pSimpleRend = New SimpleRenderer
                pLegendInfo = pGeoLayer.Renderer

                pSimpleRend.Symbol = pLegendInfo.LegendGroup(0).Class(I).Symbol

                ' finally, set the new renderer to the layer and refresh the map
                pGeoFL = pNewFeatureLayer
                pGeoFL.Renderer = pSimpleRend

                pGeoFL.AnnotationProperties.Clear()
                For J = 0 To pGeoLayer.AnnotationProperties.Count - 1
                    pGeoLayer.AnnotationProperties.QueryItem(J, pAnnotateLayerProperties)
                    pGeoFL.AnnotationProperties.Add(pAnnotateLayerProperties)
                    'pAnnotateLayerProperties.DisplayAnnotation = True
                Next J
                If pGeoLayer.DisplayAnnotation Then
                    pGeoFL.DisplayAnnotation = False
                    pGeoFL.DisplayAnnotation = True
                End If
                ReDim Preserve StratifyGradColors.FeatureLayer(I)
                StratifyGradColors.FeatureLayer(I) = pNewFeatureLayer
            Next I


            Return StratifyGradColors

err_sub:
            MsgBox(Err.Description)

        End Function

        Public Function RemoveDuplicates(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As FeatureClassArray Implements IFeatureTypeConverter.RemoveDuplicates
           
            Dim pFeatureCursor As IFeatureCursor
            Dim pFeature As IFeature
            Dim pDelSelSet As ISelectionSet
            Dim pAllDelSelSet As ISelectionSet
            Dim pCleanSelSet As ISelectionSet

            Dim pScratchWorkspace As IWorkspace
            Dim pScratchWorkspaceFactory As IScratchWorkspaceFactory
            pScratchWorkspaceFactory = New ESRI.ArcGIS.DataSourcesGDB.ScratchWorkspaceFactory
            pScratchWorkspace = pScratchWorkspaceFactory.DefaultScratchWorkspace

            Dim pQFilt As IQueryFilter
            pQFilt = New QueryFilter
            pAllDelSelSet = pFeatureClass.Select(Nothing, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionEmpty, pScratchWorkspace)
            pCleanSelSet = pFeatureClass.Select(Nothing, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionEmpty, pScratchWorkspace)

            If pSelSet Is Nothing Then
                pFeatureCursor = pFeatureClass.Search(Nothing, False)
                iFeatureCount = pFeatureClass.FeatureCount(Nothing)
            Else
                pSelSet.Search(Nothing, False, pFeatureCursor)
                iFeatureCount = pSelSet.Count
            End If

            pFeature = pFeatureCursor.NextFeature

            StartProgress(iFeatureCount)
            Dim pSpatFilter As ISpatialFilter = New SpatialFilter
            pSpatFilter.GeometryField = pFeatureClass.ShapeFieldName
            Do While Not pFeature Is Nothing
                pQFilt.WhereClause = pFeatureClass.OIDFieldName & "=" & pFeature.OID
                If pAllDelSelSet.Select(pQFilt, esriSelectionType.esriSelectionTypeSnapshot, esriSelectionOption.esriSelectionOptionOnlyOne, pScratchWorkspace).Count = 0 Then
                    pSpatFilter.Geometry = pFeature.ShapeCopy
                    pSpatFilter.WhereClause = pFeatureClass.OIDFieldName & "<>" & pFeature.OID
                    pSpatFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelContains
                    If pSelSet Is Nothing Then
                        pDelSelSet = pFeatureClass.Select(pSpatFilter, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pScratchWorkspace)
                    Else
                        pDelSelSet = pSelSet.Select(pSpatFilter, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pScratchWorkspace)
                    End If
                    pSpatFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelWithin
                    pDelSelSet = pDelSelSet.Select(pSpatFilter, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pScratchWorkspace)
                    If pDelSelSet.Count > 0 Then
                        pAllDelSelSet.Combine(pDelSelSet, esriSetOperation.esriSetUnion, pAllDelSelSet)
                    End If
                    pCleanSelSet.Add(pFeature.OID)
                End If
                pFeature = pFeatureCursor.NextFeature
                pStepPro.Message = "Processing feature: " & iPosition & " of " & iFeatureCount
                iPosition = iPosition + 1
                pStepPro.Position = iPosition
                bContinue = pTrackCancel.[Continue]
                If Not bContinue Then GoTo exit_Sub
            Loop

exit_Sub:
            pPDlg.HideDialog()
            Dim pNewFeatureClass As IFeatureClass
            Dim pRemFeatureClass As IFeatureClass

            If pAllDelSelSet.Count = 0 Then
                MsgBox("Duplicates are not found")
                ReDim RemoveDuplicates.FeatureClass(0)
                RemoveDuplicates.FeatureClass(0) = Nothing
            Else
                pNewFeatureClass = ExportFC(pDSName.Name(), pFeatureClass, pCleanSelSet)
                pRemFeatureClass = ExportFC(pDSName.Name() & "_duplicate", pFeatureClass, pAllDelSelSet)
                MsgBox(pAllDelSelSet.Count & " duplicates are removed.")
                ReDim RemoveDuplicates.FeatureClass(1)
                RemoveDuplicates.FeatureClass(0) = pNewFeatureClass
                RemoveDuplicates.FeatureClass(1) = pRemFeatureClass
            End If

        End Function

        Public Function DivideSegments(ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet, ByVal pMap As IMap, ByVal Lp As Double) As Boolean Implements IFeatureTypeConverter.DivideSegments
            Dim i As Integer
            'Dim pClone As IClone = pFeatureClass.Fields
            'Dim pFields = pClone.Clone
            'Dim pNewFeatureClass As IFeatureClass
            'pNewFeatureClass = pOutWorkspace.CreateFeatureClass(pDSName.Name(), pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
            'Try
            'CType(pNewFeatureClass, IClass).DeleteField(pNewFeatureClass.Fields.Field(pNewFeatureClass.FindField("Id")))
            'Catch ex As Exception
            'End Try

            Dim pFeature As IFeature
            Dim pFCursor As IFeatureCursor
            If pSelSet Is Nothing Then
                pFCursor = pFeatureClass.Search(Nothing, False)
            Else
                pSelSet.Search(Nothing, False, pFCursor)
            End If

            pFeature = pFCursor.NextFeature
            'Dim pPointColl As IPointCollection
            Dim pSegmentColl As ISegmentCollection
            Dim pNewPointColl As IPointCollection
            Dim pGeomColl As IGeometryCollection
            Dim pNewGeomColl As IGeometryCollection
            Select Case pFeatureClass.ShapeType
                Case esriGeometryType.esriGeometryPolygon
                    pNewGeomColl = New Polygon
                Case esriGeometryType.esriGeometryPolyline
                    pNewGeomColl = New Polyline
            End Select

            Dim pWSE As IWorkspaceEdit = CType(pFeatureClass, IDataset).Workspace
            'Dim pGeometry As IGeometry

            pWSE.StartEditing(False)
            pWSE.StartEditOperation()
            While Not pFeature Is Nothing
                pGeomColl = pFeature.ShapeCopy
                Dim GeomIndex As Integer
                For GeomIndex = 0 To pGeomColl.GeometryCount - 1
                    Select Case pFeatureClass.ShapeType
                        Case esriGeometryType.esriGeometryPolygon
                            pNewPointColl = New Ring
                        Case esriGeometryType.esriGeometryPolyline
                            pNewPointColl = New ESRI.ArcGIS.Geometry.Path
                    End Select
                    'pGeometry = pNewPointColl
                    'pGeometry.SpatialReference = pNewSR
                    'pNewPointColl = pGeometry
                    pSegmentColl = pGeomColl.Geometry(GeomIndex)
                    For i = 0 To pSegmentColl.SegmentCount - 1
                        Dim Dist = pMap.ComputeDistance(pSegmentColl.Segment(i).FromPoint, pSegmentColl.Segment(i).ToPoint)
                        Dim n = 1
                        pNewPointColl.AddPoint(pSegmentColl.Segment(i).FromPoint)
                        While Lp * n < Dist
                            Dim outPoint As New Point
                            pSegmentColl.Segment(i).QueryPoint(esriSegmentExtension.esriNoExtension, Lp * n / Dist, True, outPoint)
                            pNewPointColl.AddPoint(outPoint)
                            n = n + 1
                        End While
                        pNewPointColl.AddPoint(pSegmentColl.Segment(i).ToPoint)
                    Next
                    pNewGeomColl.AddGeometry(pNewPointColl)

                    'Dim x, y As Double
                    'With pFeature.ShapeCopy.Envelope
                    '    x = .XMin + (.XMax - .XMin) / 50
                    '    While x < .XMax
                    '        y = .YMin + (.YMax - .YMin) / 50
                    '        While y < .YMax
                    '            y = y + (.YMax - .YMin) / 50
                    '            Dim pPoint As New Point
                    '            pNewPointColl = New TriangleStrip
                    '            pPoint.PutCoords(x, y)
                    '            pNewPointColl.AddPoint(pPoint)
                    '            pPoint.PutCoords(x - 0.001, y + 0.001)
                    '            pNewPointColl.AddPoint(pPoint)
                    '            pPoint.PutCoords(x + 0.001, y - 0.001)
                    '            pNewPointColl.AddPoint(pPoint)
                    '            pPoint.PutCoords(x, y)
                    '            pNewPointColl.AddPoint(pPoint)
                    '            pNewGeomColl.AddGeometry(pNewPointColl)
                    '        End While
                    '        x = x + (.XMax - .XMin) / 50
                    '    End While
                    'End With
                Next

                'pNewFeature = pNewFeatureClass.CreateFeature
                pFeature.Shape = pNewGeomColl


                'For i = 0 To pFeatureClass.Fields.FieldCount - 1
                '    If pFeatureClass.Fields.Field(i).Type <> esriFieldType.esriFieldTypeOID And pFeatureClass.Fields.Field(i).Type <> esriFieldType.esriFieldTypeGeometry And pNewFeatureClass.Fields.Field(i).Editable Then
                '        pNewFeature.Value(i) = pFeature.Value(i)
                '    End If
                'Next
                pFeature.Store()
                pFeature = pFCursor.NextFeature
            End While

            pWSE.StartEditOperation()
            pWSE.StopEditing(True)

            Return True

        End Function
        Private Function GeomIsEquals(ByVal Geom1 As IGeometry, ByVal Geom2 As IGeometry) As Object
            Dim pRelOp As IRelationalOperator

            pRelOp = Geom1

            GeomIsEquals = pRelOp.Equals(Geom2)

        End Function

        Private Function ExportFC(ByVal FCName As String, ByVal pFeatureClass As IFeatureClass, ByVal pSelSet As ISelectionSet) As IFeatureClass
            Dim pWorkspace As IWorkspace
            Dim pInputDataset As IDataset
            Dim pInputDatasetName As IDatasetName
            pInputDataset = pFeatureClass
            pInputDatasetName = pInputDataset.FullName
            Dim pInWS As IDataset
            Dim pInWSName As IWorkspaceName
            pWorkspace = pOutWorkspace
            pInWS = pWorkspace
            pInWSName = pInWS.FullName

            Dim OutFCName As IFeatureClassName
            Dim OutDSName As IDatasetName
            OutFCName = New FeatureClassName
            OutDSName = OutFCName
            OutDSName.Name = FCName

            Dim pOutWSName As IWorkspaceName
            pOutWSName = New WorkspaceName
            pOutWSName.WorkspaceFactoryProgID = pInWSName.WorkspaceFactoryProgID
            pOutWSName.PathName = pWorkspace.PathName
            OutDSName.WorkspaceName = pOutWSName

            Dim pGeomDef As IGeometryDef
            Dim pGeomDefEdit As IGeometryDefEdit
            Dim lGeomIndex As Integer
            Dim sShpName As String
            Dim pFields As IFields
            Dim pField As IField
            Dim pGeometryDef As IGeometryDef
            sShpName = pFeatureClass.ShapeFieldName
            'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureClass.Fields. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            pFields = pFeatureClass.Fields
            lGeomIndex = pFields.FindField(sShpName)
            pField = pFields.Field(lGeomIndex)
            pGeometryDef = pField.GeometryDef
            pGeomDef = pField.GeometryDef
            pGeomDefEdit = pGeomDef
            pGeomDefEdit.SpatialReference_2 = pNewSR
            pGeomDefEdit.HasM_2 = pFields.Field(lGeomIndex).GeometryDef.HasM
            pGeomDefEdit.HasZ_2 = pFields.Field(lGeomIndex).GeometryDef.HasZ
            Dim pExOp As IExportOperation
            pExOp = New ExportOperation
            pExOp.ExportFeatureClass(pInputDatasetName, Nothing, pSelSet, pGeomDef, OutFCName, 0)
            Dim pName As IName
            pName = OutFCName
            ExportFC = pName.Open
        End Function

        Private Function NameCheck(ByVal pName As String) As String
            Dim I As Byte
            Dim RestrictSymbols(15) As Object
            RestrictSymbols(1) = " "
            RestrictSymbols(2) = "."
            RestrictSymbols(3) = ","
            RestrictSymbols(4) = ":"
            RestrictSymbols(5) = "/"
            RestrictSymbols(6) = "\"
            RestrictSymbols(7) = "-"
            RestrictSymbols(8) = "("
            RestrictSymbols(9) = ")"
            RestrictSymbols(10) = "'"
            RestrictSymbols(11) = "#"
            RestrictSymbols(12) = "!"
            RestrictSymbols(13) = "?"
            RestrictSymbols(14) = "*"
            RestrictSymbols(15) = "@"
            NameCheck = pName

            Dim LayerNameArray() As String
            For I = 1 To UBound(RestrictSymbols)
                If InStr(1, pName, RestrictSymbols(I)) <> 0 Then
                    LayerNameArray = Split(NameCheck, RestrictSymbols(I), , CompareMethod.Text)
                    NameCheck = Join(LayerNameArray, "_")
                End If
            Next I

        End Function

        Private Function UniqueValuesArray(ByVal pCursor As ICursor, ByVal sFieldName As String) As Object

            Dim pData As IDataStatistics

            pData = New DataStatistics
            pData.Field = sFieldName
            pData.Cursor = pCursor

            Dim value As Object
            'Dim UniqueValues() As Integer
            Dim UniqueValues() As Object
            ReDim UniqueValues(0)
            Dim pEnumVar As IEnumerator
            pEnumVar = pData.UniqueValues
            pEnumVar.MoveNext()
            value = pEnumVar.Current
            UniqueValues(0) = value

            Do Until IsNothing(value)

                pEnumVar.MoveNext()
                value = pEnumVar.Current

                If Not IsNothing(value) Then
                    ReDim Preserve UniqueValues(UniqueValues.Length)
                    UniqueValues(UniqueValues.Length - 1) = value
                End If
            Loop
            If pData.UniqueValueCount <> 0 Then
                ReDim Preserve UniqueValues(UniqueValues.Length)
                UniqueValues(UniqueValues.Length - 1) = Nothing
            End If
            Return UniqueValues
        End Function
        Private Function fnMakeFields(ByVal pOldFields As IFields, ByVal strShapeFieldName As String, ByVal iGeometryType As esriGeometryType) As IFields
            '
            ' This function creates a new Fields collection from an existing object,
            ' but changes the Geometry type held to Polylines. We pass in the existing
            ' Fields collection by value to avoid changing the existing Fields.
            '
            Dim pFieldsEdit As IFieldsEdit
            Dim pField As IField
            Dim pFieldEdit As IFieldEdit
            fnMakeFields = New Fields
            pFieldsEdit = fnMakeFields
            pFieldsEdit.FieldCount_2 = pOldFields.FieldCount

            Dim I As Short
            Dim pGeomDef As IGeometryDef
            Dim pGeomDefEdit As IGeometryDefEdit
            For I = 0 To pOldFields.FieldCount - 1
                If pOldFields.Field(I).Type = esriFieldType.esriFieldTypeGeometry Then
                    pField = New Field
                    pFieldEdit = pField
                    pFieldEdit.Name_2 = "Shape"
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeGeometry
                    pGeomDef = New GeometryDef
                    pGeomDefEdit = pGeomDef
                    pGeomDefEdit.GeometryType_2 = iGeometryType
                    pGeomDefEdit.HasM_2 = pOldFields.Field(I).GeometryDef.HasM
                    pGeomDefEdit.HasZ_2 = pOldFields.Field(I).GeometryDef.HasZ
                    pGeomDefEdit.SpatialReference_2 = pNewSR
                    pFieldEdit.GeometryDef_2 = pGeomDef
                    pFieldsEdit.Field_2(I) = pField
                Else
                    pFieldsEdit.Field_2(I) = pOldFields.Field(I)
                End If
            Next I

        End Function

        Private Function fnHasClosedLines(ByVal pFeatClass As IFeatureClass) As Boolean
            fnHasClosedLines = False
            If Not pFeatClass.ShapeType = esriGeometryType.esriGeometryPolyline Then Exit Function
            '
            ' Check each feature in the FeatureClass to see if any features can be classed as
            ' closed. We could use the ICurve.IsClosed property, but this way we can allow
            ' for a small tolerance in the distance between the FromPoint and ToPoint.
            '
            Dim pFeatureCursor As IFeatureCursor
            Dim pFeature As IFeature
            pFeatureCursor = pFeatClass.Search(Nothing, False)
            pFeature = pFeatureCursor.NextFeature
            '
            ' Iterate all Features.
            '
            Do While Not pFeature Is Nothing
                'UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Shape. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                If fnCurveClosed((pFeature.Shape)) Then
                    fnHasClosedLines = True
                    Exit Function
                End If
                pFeature = pFeatureCursor.NextFeature
            Loop
        End Function

        Private Function fnCurveClosed(ByVal pCurve As ICurve) As Boolean
            fnCurveClosed = False
            '
            ' We use the IProximityOperator interface to work out the distance between
            ' the FromPoint and ToPoint of the curve passed in.
            '
            Dim pProxOp As IProximityOperator
            Dim dblDist As Double
            pProxOp = pCurve.FromPoint
            'UPGRADE_WARNING: Couldn't resolve default property of object pCurve.ToPoint. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            dblDist = pProxOp.ReturnDistance(pCurve.ToPoint)
            '
            ' This is our tolerance for a 'closed' curve. We could relate this to a snapping
            ' tolerance, or to the spatial reference, or to the size of the shape, instead
            ' of hardcoding the distance.
            '
            If (dblDist < 0.0001) Then fnCurveClosed = True
        End Function

        Private Function fnPolylineToPolygon(ByVal pGeometry As IGeometry) As IPolygon
            '
            ' This function converts a Polyline to a Polygon, by creating a new Polygon
            ' object and copying the Segments from the Polyline to the new Polygon, and
            ' then ensuring the Polygon is Simple.
            '
            fnPolylineToPolygon = New Polygon
            '
            ' Passing the Polyline by value means that we do not need to clone the Polyline.
            '
            Dim pPolygonSegs, pPolylineSegs As ISegmentCollection
            pPolygonSegs = fnPolylineToPolygon
            pPolylineSegs = pGeometry
            '
            ' Here we copy the Segment objects by using the QuerySegments and AddSegments
            ' methods on the ISegmentCollection interface, which is implemented by both
            ' a Polygon and Polyline coclass.
            '
            Dim pSegs() As ISegment
            ReDim pSegs(pPolylineSegs.SegmentCount - 1)
            pPolylineSegs.QuerySegments(0, pPolylineSegs.SegmentCount, pSegs(0))
            If UBound(pSegs) > 0 Then
                pPolygonSegs.AddSegments(UBound(pSegs) + 1, pSegs(0))
            End If
            '
            ' The Polygon may have it's rings oriented incorrectly, or have overlapping Rings.
            ' We call simplify here to ensure the new Polygon is Simple, which is a requirement
            ' for adding to a FeatureClass.
            '
            fnPolylineToPolygon.SimplifyPreserveFromTo()
        End Function

        Private Function fnPolygonToPolyline(ByVal pGeometry As IGeometry) As IPolyline
            '
            ' This function converts a Polyline to a Polygon, by creating a new Polygon
            ' object and copying the Segments from the Polyline to the new Polygon, and
            ' then ensuring the Polygon is Simple.
            '
            fnPolygonToPolyline = New Polyline
            '
            '
            Dim pPolylineSegs, pPolygonSegs As ISegmentCollection
            pPolylineSegs = fnPolygonToPolyline
            pPolygonSegs = pGeometry
            '
            ' Here we copy the Segment objects by using the QuerySegments and AddSegments
            ' methods on the ISegmentCollection interface, which is implemented by both
            ' a Polygon and Polyline coclass.
            '
            Dim pSegs() As ISegment
            ReDim pSegs(pPolygonSegs.SegmentCount - 1)
            pPolygonSegs.QuerySegments(0, pPolygonSegs.SegmentCount, pSegs(0))
            If UBound(pSegs) > 0 Then
                pPolylineSegs.AddSegments(UBound(pSegs) + 1, pSegs(0))
            End If
            '
            ' The Polygon may have it's rings oriented incorrectly, or have overlapping Rings.
            ' We call simplify here to ensure the new Polygon is Simple, which is a requirement
            ' for adding to a FeatureClass.

        End Function
        Public Function ConvertGraphics(ByVal pGraphicsContainer As IGraphicsContainer, ByVal pGraphicElementEnum As IEnumElement) As FeatureClassArray Implements IFeatureTypeConverter.ConvertGraphics

            Dim I As Object
            Dim pGeometry As IGeometry
            Dim pNewPointFC As IFeatureClass
            Dim pNewPolylineFC As IFeatureClass
            Dim pNewPolygonFC As IFeatureClass
            Dim pPointCursor As IFeatureCursor
            Dim pPolylineCursor As IFeatureCursor
            Dim pPolygonCursor As IFeatureCursor
            Dim pPointBuffer As IFeatureBuffer
            Dim pPolylineBuffer As IFeatureBuffer
            Dim pPolygonBuffer As IFeatureBuffer
            Dim sText As Object
            Dim pTextElement As ITextElement
            Dim pGroupElement As IGroupElement
            Dim pEnumGroupElements As IEnumElement
            Dim pGElement As IElement
            Dim PolylineFlag, PointFlag, PolygonFlag As Boolean
            Dim pElement As IElement
            Dim iElementCount As Object

            If pGraphicElementEnum Is Nothing Then
                pGraphicsContainer.Reset()
                pElement = pGraphicsContainer.Next
            Else
                pGraphicElementEnum.Reset()
                pElement = pGraphicElementEnum.Next
            End If



            PointFlag = False
            PolylineFlag = False
            PolygonFlag = False

            Do While Not pElement Is Nothing

                sText = ""
                pGElement = pElement
                iElementCount = 1

                If TypeOf pElement Is IGroupElement Then
                    pGroupElement = pElement
                    pEnumGroupElements = pGroupElement.Elements
                    pEnumGroupElements.Reset()
                    pGElement = pEnumGroupElements.Next
                    iElementCount = pGroupElement.ElementCount
                End If

                For I = 1 To iElementCount

                    pGeometry = pGElement.Geometry

                    If TypeOf pGElement Is ITextElement Then
                        pTextElement = pGElement
                        sText = pTextElement.Text
                    End If

                    pGeometry.Project(pNewSR)

                    Select Case pGeometry.GeometryType

                        Case esriGeometryType.esriGeometryPoint

                            If Not PointFlag Then
                                pNewPointFC = CreateFC(pOutWorkspace, pDSName.Name() & "_point", esriGeometryType.esriGeometryPoint, pNewSR)
                                pPointCursor = pNewPointFC.Insert(True)
                                pPointBuffer = pNewPointFC.CreateFeatureBuffer
                                PointFlag = True
                            End If
                            pPointBuffer.Shape = pGeometry
                            pPointBuffer.Value(2) = sText
                            pPointCursor.InsertFeature(pPointBuffer)

                        Case esriGeometryType.esriGeometryPolyline
                            If Not PolylineFlag Then
                                pNewPolylineFC = CreateFC(pOutWorkspace, pDSName.Name() & "_polyline", esriGeometryType.esriGeometryPolyline, pNewSR)
                                pPolylineCursor = pNewPolylineFC.Insert(True)
                                pPolylineBuffer = pNewPolylineFC.CreateFeatureBuffer
                                PolylineFlag = True
                            End If
                            pPolylineBuffer.Shape = pGeometry
                            pPolylineBuffer.Value(2) = sText
                            pPolylineCursor.InsertFeature(pPolylineBuffer)


                        Case esriGeometryType.esriGeometryPolygon
                            If Not PolygonFlag Then

                                pNewPolygonFC = CreateFC(pOutWorkspace, pDSName.Name() & "_polygon", esriGeometryType.esriGeometryPolygon, pNewSR)
                                pPolygonCursor = pNewPolygonFC.Insert(True)
                                pPolygonBuffer = pNewPolygonFC.CreateFeatureBuffer
                                PolygonFlag = True
                            End If
                            pPolygonBuffer.Shape = pGeometry
                            pPolygonBuffer.Value(2) = sText
                            pPolygonCursor.InsertFeature(pPolygonBuffer)


                    End Select

                    If Not (pEnumGroupElements Is Nothing) Then pGElement = pEnumGroupElements.Next

                Next I

                If pGraphicElementEnum Is Nothing Then
                    pElement = pGraphicsContainer.Next
                Else
                    pElement = pGraphicElementEnum.Next
                End If

            Loop

            ReDim ConvertGraphics.FeatureClass(2)

            ConvertGraphics.FeatureClass(0) = pNewPointFC
            ConvertGraphics.FeatureClass(1) = pNewPolylineFC
            ConvertGraphics.FeatureClass(2) = pNewPolygonFC

            Exit Function

err_sub:
            MsgBox(Err.Description)
        End Function
        Private Function CreateFC(ByVal OutWS As IFeatureWorkspace, ByVal FCName As String, ByVal GeomType As esriGeometryType, ByVal pNewSR As ISpatialReference) As IFeatureClass
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
                'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                I = I + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                sName = FCName & "_" & CStr(I)
                pFC = OutWS.OpenFeatureClass(sName)
            Loop

            pFC = OutWS.CreateFeatureClass(sName, pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, "Shape", "")

            CreateFC = pFC

        End Function
        Public Sub ExportToBLN(ByVal pFeatureClass As IFeatureClass, ByVal FileName As String, ByVal Outside As Boolean) Implements IFeatureTypeConverter.ExportToBLN

            Dim pFeatureCursor As IFeatureCursor
            Dim pFeature As IFeature
            Dim pPtColl As IPointCollection
            Dim pEnumVertices As IEnumVertex
            Dim pPt As IPoint
            Dim lPartIndex, lVertexIndex As Integer
            Dim iPointCount As Short
            Dim Y, X, Z As Double

            pPt = New ESRI.ArcGIS.Geometry.Point

            FileOpen(1, FileName, OpenMode.Output) ' Open file.

            pFeatureCursor = pFeatureClass.Search(Nothing, False)
            pFeature = pFeatureCursor.NextFeature

            Do While Not pFeature Is Nothing
                If pFeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Or pFeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then

                    pPtColl = pFeature.ShapeCopy
                    pEnumVertices = pPtColl.EnumVertices
                    iPointCount = pPtColl.PointCount

                    PrintLine(1, CStr(iPointCount), ",", CStr(IIf(Outside, 0, 1)))
                    pEnumVertices.QueryNext(pPt, lPartIndex, lVertexIndex)
                    Do While Not pPt.IsEmpty
                        X = pPt.X
                        Y = pPt.Y
                        If pFeatureClass.Fields.Field(pFeatureClass.FindField("Shape")).GeometryDef.HasZ Then
                            Z = pPt.Z
                            PrintLine(1, CStr(X), ",", CStr(Y), ",", CStr(Z))
                        Else
                            PrintLine(1, CStr(X), ",", CStr(Y))
                        End If
                        pEnumVertices.QueryNext(pPt, lPartIndex, lVertexIndex)
                    Loop

                Else

                    'Set pPt = pFeature.ShapeCopy
                    'X = pPt.X
                    'Y = pPt.Y
                End If

                pFeature = pFeatureCursor.NextFeature

            Loop
            FileClose(1)
        End Sub
        Public Sub ConvertToKML(ByVal pFeatureLayer As IFeatureLayer, ByVal FileName As String, ByVal LabelFieldName As String, ByVal PMFlag As Boolean, Optional ByVal blnExtrude As Integer = 0, Optional ByVal AltitudeMode As Integer = 0, Optional ByVal AttributeFieldName As String = "", Optional ByVal SetValue As Double = 0, Optional ByVal Distribute As Boolean = False) Implements IFeatureTypeConverter.ConvertToKML
            Dim AltitudeModes(3) As String
            AltitudeModes(0) = "clampedToGround"
            AltitudeModes(1) = "relativeToGround"
            AltitudeModes(2) = "absolute"

            Dim xmlw As XmlTextWriter = New XmlTextWriter(FileName, Encoding.UTF8)

            Dim Xcoord, Ycoord As Single
            Dim nextline As String
            Dim pFCursor As IFeatureCursor
            Dim pFeature As IFeature
            Dim pGeometry, pGeometry1, pGeometry2 As IGeometry
            Dim pSpatRefFact As ISpatialReferenceFactory = New SpatialReferenceEnvironment
            Dim pSR As ISpatialReference = New GeographicCoordinateSystem
            pSR = pSpatRefFact.CreateGeographicCoordinateSystem(esriSRGeoCSType.esriSRGeoCS_WGS1984)
            pSR.SetDomain(-360, 360, -360, 360)
            Dim pTableFields As ITableFields = pFeatureLayer
            Dim pGeoLayer As IGeoFeatureLayer = pFeatureLayer
            Dim pTable As ITable = pFeatureLayer
            Dim pFClass As IFeatureClass = pFeatureLayer.FeatureClass
            Dim sLayername As String = pFeatureLayer.Name
            Dim GeomType As esriGeometryType = pFeatureLayer.FeatureClass.ShapeType
            Dim sFieldName As String
            Dim pAttribField As IField = Nothing

            Dim pQF() As IQueryFilter
            Dim pUVRend As IUniqueValueRenderer
            If (TypeOf pGeoLayer.Renderer Is IUniqueValueRenderer) And Distribute Then
                pUVRend = pGeoLayer.Renderer
                Dim i As Short
                Dim AttribStringType As Boolean
                Dim AllValues As String
                ReDim pQF(pUVRend.ValueCount - 1)
                sFieldName = pUVRend.Field(0)
                pAttribField = pTableFields.Field(pTableFields.FindField(sFieldName))
                AttribStringType = (pAttribField.Type = esriFieldType.esriFieldTypeString)
                If pUVRend.Value(0) <> "<Null>" Then
                    If Not AttribStringType Then
                        AllValues = CStr(pUVRend.Value(0))
                    Else
                        AllValues = "'" & pUVRend.Value(0) & "'"
                    End If
                Else
                    AllValues = "NULL"
                End If
                For i = 0 To pUVRend.ValueCount - 1
                    pQF(i) = New QueryFilter
                    If pUVRend.Value(I) <> "<Null>" Then
                        If Not AttribStringType Then
                            pQF(i).WhereClause = sFieldName & "=" & pUVRend.Value(I)
                        Else
                            pQF(i).WhereClause = sFieldName & "=" & "'" & pUVRend.Value(I) & "'"
                        End If
                    Else
                        pQF(i).WhereClause = sFieldName & " IS NULL"
                    End If
                    If I > 0 Then
                        If pUVRend.Value(0) <> "<Null>" Then
                            If Not AttribStringType Then
                                AllValues = AllValues & "," & CStr(pUVRend.Value(I))
                            Else
                                AllValues = AllValues & "," & "'" & pUVRend.Value(I) & "'"
                            End If
                        End If
                    Else
                        AllValues = AllValues & ",NULL"
                    End If
                Next i
                If pUVRend.UseDefaultSymbol Then
                    ReDim Preserve pQF(pUVRend.ValueCount)
                    pQF(pUVRend.ValueCount) = New QueryFilter
                    pQF(pUVRend.ValueCount).WhereClause = " NOT " & sFieldName & " IN (" & AllValues & ")"
                End If

            Else
                ReDim pQF(0)
                pQF(0) = New QueryFilter
                pQF(0) = Nothing
            End If

            xmlw.Formatting = Formatting.Indented
            xmlw.Indentation = 4
            xmlw.WriteStartDocument()
            xmlw.WriteStartElement("kml", "http://earth.google.com/kml/2.0")

            Select Case GeomType
                Case esriGeometryType.esriGeometryPolygon, esriGeometryType.esriGeometryPolyline
                    'write header tags
                    If PMFlag Or Distribute Then
                        xmlw.WriteStartElement("Folder")
                    Else
                        xmlw.WriteStartElement("Placemark")
                    End If
                    xmlw.WriteElementString("name", sLayername)
                    xmlw.WriteStartElement("description")
                    xmlw.WriteCData("Generated by <a href=" & Chr(34) & "http://dev.gis-lab.info/typeconvert" & Chr(34) & ">TypeConvert</a><br>" & Now())
                    xmlw.WriteEndElement()
                    xmlw.WriteElementString("visibility", "1")
                    xmlw.WriteElementString("open", "0")
                    If Not PMFlag And Not Distribute Then
                        'Здесь можно пустить стиль
                        xmlw.WriteStartElement("MultiGeometry")
                    End If

                    Dim qfIndex As Integer
                    For qfIndex = 0 To pQF.Length - 1
                        If pFClass.FeatureCount(pQF(qfIndex)) > 0 Then
                            If Not pQF(qfIndex) Is Nothing Then
                                If PMFlag Then
                                    xmlw.WriteStartElement("Folder")
                                Else
                                    xmlw.WriteStartElement("Placemark")
                                End If
                                If qfIndex = pUVRend.ValueCount Then
                                    xmlw.WriteElementString("name", "Other values")
                                Else
                                    If pUVRend.Label(pUVRend.Value(qfIndex)) = "<Null>" Then
                                        xmlw.WriteElementString("name", "")
                                    Else
                                        xmlw.WriteElementString("name", pUVRend.Label(pUVRend.Value(qfIndex)))
                                    End If
                                End If
                                If Not PMFlag Then
                                    xmlw.WriteStartElement("MultiGeometry")
                                End If
                            End If

                            pFCursor = pTable.Search(pQF(qfIndex), False)
                            pFeature = pFCursor.NextFeature

                            Dim pGeomColl As IGeometryCollection
                            Dim pSegColl As ISegmentCollection
                            Dim lngExtRing, lngIntRing As Long
                            Dim lSeg As Long
                            Dim pPolygon As IPolygon4
                            Dim GeometryCount As Integer
                            Dim sDesc As String
                            Dim i As Byte
                            'loop through features
                            Do While Not pFeature Is Nothing
                                If GeomType = esriGeometryType.esriGeometryPolyline Then
                                    pGeomColl = pFeature.Shape
                                Else
                                    pPolygon = pFeature.Shape
                                    pGeomColl = pPolygon.ExteriorRingBag
                                End If
                                GeometryCount = pGeomColl.GeometryCount

                                For lngExtRing = 0 To GeometryCount - 1
                                    If PMFlag Then
                                        xmlw.WriteStartElement("Placemark")
                                        xmlw.WriteElementString("name", Convert.ToString(pFeature.Value(pTableFields.FindField(LabelFieldName))))
                                        sDesc = ""
                                        For i = 0 To pTableFields.FieldCount - 1
                                            If pTableFields.FieldInfo(i).Visible AndAlso pTableFields.Field(i).Type <> esriFieldType.esriFieldTypeOID AndAlso pTableFields.Field(i).Type <> esriFieldType.esriFieldTypeGeometry And pTableFields.Field(i).Name <> LabelFieldName And Not pTableFields.Field(i) Is pAttribField Then
                                                sDesc = sDesc & pTableFields.FieldInfo(i).Alias & "=" & pFeature.Value(i) & "<br>"
                                            End If
                                        Next i
                                        xmlw.WriteStartElement("description")
                                        xmlw.WriteCData(sDesc)
                                        xmlw.WriteEndElement()
                                        ' Здесь можно пустить стиль
                                    End If
                                    If GeomType = esriGeometryType.esriGeometryPolygon Then
                                        xmlw.WriteStartElement("Polygon")
                                    Else
                                        xmlw.WriteStartElement("LineString")
                                    End If
                                    xmlw.WriteElementString("extrude", blnExtrude)
                                    xmlw.WriteElementString("tesselate", "1")
                                    xmlw.WriteElementString("altitudeMode", AltitudeModes(AltitudeMode))

                                    If GeomType = esriGeometryType.esriGeometryPolygon Then
                                        xmlw.WriteStartElement("outerBoundaryIs")
                                        xmlw.WriteStartElement("LinearRing")
                                    End If
                                    xmlw.WriteStartElement("coordinates")

                                    'write endpoints of each line segment
                                    Dim absheight As Single
                                    Dim fromPoint As IPoint
                                    Dim toPoint As IPoint

                                    Dim UseZValues As Boolean = False
                                    If blnExtrude = 1 Then
                                        If AttributeFieldName <> "" Then
                                            If AttributeFieldName <> "Z values" Then
                                                absheight = IIf(IsDBNull(pFeature.Value(pTableFields.FindField(AttributeFieldName))), 0, pFeature.Value(pTableFields.FindField(AttributeFieldName)))
                                            Else
                                                UseZValues = True
                                            End If
                                        Else
                                            absheight = SetValue
                                        End If
                                    End If

                                    pSegColl = pGeomColl.Geometry(lngExtRing)
                                    For lSeg = 0 To pSegColl.SegmentCount - 1
                                        If lSeg = 0 Then
                                            fromPoint = pSegColl.Segment(lSeg).FromPoint
                                            pGeometry1 = fromPoint
                                            If Not pGeometry1.SpatialReference Is Nothing Then
                                                pGeometry1.Project(pSR)
                                            End If
                                            Xcoord = fromPoint.X
                                            Ycoord = fromPoint.Y
                                            If UseZValues Then
                                                absheight = fromPoint.Z
                                            End If
                                            nextline = Xcoord & "," & Ycoord & "," & absheight
                                            xmlw.WriteString(nextline)
                                        End If
                                        toPoint = pSegColl.Segment(lSeg).ToPoint
                                        pGeometry2 = toPoint
                                        If Not pGeometry2.SpatialReference Is Nothing Then
                                            pGeometry2.Project(pSR)
                                        End If
                                        Xcoord = toPoint.X
                                        Ycoord = toPoint.Y
                                        If UseZValues Then
                                            absheight = toPoint.Z
                                        End If
                                        nextline = "," & Xcoord & "," & Ycoord & "," & absheight
                                        xmlw.WriteString(nextline)
                                    Next lSeg
                                    xmlw.WriteEndElement()

                                    If GeomType = esriGeometryType.esriGeometryPolygon Then
                                        xmlw.WriteEndElement()
                                        xmlw.WriteEndElement()
                                        If pPolygon.InteriorRingCount(pGeomColl.Geometry(lngExtRing)) > 0 Then
                                            Dim pIntGeomColl As IGeometryCollection
                                            pIntGeomColl = pPolygon.InteriorRingBag(pGeomColl.Geometry(lngExtRing))
                                            For lngIntRing = 0 To pPolygon.InteriorRingCount(pGeomColl.Geometry(lngExtRing)) - 1
                                                xmlw.WriteStartElement("innerBoundaryIs")
                                                xmlw.WriteStartElement("LinearRing")
                                                xmlw.WriteStartElement("coordinates")

                                                pSegColl = pIntGeomColl.Geometry(lngIntRing)
                                                For lSeg = 0 To pSegColl.SegmentCount - 1
                                                    If lSeg = 0 Then
                                                        Xcoord = pSegColl.Segment(lSeg).FromPoint.X
                                                        Ycoord = pSegColl.Segment(lSeg).FromPoint.Y
                                                        If UseZValues Then
                                                            absheight = fromPoint.Z
                                                        End If
                                                        nextline = Xcoord & "," & Ycoord & "," & absheight
                                                        xmlw.WriteString(nextline)
                                                    End If
                                                    Xcoord = pSegColl.Segment(lSeg).ToPoint.X
                                                    Ycoord = pSegColl.Segment(lSeg).ToPoint.Y
                                                    If UseZValues Then
                                                        absheight = toPoint.Z
                                                    End If
                                                    nextline = "," & Xcoord & "," & Ycoord & "," & absheight
                                                    xmlw.WriteString(nextline)
                                                Next lSeg

                                                xmlw.WriteEndElement()
                                                xmlw.WriteEndElement()
                                                xmlw.WriteEndElement()
                                            Next lngIntRing
                                        End If
                                        xmlw.WriteEndElement()
                                    Else
                                        xmlw.WriteEndElement()
                                    End If
                                    If PMFlag Then
                                        xmlw.WriteEndElement()
                                    End If
                                Next lngExtRing
                                pFeature = pFCursor.NextFeature
                            Loop
                            If Not pQF(qfIndex) Is Nothing Then
                                If PMFlag Then
                                    xmlw.WriteEndElement()
                                Else
                                    xmlw.WriteEndElement()
                                    xmlw.WriteEndElement()
                                End If
                            End If
                        End If
                    Next qfIndex

                    'write footer
                    If PMFlag Or Distribute Then
                        xmlw.WriteEndElement()
                    Else
                        xmlw.WriteEndElement()
                        xmlw.WriteEndElement()
                    End If

                Case esriGeometryType.esriGeometryPoint

                    Dim absheight As Single

                    xmlw.WriteStartElement("Folder")
                    xmlw.WriteElementString("name", sLayername)
                    xmlw.WriteStartElement("description")
                    xmlw.WriteCData("Generated by <a href=" & Chr(34) & "http://dev.gis-lab.info/typeconvert" & Chr(34) & ">TypeConvert</a><br>" & Now())
                    xmlw.WriteEndElement()
                    xmlw.WriteElementString("visibility", "1")
                    xmlw.WriteElementString("open", "0")
                    Dim qfIndex As Integer
                    For qfIndex = 0 To pQF.Length - 1
                        If pFClass.FeatureCount(pQF(qfIndex)) > 0 Then
                            If Not pQF(qfIndex) Is Nothing Then
                                xmlw.WriteStartElement("Folder")
                                If qfIndex = pUVRend.ValueCount Then
                                    xmlw.WriteElementString("name", "Other values")
                                Else
                                    If pUVRend.Label(pUVRend.Value(qfIndex)) = "<Null>" Then
                                        xmlw.WriteElementString("name", "")
                                    Else
                                        xmlw.WriteElementString("name", pUVRend.Label(pUVRend.Value(qfIndex)))
                                    End If
                                End If
                            End If

                            pFCursor = pTable.Search(pQF(qfIndex), False)
                            pFeature = pFCursor.NextFeature

                            'Dim pArea As IArea
                            Dim pPoint As IPoint
                            Dim sValue As String
                            Dim sDesc As String
                            Dim i As Byte
                            Do While Not pFeature Is Nothing
                                xmlw.WriteElementString("extrude", blnExtrude)
                                xmlw.WriteElementString("altitudeMode", AltitudeModes(AltitudeMode))

                                Dim UseZValues As Boolean = False
                                If blnExtrude = 1 Then
                                    If AttributeFieldName <> "" Then
                                        If AttributeFieldName <> "Z values" Then
                                            absheight = IIf(IsDBNull(pFeature.Value(pTableFields.FindField(AttributeFieldName))), 0, pFeature.Value(pTableFields.FindField(AttributeFieldName)))
                                        Else
                                            UseZValues = True
                                        End If
                                    Else
                                        absheight = SetValue
                                    End If
                                End If

                                sValue = Convert.ToString(pFeature.Value(pTableFields.FindField(LabelFieldName)))
                                sDesc = ""
                                For i = 0 To pTableFields.FieldCount - 1
                                    If pTableFields.FieldInfo(i).Visible AndAlso pTableFields.Field(i).Type <> esriFieldType.esriFieldTypeOID AndAlso pTableFields.Field(i).Type <> esriFieldType.esriFieldTypeGeometry AndAlso pTableFields.Field(i).Name <> LabelFieldName And Not pTableFields.Field(i) Is pAttribField Then
                                        sDesc = sDesc & pTableFields.FieldInfo(i).Alias & "=" & pFeature.Value(i) & "<br>"
                                    End If
                                Next i

                                xmlw.WriteStartElement("Placemark")
                                xmlw.WriteElementString("name", sValue)
                                xmlw.WriteStartElement("description")
                                xmlw.WriteCData(sDesc)
                                xmlw.WriteEndElement()
                                xmlw.WriteStartElement("Point")
                                xmlw.WriteStartElement("coordinates")

                                'pArea = pFeature.Shape
                                pPoint = pFeature.Shape
                                pGeometry = pPoint
                                pGeometry.Project(pSR)
                                Xcoord = pPoint.X
                                Ycoord = pPoint.Y
                                If UseZValues Then
                                    absheight = pPoint.Z
                                End If
                                nextline = "," & Xcoord & "," & Ycoord & "," & absheight
                                xmlw.WriteString(nextline)

                                xmlw.WriteEndElement()
                                xmlw.WriteEndElement()
                                xmlw.WriteEndElement()

                                pFeature = pFCursor.NextFeature
                            Loop
                            If Not pQF(qfIndex) Is Nothing Then
                                xmlw.WriteEndElement()
                            End If
                        End If
                    Next qfIndex
                    'write footer
                    xmlw.WriteEndElement()

            End Select

            xmlw.WriteEndDocument()
            xmlw.Close()

        End Sub
        Public Sub New()
            MyBase.New()
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
End Namespace