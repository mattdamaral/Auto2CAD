Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class LoadTableClass

    Public circList As List(Of CircuitClass)
    Public DRList As New List(Of DRClass)

    Public totalDesc As String
    Public totalLoad As Double
    Public totalPhases As String
    Public totalR As Double
    Public totalS As Double
    Public totalT As Double
    Public totalWire As String
    Public totalBreaker As String

    Public fromCSV As Boolean = False

    Public Sub New(circListAux As List(Of CircuitClass))
        circList = circListAux
        DistributeLoads()

        totalDesc = "QDXX - Description"
        totalLoad = 0
        totalR = 0
        totalS = 0
        totalT = 0
        totalWire = "-"
        totalBreaker = "-"
        GetTotal()

    End Sub

    Public Sub New(circListAux As List(Of CircuitClass), totalDescAux As String, totalLoadAux As Double, totalPhasesAux As String, totalRAux As Double, totalSAux As Double, totalTAux As Double, totalWireAux As Double, totalBreakerAux As String)
        circList = circListAux
        fromCSV = True

        totalDesc = totalDescAux
        totalLoad = totalLoadAux
        totalPhases = totalPhasesAux
        totalR = totalRAux
        totalS = totalSAux
        totalT = totalTAux
        totalWire = totalWireAux
        totalBreaker = totalBreakerAux

    End Sub

    ' Possibily obsolete ************************************************************
    'Public Sub New(circListAux As List(Of CircuitClass), fromCSVAux As Boolean)
    '    circList = circListAux
    '    fromCSV = fromCSVAux

    '    totalDesc = "QDXX - Description"
    '    totalLoad = 0
    '    totalR = 0
    '    totalS = 0
    '    totalT = 0
    '    totalWire = "-"
    '    totalBreaker = "-"

    'End Sub

    Public Sub DistributeLoads()   ' Distributes the each circuit's load beetween phases R, S and T

        'MsgBox(circList.Count - 1)

        For index = 0 To (circList.Count - 1)   ' Runs through the circuits in the load table

            Dim currentCirc As CircuitClass = circList(index)

            ' If it's the first circuit, don't check the previous circuits ('cause there ain't none)
            If index = 0 Then

                Select Case currentCirc.connection
                    Case "Mono", ""
                        currentCirc.r = currentCirc.load
                        currentCirc.s = 0
                        currentCirc.t = 0
                        currentCirc.phases = "R"
                        GoTo end_of_for_01
                    Case "Bi"
                        currentCirc.r = currentCirc.load / 2
                        currentCirc.s = currentCirc.load / 2
                        currentCirc.t = 0
                        currentCirc.phases = "R+S"
                        GoTo end_of_for_01
                    Case "Tri"
                        currentCirc.r = currentCirc.load / 3
                        currentCirc.s = currentCirc.load / 3
                        currentCirc.t = currentCirc.load / 3
                        currentCirc.phases = "R+S+T"
                        GoTo end_of_for_01
                End Select

            Else                                                    ' If else checks the previous circuits, unless it's a three phase

                If currentCirc.connection.Contains("Tri") Then

                    currentCirc.r = currentCirc.load / 3
                    currentCirc.s = currentCirc.load / 3
                    currentCirc.t = currentCirc.load / 3
                    currentCirc.phases = "R+S+T"
                    GoTo end_of_for_01

                Else

                    For previousIndex = (index - 1) To 0 Step -1

                        Dim previousCirc As CircuitClass = circList(previousIndex)

                        If previousCirc.r > 0 And previousCirc.s > 0 And previousCirc.t > 0 Then

                            GoTo end_of_for_02

                        Else

                            If previousCirc.r > 0 Then

                                If previousCirc.s > 0 Then

                                    Select Case currentCirc.connection
                                        Case "Mono", ""
                                            currentCirc.r = 0
                                            currentCirc.s = 0
                                            currentCirc.t = currentCirc.load
                                            currentCirc.phases = "T"
                                            GoTo end_of_for_01
                                        Case "Bi"
                                            currentCirc.r = currentCirc.load / 2
                                            currentCirc.s = 0
                                            currentCirc.t = currentCirc.load / 2
                                            currentCirc.phases = "R+T"
                                            GoTo end_of_for_01
                                    End Select

                                Else

                                    Select Case currentCirc.connection
                                        Case "Mono", ""
                                            currentCirc.r = 0
                                            currentCirc.s = currentCirc.load
                                            currentCirc.t = 0
                                            currentCirc.phases = "S"
                                            GoTo end_of_for_01
                                        Case "Bi"
                                            currentCirc.r = 0
                                            currentCirc.s = currentCirc.load / 2
                                            currentCirc.t = currentCirc.load / 2
                                            currentCirc.phases = "S+T"
                                            GoTo end_of_for_01
                                    End Select

                                End If

                            ElseIf previousCirc.s > 0 Then

                                If previousCirc.t > 0 Then

                                    Select Case currentCirc.connection
                                        Case "Mono", ""
                                            currentCirc.r = currentCirc.load
                                            currentCirc.s = 0
                                            currentCirc.t = 0
                                            currentCirc.phases = "R"
                                            GoTo end_of_for_01
                                        Case "Bi"
                                            currentCirc.r = currentCirc.load / 2
                                            currentCirc.s = currentCirc.load / 2
                                            currentCirc.t = 0
                                            currentCirc.phases = "R+S"
                                            GoTo end_of_for_01
                                    End Select

                                Else

                                    Select Case currentCirc.connection
                                        Case "Mono", ""
                                            currentCirc.r = 0
                                            currentCirc.s = 0
                                            currentCirc.t = currentCirc.load
                                            currentCirc.phases = "T"
                                            GoTo end_of_for_01
                                        Case "Bi"
                                            currentCirc.r = currentCirc.load / 2
                                            currentCirc.s = 0
                                            currentCirc.t = currentCirc.load / 2
                                            currentCirc.phases = "R+T"
                                            GoTo end_of_for_01
                                    End Select

                                End If

                            Else

                                Select Case currentCirc.connection
                                    Case "Mono", ""
                                        currentCirc.r = currentCirc.load
                                        currentCirc.s = 0
                                        currentCirc.t = 0
                                        currentCirc.phases = "R"
                                        GoTo end_of_for_01
                                    Case "Bi"
                                        currentCirc.r = currentCirc.load / 2
                                        currentCirc.s = currentCirc.load / 2
                                        currentCirc.t = 0
                                        currentCirc.phases = "R+S"
                                        GoTo end_of_for_01
                                End Select

                            End If

                        End If

end_of_for_02:

                    Next

                End If

            End If

end_of_for_01:

        Next

    End Sub

    ' Adds the Loads to get the total ---------------------------------------------------------------------------------------------------------------

    Private Sub GetTotal()

        For index = 0 To circList.Count - 1

            totalLoad += circList(index).load
            totalR += circList(index).r
            totalS += circList(index).s
            totalT += circList(index).t

        Next

        If totalLoad > 0 And totalLoad <= 25000 Then
            totalWire = "6"
            totalBreaker = "30 A - C"
        ElseIf totalLoad > 25000 And totalLoad <= 30000 Then
            totalWire = "10"
            totalBreaker = "40 A - C"
        ElseIf totalLoad > 30000 And totalLoad <= 35000 Then
            totalWire = "10"
            totalBreaker = "50 A - C"
        ElseIf totalLoad > 35000 And totalLoad <= 40000 Then
            totalWire = "16"
            totalBreaker = "60 A - C"
        ElseIf totalLoad > 40000 And totalLoad <= 50000 Then
            totalWire = "25"
            totalBreaker = "70 A - C"
        ElseIf totalLoad > 50000 And totalLoad <= 65000 Then
            totalWire = "35"
            totalBreaker = "100 A - C"
        ElseIf totalLoad > 65000 And totalLoad <= 75000 Then
            totalWire = "50"
            totalBreaker = "125 A - C"
        ElseIf totalLoad > 75000 Then
            totalWire = "-"
            totalBreaker = "-"
        End If

    End Sub


    ' Creates the Table itself ----------------------------------------------------------------------------------------------------------------------
    <Obsolete>
    Public Sub CreateTable()

        ' Accessess the database -------------------------------------------------------------------------------------------
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim pr As PromptPointResult = ed.GetPoint("Enter Table insertion point: ")  ' Asks the user for the table's insertion point

        If pr.Status = PromptStatus.OK Then

            Dim table As Table = New Table()
            'Dim circuits As List(Of CircuitClass) = loadTableAux.circList
            Dim circuits As List(Of CircuitClass) = circList

            With table

                Dim headerRows As Integer = 2                                                           ' Amount of Rows for the Header
                Dim totalRows As Integer = 1                                                           ' Amount of Rows for the Total
                Dim rows As Integer = circuits.Count + headerRows + totalRows                            ' Amount of Rows (w/ header included)
                Dim columnTextList() As String = {"Circuito", "Descrição", "Potência (W)", "R (W)", "S (W)", "T (W)", "Seção (mm²)", "Disjuntor"}
                Dim columns As Integer = UBound(columnTextList) - LBound(columnTextList) + 1            ' Amount of Columns

                .SetSize(rows, columns)
                .SetRowHeight(16)
                .SetColumnWidth(100)
                .Position = pr.Value

                .SetTextHeight(0, 0, 8)
                .SetAlignment(0, 0, CellAlignment.MiddleCenter)
                .SetTextString(0, 0, "Quadro de Cargas")

                For indexColumn = 0 To (columns - 1)

                    .SetTextHeight(1, indexColumn, 8)
                    .SetAlignment(1, indexColumn, CellAlignment.MiddleCenter)
                    .SetTextString(1, indexColumn, columnTextList(indexColumn))

                Next

                For indexRow = headerRows To (rows - totalRows - 1)

                    For indexColumn = 0 To (columns - 1)

                        .SetTextHeight(indexRow, indexColumn, 8)
                        .SetAlignment(indexRow, indexColumn, CellAlignment.MiddleCenter)

                        Dim circRow = indexRow - headerRows

                        Select Case indexColumn
                            Case 0
                                .SetTextString(indexRow, indexColumn, circuits(circRow).circ.ToString())
                            Case 1
                                .SetTextString(indexRow, indexColumn, "Descrição")
                                .SetColumnWidth(indexColumn, 300)
                            Case 2
                                .SetTextString(indexRow, indexColumn, circuits(circRow).load.ToString())
                            Case 3
                                .SetTextString(indexRow, indexColumn, circuits(circRow).r.ToString())
                            Case 4
                                .SetTextString(indexRow, indexColumn, circuits(circRow).s.ToString())
                            Case 5
                                .SetTextString(indexRow, indexColumn, circuits(circRow).t.ToString())
                            Case 6
                                .SetTextString(indexRow, indexColumn, circuits(circRow).wire.ToString())
                            Case 7
                                .SetTextString(indexRow, indexColumn, circuits(circRow).breaker.ToString())
                        End Select

                    Next

                Next

                ' Populates the Total row (last row)
                For indexColumn = 0 To (columns - 1)
                    .SetTextHeight(rows - totalRows, indexColumn, 8)
                    .SetAlignment(rows - totalRows, indexColumn, CellAlignment.MiddleCenter)
                Next
                .SetTextString(rows - totalRows, 0, "Total")
                Dim range As CellRange = CellRange.Create(table, rows - totalRows, 0, rows - totalRows, 1)
                .MergeCells(range)
                '.SetTextString(rows - totalRows, 2, loadTableAux.totalLoad.ToString)
                '.SetTextString(rows - totalRows, 3, loadTableAux.totalR.ToString)
                '.SetTextString(rows - totalRows, 4, loadTableAux.totalS.ToString)
                '.SetTextString(rows - totalRows, 5, loadTableAux.totalT.ToString)
                '.SetTextString(rows - totalRows, 6, loadTableAux.totalWire.ToString)
                '.SetTextString(rows - totalRows, 7, loadTableAux.totalBreaker.ToString)
                .SetTextString(rows - totalRows, 2, totalLoad.ToString)
                .SetTextString(rows - totalRows, 3, totalR.ToString)
                .SetTextString(rows - totalRows, 4, totalS.ToString)
                .SetTextString(rows - totalRows, 5, totalT.ToString)
                .SetTextString(rows - totalRows, 6, totalWire.ToString)
                .SetTextString(rows - totalRows, 7, totalBreaker.ToString)

                table.GenerateLayout()

            End With

            ' Creates a transaction for the database ---------------------------------------------------------------------------
            Using trans As Transaction = doc.TransactionManager.StartTransaction()

                Call Commands.ChangeLayer("MD - Quadro de Cargas")

                Dim lt As LayerTable
                Dim ltr As New LayerTableRecord
                Dim layerID As ObjectId
                Dim layerName As String = "MD - Quadro de Cargas"

                ' Checks if layer exists
                Try
                    ' If layer exists, get layerid
                    lt = CType(trans.GetObject(db.LayerTableId, OpenMode.ForRead, True, True), LayerTable)
                    layerID = lt.Item(layerName)

                    ' If the layer was deleted, recover layer
                    If layerID.IsErased Then
                        lt.UpgradeOpen()
                        lt.Item(layerName).GetObject(OpenMode.ForWrite, True, True).Erase(False)
                    End If

                Catch ex As Exception
                    ' If the layer doesn't exist, create it
                    lt = db.LayerTableId.GetObject(OpenMode.ForWrite, True, True)
                    ltr.Name = layerName
                    lt.Add(ltr)
                    ' Adds layer to db
                    trans.AddNewlyCreatedDBObject(ltr, True)
                    ' Recovers layerid of newly created layer
                    lt = CType(trans.GetObject(db.LayerTableId, OpenMode.ForRead, False), LayerTable)
                    layerID = lt.Item(layerName)
                End Try

                ' Sets layer as current
                db.Clayer = layerID

                Dim bt As BlockTable = CType(trans.GetObject(doc.Database.BlockTableId, OpenMode.ForRead), BlockTable)

                Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                btr.AppendEntity(table)

                trans.AddNewlyCreatedDBObject(table, True)

                trans.Commit()

            End Using

        Else

        End If

    End Sub

    <Obsolete>
    Public Sub CreateDiagram()
        'Public Sub CreateDiagram(loadTableAux As LoadTableClass)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim pr As PromptPointResult = ed.GetPoint("Enter the Diagram insertion point: ")  ' Asks the user for the table's insertion point

        If pr.Status = PromptStatus.OK Then

            Using trans As Transaction = db.TransactionManager.StartTransaction()

                Call Commands.ChangeLayer("MD - Diagrama Unifilar")

                'Dim circuits As List(Of CircuitClass) = loadTableAux.circList
                Dim circuits As List(Of CircuitClass) = circList

                Dim blkPos As Point3d = New Point3d(pr.Value.X, pr.Value.Y, pr.Value.Z)

                ' Opens the Block table for read
                Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                For index = 0 To (circuits.Count - 1)
                    Dim blkID As ObjectId = ObjectId.Null

                    If Not bt.Has("MD - DU Circuito") Then
                        ' Add block from file

                    Else
                        blkID = bt("MD - DU Circuito")

                    End If

                    ' Create and insert the new block reference
                    If blkID <> ObjectId.Null Then
                        Dim blkPosCurrent As Point3d = New Point3d(pr.Value.X, pr.Value.Y + (index * (-60)), pr.Value.Z)

                        'Dim btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)
                        Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)
                            'Dim blkName As String = "MD - DU Circuito"

                            Using blkRef As New BlockReference(blkPosCurrent, btr.Id)

                                Dim curBtr As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                                curBtr.AppendEntity(blkRef)
                                trans.AddNewlyCreatedDBObject(blkRef, True)

                                ' Verify block table record has attribute definitions associated with it
                                If btr.HasAttributeDefinitions Then

                                    ' Add attributes from the block table record
                                    For Each objID As ObjectId In btr

                                        Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)

                                        If TypeOf dbObj Is AttributeDefinition Then

                                            Dim attDef As AttributeDefinition = dbObj

                                            If Not attDef.Constant Then

                                                Using attRef As New AttributeReference

                                                    attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                                    attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                                    Select Case attRef.Tag
                                                        Case "CIRC_DISJUNTOR"
                                                            'If loadTableAux.fromCSV = True Then
                                                            If fromCSV = True Then
                                                                If circuits(index).desc.Contains("Iluminação") Then
                                                                    attRef.TextString = circuits(index).breaker.ToString & " A - B"
                                                                Else
                                                                    attRef.TextString = circuits(index).breaker.ToString & " A - C"
                                                                End If
                                                            Else
                                                                attRef.TextString = circuits(index).breaker.ToString
                                                            End If
                                                            Exit Select
                                                        Case "CIRC_SEÇÃO"
                                                            'If loadTableAux.fromCSV = True Then
                                                            If fromCSV = True Then
                                                                attRef.TextString = circuits(index).wire.ToString & " mm²"
                                                            Else
                                                                attRef.TextString = circuits(index).wire.ToString
                                                            End If
                                                            Exit Select
                                                        Case "CIRC_POTÊNCIA"
                                                            attRef.TextString = "(" + circuits(index).load.ToString + " W)"
                                                            Exit Select
                                                        Case "CIRC_FASE"
                                                            attRef.TextString = circuits(index).phases
                                                            Exit Select
                                                        Case "CIRC_N_CIRCUITO"
                                                            attRef.TextString = circuits(index).circ.ToString
                                                            Exit Select
                                                        Case "CIRC_DESCRIÇÃO"
                                                            'If loadTableAux.fromCSV = True Then
                                                            If fromCSV = True Then
                                                                attRef.TextString = "(" & circuits(index).desc.ToString & ")"
                                                            Else
                                                                attRef.TextString = "(Description)"
                                                            End If
                                                            Exit Select
                                                    End Select

                                                    ' Add DU block to the block table record and to the transaction
                                                    blkRef.AttributeCollection.AppendAttribute(attRef)
                                                    trans.AddNewlyCreatedDBObject(attRef, True)

                                                End Using
                                            End If
                                        End If
                                    Next
                                End If
                            End Using
                        End Using
                    End If
                Next
                'Dim descText As DBText
                'Dim totalLoadText As DBText
                'Dim frame As Polyline

                Using btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    ' Creates the Total Load And Description texts
                    Dim descText As DBText = New DBText
                    descText.SetDatabaseDefaults()
                    descText.Position = New Point3d((blkPos.X - 200), (blkPos.Y + 65), 0)
                    descText.Height = 15
                    'descText.TextString = loadTableAux.totalDesc
                    descText.TextString = totalDesc
                    btr.AppendEntity(descText)
                    trans.AddNewlyCreatedDBObject(descText, True)

                    Dim totalLoadText As DBText = New DBText
                    totalLoadText.SetDatabaseDefaults()
                    totalLoadText.Position = New Point3d((blkPos.X - 195), (blkPos.Y + 45), 0)
                    totalLoadText.Height = 10
                    'totalLoadText.TextString = "(" & loadTableAux.totalLoad.ToString & " W)"
                    totalLoadText.TextString = "(" & totalLoad.ToString & " W)"
                    btr.AppendEntity(totalLoadText)
                    trans.AddNewlyCreatedDBObject(totalLoadText, True)

                    ' Creates the dashed frame around the circuits
                    ' Initializes a New polyline with 5 vertices
                    Dim frame As Polyline = New Polyline
                    frame.AddVertexAt(0, New Point2d(blkPos.X - 200, blkPos.Y + 60), 0, 0, 0)
                    frame.AddVertexAt(1, New Point2d(blkPos.X + 125, blkPos.Y + 60), 0, 0, 0)
                    frame.AddVertexAt(2, New Point2d(blkPos.X + 125, blkPos.Y - (60 * (circuits.Count - 1) + 200)), 0, 0, 0)
                    frame.AddVertexAt(3, New Point2d(blkPos.X - 200, blkPos.Y - (60 * (circuits.Count - 1) + 200)), 0, 0, 0)
                    frame.AddVertexAt(4, New Point2d(blkPos.X - 200, blkPos.Y + 60), 0, 0, 0)
                    frame.Closed = True
                    btr.AppendEntity(frame)
                    trans.AddNewlyCreatedDBObject(frame, True)

                    'descText = Nothing
                    'totalLoadText = Nothing
                    'frame = Nothing

                    If fromCSV = True Then
                        SortDRs(circuits, blkPos)
                    End If

                End Using

                ' Saves the new object to the database
                trans.Commit()

                ' Disposes of the transaction
            End Using
        End If
    End Sub

    Public Sub SortDRs(circuits As List(Of CircuitClass), blkPos As Point3d)
        Dim nofDR As Integer = circuits.Count ' Number of DRs in the load table
        Dim drArray(nofDR) As String ' Array with all the DR's values in the load table
        'Dim drStartFinish(nofDR, 2) ' Where the DRs start and finish (related to the circuits they protect)
        Dim busStartPoint As Point3d ' Where the circuits' buses start
        Dim busEndPoint As Point3d ' Where the circuits' buses finish
        'Dim mainBusStartPoint As Point3d ' Where the bus that connects all the DRs starts
        'Dim mainBusEndPoint As Point3d ' Where the bus that connects all the DRs ends
        Dim currentDRValue As String
        Dim protEntradaStartPoint As Point3d = Nothing
        Dim protEntradaEndPoint As Point3d = Nothing
        Dim DRType As String = "DDR"

        ' Passes the values of the DR's from the circuits to the DR array
        For index = 0 To (nofDR - 1)
            drArray(index) = circuits(index).dr
        Next

        ' Defines the circuits that each DR protects
        Dim nextDR As String = "-"
        Dim newDRAhead As Boolean = True
        For index = 0 To (nofDR - 1)
            Dim currentDR As String = drArray(index)

            'If nofDR > index Then
            '    nextDR = drArray(index + 1)
            'End If

            If index = 0 Then
                busStartPoint = New Point3d(blkPos.X, blkPos.Y + 20, blkPos.Z)
                currentDRValue = currentDR
                newDRAhead = False
                If (nofDR - 1) > index Then
                    nextDR = drArray(index + 1)

                    If currentDR IsNot nextDR And nextDR IsNot "" Then
                        busEndPoint = New Point3d(blkPos.X, blkPos.Y - 20, blkPos.Z)
                        Call CommonFunctions.DrawLine(busStartPoint, busEndPoint, 6)
                        'Call DRClass.DrawDR(currentDRValue, New Point3d(busStartPoint.X, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z))
                        DRList.Add(New DRClass(circList, currentDRValue, DRType, New Point3d(busStartPoint.X, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z)))
                        newDRAhead = True

                        If protEntradaStartPoint = Nothing Then
                            protEntradaStartPoint = New Point3d(busStartPoint.X - 55, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z)
                        Else
                            protEntradaEndPoint = New Point3d(busStartPoint.X - 55, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z)
                        End If

                    End If
                End If
            Else
                If newDRAhead = True Then
                    busStartPoint = New Point3d(blkPos.X, blkPos.Y + (index * (-60)) + 20, blkPos.Z)
                    currentDRValue = currentDR
                    newDRAhead = False
                End If

                If (nofDR - 1) > index Then
                    nextDR = drArray(index + 1)

                    If currentDR IsNot nextDR And nextDR IsNot "" Then
                        busEndPoint = New Point3d(blkPos.X, blkPos.Y + (index * (-60)) - 20, blkPos.Z)
                        Call CommonFunctions.DrawLine(busStartPoint, busEndPoint, 6)
                        'Call DRClass.DrawDR(currentDRValue, New Point3d(busStartPoint.X, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z))
                        DRList.Add(New DRClass(circList, currentDRValue, DRType, New Point3d(busStartPoint.X, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z)))
                        newDRAhead = True

                        If protEntradaStartPoint = Nothing Then
                            protEntradaStartPoint = New Point3d(busStartPoint.X - 55, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z)
                        Else
                            protEntradaEndPoint = New Point3d(busStartPoint.X - 55, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z)
                        End If

                    Else
                        newDRAhead = False
                    End If
                Else
                    busEndPoint = New Point3d(blkPos.X, blkPos.Y + (index * (-60)) - 20, blkPos.Z)
                    Call CommonFunctions.DrawLine(busStartPoint, busEndPoint, 6)
                    'Call DRClass.DrawDR(currentDRValue, New Point3d(busStartPoint.X, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z))
                    DRList.Add(New DRClass(circList, currentDRValue, DRType, New Point3d(busStartPoint.X, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z)))

                    If protEntradaStartPoint = Nothing Then
                        protEntradaStartPoint = New Point3d(busStartPoint.X - 55, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z)
                    Else
                        protEntradaEndPoint = New Point3d(busStartPoint.X - 55, (busStartPoint.Y + busEndPoint.Y) / 2, busStartPoint.Z)
                    End If

                End If
            End If
        Next

        For index = 0 To (DRList.Count - 1)

            'Dim DR As DRClass
            'DR = DRList(index)
            DRList(index).DrawDR()

        Next

        Call CommonFunctions.DrawLine(protEntradaStartPoint, protEntradaEndPoint, 7)
        DrawMainProtection(New Point3d(protEntradaStartPoint.X, (protEntradaStartPoint.Y + protEntradaEndPoint.Y) / 2, protEntradaStartPoint.Z), totalBreaker, totalWire, "??")

    End Sub


    ' Draws the main protection
    Public Sub DrawMainProtection(protEntradaPos As Point3d, breaker As String, wire As Double, nomeEntrada As String)
        Call Commands.ChangeLayer("MD - Diagrama Unifilar")

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            ' Opens the Block table for read
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            Dim blkID As ObjectId = ObjectId.Null

            If Not bt.Has("MD - DU Proteção Entrada") Then
                ' Add block from file
            Else
                blkID = bt("MD - DU Proteção Entrada")
            End If

            ' Create and insert the new block reference
            If blkID <> ObjectId.Null Then

                Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)
                    Using blkRef As New BlockReference(protEntradaPos, btr.Id)

                        Dim curBtr As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                        curBtr.AppendEntity(blkRef)
                        trans.AddNewlyCreatedDBObject(blkRef, True)

                        ' Verify block table record has attribute definitions associated with it
                        If btr.HasAttributeDefinitions Then

                            ' Add attributes from the block table record
                            For Each objID As ObjectId In btr

                                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)

                                If TypeOf dbObj Is AttributeDefinition Then

                                    Dim attDef As AttributeDefinition = dbObj

                                    If Not attDef.Constant Then

                                        Using attRef As New AttributeReference

                                            attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                            attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                            Select Case attRef.Tag
                                                Case "DISJUNTOR"
                                                    attRef.TextString = breaker.ToString & " A - C"
                                                    Exit Select
                                                Case "SEÇÃO"
                                                    attRef.TextString = "[3#" & wire.ToString & "(" & wire.ToString & ")" & wire.ToString & "] mm²"
                                                    Exit Select
                                                Case "NOME_ENTRADA"
                                                    attRef.TextString = nomeEntrada.ToString
                                                    Exit Select
                                            End Select

                                            ' Add DU block to the block table record and to the transaction
                                            blkRef.AttributeCollection.AppendAttribute(attRef)
                                            trans.AddNewlyCreatedDBObject(attRef, True)

                                        End Using
                                    End If
                                End If
                            Next
                        End If
                    End Using
                End Using
            End If

            trans.Commit()

        End Using
    End Sub
End Class
