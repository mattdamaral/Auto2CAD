Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class LoadTable

    '<CommandMethod("loadtable")>
    <CommandMethod("zz")>
    Public Sub CreateLoadTable()

        Dim loadTableAux As LoadTableObj = New LoadTableObj(GetList())

        '' Displays all the circuits in circList
        'For index = 0 To (CircObjAux.Count - 1)

        '    MsgBox("Circuit: " + CircObjAux(index).circ.ToString + vbNewLine + "Load: " + CircObjAux(index).load.ToString)

        'Next

        CreateTable(loadTableAux)

        CreateDiagram(loadTableAux)

    End Sub

    ' Gets the circuits and loads -------------------------------------------------------------------------------------------------------------------
    Public Function GetList() As List(Of CircObj) ' GetList(0) returns circList, while GetList(1) return loadList

        Dim circList As New List(Of Integer)()
        Dim loadList As New List(Of Integer)()
        Dim circLoadList As New List(Of CircObj)()

        ' Accessess the database -------------------------------------------------------------------------------------------
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Creates a transaction for the database ---------------------------------------------------------------------------
        Using trans As Transaction = doc.TransactionManager.StartTransaction()

            ' Filter for the selection
            Dim tv As TypedValue() = New TypedValue(0) {}
            tv.SetValue(New TypedValue(CInt(DxfCode.Start), "INSERT"), 0)
            Dim filter As SelectionFilter = New SelectionFilter(tv)

            ' Selects objects using filter
            Dim psr As PromptSelectionResult = ed.GetSelection(filter)

            ' Checks if selection is ok
            If psr.Status = PromptStatus.OK Then

                Dim ss As SelectionSet = psr.Value

                ' Repeats code for every item selected
                For Each so As SelectedObject In ss

                    ' Blocks parameters
                    Dim br As BlockReference = CType(trans.GetObject(so.ObjectId, OpenMode.ForRead), BlockReference)
                    Dim ac As AttributeCollection = br.AttributeCollection

                    ' Attribute related parameters
                    Dim circAux As Integer = Nothing
                    Dim loadAux As Integer = Nothing
                    Dim connAux As String = ""

                    ' Fills the attribute list (attList) with attribute references
                    For Each oid As ObjectId In ac

                        Dim attR As AttributeReference = CType(trans.GetObject(oid, OpenMode.ForRead), AttributeReference)  ' Stores the attribute
                        'MsgBox("Att Tag: " + attR.Tag.ToString + vbNewLine + "Att Value: " + attR.TextString)               ' Displays the attribute tag and value

                        'Dim attRTag As String = attR.Tag            ' Stores the attribute's tag as string
                        'Dim attRValue As String = attR.TextString   ' Stores the attribute's value as string

                        ' Checks if the attribute's tag is either "CIRCUITO" or "POTÊNCIA"
                        If attR.Tag.Contains("CIRCUITO") Then

                            Integer.TryParse(attR.TextString, circAux)             ' If the attribute's value is an intenger, save it to circAux

                        ElseIf attR.Tag.Contains("POTÊNCIA") Then

                            Integer.TryParse(attR.TextString, loadAux)      ' If the attribute's value is an intenger, save it to loadAux

                        ElseIf attR.Tag.Contains("CONNECTION") Then

                            Integer.TryParse(attR.TextString, connAux)

                        End If

                        'MsgBox("circAux: " + circAux.ToString + vbNewLine + "loadAux: " + loadAux.ToString)

                    Next

                    ' Checks if circAux and loadAux were assigned a value
                    If circAux > 0 And loadAux > 0 Then

                        'MsgBox("List size: " + circList.Count.ToString + vbNewLine + "Circuit: " + circAux.ToString + vbNewLine + "Load: " + loadAux.ToString)

                        If circLoadList.Count > 0 Then                  ' If circuits list is not empty, either adds the load to an existing circuit or create a new one

                            Dim foundSimiliarCirc = False

                            For index = 0 To (circLoadList.Count - 1)   ' Checks if there's a circuit in circList with the same value as the attribute's value

                                If circAux = circLoadList(index).circ Then

                                    foundSimiliarCirc = True
                                    circLoadList(index).load += loadAux

                                    GoTo goto_01

                                End If

                            Next

                            If foundSimiliarCirc = False Then      ' If there's no circAux in circList, adds a new one to the end of the list

                                circLoadList.Add(New CircObj(circAux, loadAux, connAux))

                            End If

                        Else                                       ' If circList is empty, adds a new circuit and assigns the load to it

                            circLoadList.Add(New CircObj(circAux, loadAux, connAux))

                        End If

goto_01:

                    End If

                Next

                circLoadList.Sort(Function(x, y) x.circ.CompareTo(y.circ))

                '' Displays all the circuits in circList
                'For index = 0 To (circList.Count - 1)

                '    MsgBox("Circuit: " + circList(index).ToString + vbNewLine + "Load: " + loadList(index).ToString)

                'Next

            Else    ' Displays this message in case no objects were selected in the first place

                MsgBox("No objects selected")

            End If

            trans.Commit()

        End Using

        Return circLoadList

    End Function

    Public Function DistributeLoads()



        Return 0

    End Function

    ' Creates the Table itself ----------------------------------------------------------------------------------------------------------------------
    <Obsolete>
    Public Sub CreateTable(loadTableAux As LoadTableObj)

        ' Accessess the database -------------------------------------------------------------------------------------------
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim pr As PromptPointResult = ed.GetPoint("Enter Table insertion point: ")  ' Asks the user for the table's insertion point

        If pr.Status = PromptStatus.OK Then

            Dim table As Table = New Table()

            Dim circuits As List(Of CircObj) = loadTableAux.circList

            With table

                Dim headerRows As Integer = 2                                               ' Amount of Rows for the Header
                Dim rows As Integer = circuits.Count + headerRows                           ' Amount of Rows (w/ header included)
                Dim columnTextList() As String = {"Circuito", "Descrição", "Potência", "R", "S", "T"}
                Dim columns As Integer = UBound(columnTextList) - LBound(columnTextList) + 1                                               ' Amount of Columns

                .SetSize(rows, columns)
                .SetRowHeight(12)
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

                For indexRow = headerRows To (rows - 1)

                    For indexColumn = 0 To (columns - 1)

                        .SetTextHeight(indexRow, indexColumn, 8)
                        .SetAlignment(indexRow, indexColumn, CellAlignment.MiddleCenter)

                        Select Case indexColumn
                            Case 0
                                .SetTextString(indexRow, indexColumn, circuits(indexRow - headerRows).circ.ToString())
                            Case 1
                                .SetTextString(indexRow, indexColumn, "Descrição")
                            Case 2
                                .SetTextString(indexRow, indexColumn, circuits(indexRow - headerRows).load.ToString())
                            Case 3
                                .SetTextString(indexRow, indexColumn, circuits(indexRow - headerRows).r.ToString())
                            Case 4
                                .SetTextString(indexRow, indexColumn, circuits(indexRow - headerRows).s.ToString())
                            Case 5
                                .SetTextString(indexRow, indexColumn, circuits(indexRow - headerRows).t.ToString())
                        End Select

                    Next

                Next

                table.GenerateLayout()

            End With

            ' Creates a transaction for the database ---------------------------------------------------------------------------
            Using trans As Transaction = doc.TransactionManager.StartTransaction()

                Dim bt As BlockTable = CType(trans.GetObject(doc.Database.BlockTableId, OpenMode.ForRead), BlockTable)

                Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                btr.AppendEntity(table)

                trans.AddNewlyCreatedDBObject(table, True)

                trans.Commit()

            End Using

        Else



        End If

    End Sub

    ' -----------------------------------------------------------------------------------------------------------------------------------------------

    <Obsolete>
    Public Sub CreateDiagram(loadTableAux As LoadTableObj)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim pr As PromptPointResult = ed.GetPoint("Enter the Diagram insertion point: ")  ' Asks the user for the table's insertion point

        If pr.Status = PromptStatus.OK Then

            ' Get the current database and start a transaction
            'Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
            'acCurDb = Application.DocumentManager.MdiActiveDocument.Database

            Using trans As Transaction = db.TransactionManager.StartTransaction()

                Dim circuits As List(Of CircObj) = loadTableAux.circList

                For index = 0 To (circuits.Count - 1)

                    ' Open the Block table for read
                    Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                    Dim blkID As ObjectId = ObjectId.Null

                    If Not bt.Has("MD - DU Circuito") Then

                        ' Add block from file

                    Else

                        blkID = bt("MD - DU Circuito")

                    End If

                    ' Create and insert the new block reference
                    If blkID <> ObjectId.Null Then


                        '        Dim circuits As List(Of CircObj) = loadTableAux.circList

                        '        'Dim blk
                        '        'Dim blkPath As String = "C:\Users\mathg\OneDrive\Área de Trabalho\AutoCAD Automation Stuff\Blocks\Diagrama Unifilar\MD - DU Circuito.dwg"

                        '        For index = 0 To (circuits.Count - 1)

                        Dim btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)

                        'Dim blkName As String = "MD - DU Circuito"
                        Dim blkPos As Point3d = New Point3d(pr.Value.X, pr.Value.Y + (index * (-60)), pr.Value.Z)

                        Using blkRef As New BlockReference(blkPos, btr.Id)

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
                                                        attRef.TextString = "16 A - C"
                                                        Exit Select
                                                    Case "CIRC_SEÇÃO"
                                                        attRef.TextString = "2,5 mm²"
                                                        Exit Select
                                                    Case "CIRC_POTÊNCIA"
                                                        attRef.TextString = "(" + circuits(index).load.ToString + " W)"
                                                        Exit Select
                                                    Case "CIRC_FASE"
                                                        'attRef.TextString = "R+S+T"
                                                        attRef.TextString = circuits(index).phases
                                                        Exit Select
                                                    Case "CIRC_N_CIRCUITO"
                                                        attRef.TextString = circuits(index).circ.ToString
                                                        Exit Select
                                                    Case "CIRC_DESCRIÇÃO"
                                                        attRef.TextString = "(Description)"
                                                        Exit Select
                                                End Select

                                                blkRef.AttributeCollection.AppendAttribute(attRef)

                                                trans.AddNewlyCreatedDBObject(attRef, True)

                                            End Using

                                        End If

                                    End If

                                Next

                            End If

                        End Using

                    End If

                Next

                ' Save the new object to the database
                trans.Commit()

                ' Dispose of the transaction
            End Using

        End If

    End Sub

End Class

'----------------------------------------------------------------------------------------------------------------------------------------------------

Public Class CircObj

    Public circ As Integer      ' Circuit number
    Public desc As String       ' Description
    Public load As Integer      ' Circuit load
    Public r As Integer         ' Phase R
    Public s As Integer         ' Phase S
    Public t As Integer         ' Phase T
    Public phases As String     ' Phases used by the circuit (e.g R+S)
    Public wire As Integer      ' Wire section
    Public breaker As String    ' Breaker value

    Public connection As String  ' Connection Type (Single/Two/Three-Phase)

    Public Sub New(circAux As Integer, loadAux As Integer, connAux As String)

        circ = circAux
        load = loadAux
        connection = connAux

        desc = ""
        r = 0
        s = 0
        t = 0
        phases = "R"
        wire = 0
        breaker = ""

    End Sub

    Public Sub New()

        circ = 0
        desc = ""
        load = 0
        r = 0
        s = 0
        t = 0
        phases = "R"
        wire = 0
        breaker = ""

        connection = ""

    End Sub

End Class

'----------------------------------------------------------------------------------------------------------------------------------------------------

Public Class LoadTableObj

    Public circList As List(Of CircObj)

    Public Sub New(circListAux As List(Of CircObj))

        circList = circListAux
        DistributeLoads()

    End Sub

    Private Sub DistributeLoads()   ' Distributes the each circuit's load beetween phases R, S and T

        'MsgBox(circList.Count - 1)

        For index = 0 To (circList.Count - 1)   ' Runs through the circuits in the load table

            Dim currentCirc As CircObj = circList(index)

            If index = 0 Then                                       ' If it's the first circuit, don't check the previous circuits

                Select Case currentCirc.connection
                    Case "Single", ""
                        currentCirc.r = currentCirc.load
                        currentCirc.s = 0
                        currentCirc.t = 0
                        currentCirc.phases = "R"
                        GoTo end_of_for_01
                    Case "Two"
                        currentCirc.r = currentCirc.load / 2
                        currentCirc.s = currentCirc.load / 2
                        currentCirc.t = 0
                        currentCirc.phases = "R+S"
                        GoTo end_of_for_01
                    Case "Three"
                        currentCirc.r = currentCirc.load / 3
                        currentCirc.s = currentCirc.load / 3
                        currentCirc.t = currentCirc.load / 3
                        currentCirc.phases = "R+S+T"
                        GoTo end_of_for_01
                End Select

            Else                                                    ' If else checks the previous circuits, unless it's a three phase

                If currentCirc.connection.Contains("Three") Then

                    currentCirc.r = currentCirc.load / 3
                    currentCirc.s = currentCirc.load / 3
                    currentCirc.t = currentCirc.load / 3
                    currentCirc.phases = "R+S+T"
                    GoTo end_of_for_01

                Else

                    For previousIndex = (index - 1) To 0 Step -1

                        Dim previousCirc As CircObj = circList(previousIndex)

                        If previousCirc.r > 0 And previousCirc.s > 0 And previousCirc.t > 0 Then

                            GoTo end_of_for_02

                        Else

                            If previousCirc.r > 0 Then

                                If previousCirc.s > 0 Then

                                    Select Case currentCirc.connection
                                        Case "Single", ""
                                            currentCirc.r = 0
                                            currentCirc.s = 0
                                            currentCirc.t = currentCirc.load
                                            currentCirc.phases = "T"
                                            GoTo end_of_for_01
                                        Case "Two"
                                            currentCirc.r = currentCirc.load / 2
                                            currentCirc.s = 0
                                            currentCirc.t = currentCirc.load / 2
                                            currentCirc.phases = "R+T"
                                            GoTo end_of_for_01
                                    End Select

                                Else

                                    Select Case currentCirc.connection
                                        Case "Single", ""
                                            currentCirc.r = 0
                                            currentCirc.s = currentCirc.load
                                            currentCirc.t = 0
                                            currentCirc.phases = "S"
                                            GoTo end_of_for_01
                                        Case "Two"
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
                                        Case "Single", ""
                                            currentCirc.r = currentCirc.load
                                            currentCirc.s = 0
                                            currentCirc.t = 0
                                            currentCirc.phases = "R"
                                            GoTo end_of_for_01
                                        Case "Two"
                                            currentCirc.r = currentCirc.load / 2
                                            currentCirc.s = currentCirc.load / 2
                                            currentCirc.t = 0
                                            currentCirc.phases = "R+S"
                                            GoTo end_of_for_01
                                    End Select

                                Else

                                    Select Case currentCirc.connection
                                        Case "Single", ""
                                            currentCirc.r = 0
                                            currentCirc.s = 0
                                            currentCirc.t = currentCirc.load
                                            currentCirc.phases = "T"
                                            GoTo end_of_for_01
                                        Case "Two"
                                            currentCirc.r = currentCirc.load / 2
                                            currentCirc.s = 0
                                            currentCirc.t = currentCirc.load / 2
                                            currentCirc.phases = "R+T"
                                            GoTo end_of_for_01
                                    End Select

                                End If

                            Else

                                Select Case currentCirc.connection
                                    Case "Single", ""
                                        currentCirc.r = currentCirc.load
                                        currentCirc.s = 0
                                        currentCirc.t = 0
                                        currentCirc.phases = "R"
                                        GoTo end_of_for_01
                                    Case "Two"
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

End Class
