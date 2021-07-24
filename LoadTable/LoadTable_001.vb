Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class LoadTable

    '<CommandMethod("loadtable")>
    <CommandMethod("zz")>
    Public Sub CreateLoadTable()

        Dim CircLoad = GetList()

        ' Displays all the circuits in circList
        For index = 0 To (CircLoad.circList.Count - 1)

            MsgBox("Circuit: " + CircLoad.circList(index).ToString + vbNewLine + "Load: " + CircLoad.loadList(index).ToString)

        Next

        CreateTable(CircLoad.circList, CircLoad.loadList)

    End Sub

    ' Gets the circuits and loads -------------------------------------------------------------------------------------------------------------------
    Public Function GetList() As (circList As List(Of Integer), loadList As List(Of Integer)) ' GetList(0) returns circList, while GetList(1) return loadList

        Dim circList As New List(Of Integer)()
        Dim loadList As New List(Of Integer)()

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

                    ' Fills the attribute list (attList) with attribute references
                    For Each oid As ObjectId In ac

                        Dim attR As AttributeReference = CType(trans.GetObject(oid, OpenMode.ForRead), AttributeReference)  ' Stores the attribute
                        'MsgBox("Att Tag: " + attR.Tag.ToString + vbNewLine + "Att Value: " + attR.TextString)               ' Displays the attribute tag and value

                        Dim attRTag As String = attR.Tag            ' Stores the attribute's tag as string
                        Dim attRValue As String = attR.TextString   ' Stores the attribute's value as string

                        ' Checks if the attribute's tag is either "CIRCUITO" or "POTÊNCIA"
                        If attRTag.Contains("CIRCUITO") Then

                            Integer.TryParse(attRValue, circAux)    ' If the attribute's value is an intenger, save it to circAux

                        ElseIf attRTag.Contains("POTÊNCIA") Then

                            Integer.TryParse(attRValue, loadAux)    ' If the attribute's value is an intenger, save it to loadAux

                        End If

                    Next

                    ' Checks if circAux and loadAux were assigned a value
                    If circAux.ToString IsNot Nothing And loadAux.ToString IsNot Nothing Then

                        'MsgBox("List size: " + circList.Count.ToString + vbNewLine + "Circuit: " + circAux.ToString + vbNewLine + "Load: " + loadAux.ToString)

                        If circList.Count > 0 Then                  ' If circuits list is not empty, either adds the load to an existing circuit or create a new one

                            Dim foundSimiliarCirc = False

                            For index = 0 To (circList.Count - 1)   ' Checks if there's a circuit in circList with the same value as the attribute's value

                                If circAux = circList(index) Then

                                    foundSimiliarCirc = True
                                    loadList(index) += loadAux

                                    GoTo goto_01

                                End If

                            Next

                            If foundSimiliarCirc = False Then      ' If there's no circAux in circList, adds a new one to the end of the list

                                circList.Add(circAux)
                                loadList.Add(loadAux)

                            End If

                        Else                                       ' If circList is empty, adds a new circuit and assigns the load to it

                            circList.Add(circAux)
                            loadList.Add(loadAux)

                        End If

goto_01:

                    End If

                Next

                circList.Sort()

                '' Displays all the circuits in circList
                'For index = 0 To (circList.Count - 1)

                '    MsgBox("Circuit: " + circList(index).ToString + vbNewLine + "Load: " + loadList(index).ToString)

                'Next

            Else    ' Displays this message in case no objects were selected in the first place

                MsgBox("No objects selected")

            End If

            trans.Commit()

        End Using

        Return (circList, loadList)

    End Function

    ' Creates the Table itself ----------------------------------------------------------------------------------------------------------------------
    <Obsolete>
    Public Sub CreateTable(circList As List(Of Integer), loadList As List(Of Integer))

        ' Accessess the database -------------------------------------------------------------------------------------------
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim pr As PromptPointResult = ed.GetPoint("Enter table insertion point: ")  ' Asks the user for the table's insertion point

        If pr.Status = PromptStatus.OK Then

            Dim table As Table = New Table()

            With table

                .SetSize(circList.Count + 1, 2)
                .SetRowHeight(12)
                .SetColumnWidth(100)
                .Position = pr.Value

                .SetTextHeight(0, 0, 8)
                .SetTextString(0, 0, "Circuito")
                .SetAlignment(0, 0, CellAlignment.MiddleCenter)
                .SetTextHeight(0, 1, 8)
                .SetTextString(0, 1, "Potência")
                .SetAlignment(0, 1, CellAlignment.MiddleCenter)

                For index = 0 To (circList.Count - 1)

                    .SetTextHeight(index + 1, 0, 8)
                    .SetTextString(index + 1, 0, circList(index).ToString)
                    .SetAlignment(index + 1, 0, CellAlignment.MiddleCenter)
                    .SetTextHeight(index + 1, 1, 8)
                    .SetTextString(index + 1, 1, loadList(index).ToString)
                    .SetAlignment(index + 1, 1, CellAlignment.MiddleCenter)

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

End Class

Public Class CircAndLoad

    Public circ As Integer
    Public load As Integer

End Class

'----------------------------------------------------------------------------------------------------------------------------------------------------


