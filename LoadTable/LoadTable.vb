Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class Conduit

    '<CommandMethod("loadtable")>
    <CommandMethod("zz")>
    Public Sub ConduitConnector()

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
                    Dim attList As New List(Of AttributeReference)()
                    Dim circAttIndex As Short = Nothing
                    Dim loadAttIndex As Short = Nothing
                    Dim circuits As New List(Of Short)()
                    Dim loads As New List(Of Short)()

                    ' Fills the attribute list (attList) with attribute references
                    For Each oid As ObjectId In ac

                        attList.Add(CType(trans.GetObject(oid, OpenMode.ForRead), AttributeReference))
                        MsgBox("Att Tag: " + attList(attList.Count - 1).Tag.ToString + vbNewLine + "Att Value: " + attList(attList.Count - 1).TextString)

                    Next

                    ' Checks if attribute list (from a block) has the tags "CIRCUITO" and "POTÊNCIA"
                    Dim i As Short = 0

                    For Each attr As AttributeReference In attList

                        If attr.Tag.Contains("CIRCUITO") Then

                            circAttIndex = i

                        ElseIf attr.Tag.Contains("POTÊNCIA") Then

                            loadAttIndex = i

                        End If

                        i += 1

                    Next

                    i = 0

                    ' If the blocks has both attributes, adds the load value to an array
                    If circAttIndex.ToString IsNot Nothing And loadAttIndex.ToString IsNot Nothing Then

                        MsgBox("Found Circuit and Load Attributes")

                        Dim alreadyHasCircuit As Boolean = False

                        If circuits IsNot Nothing Then

                            For Each circuit As Short In circuits

                                If attList(circAttIndex).TextString Is circuit.ToString Then

                                    loads(i) = CShort(attList(loadAttIndex).TextString)

                                    alreadyHasCircuit = True

                                    MsgBox("Found Previous Circuit")

                                    Exit For

                                End If

                                i += 1

                            Next

                            If alreadyHasCircuit.ToString Is "False" Then

                                loads(i) = CShort(attList(loadAttIndex).ToString)

                                MsgBox("Added New Circuit")

                            End If

                        Else

                            'MsgBox("3")

                            ''For Each 

                            'Dim attrCirc As AttributeReference = CType(trans.GetObject(attList(circAttIndex), OpenMode.ForRead), AttributeReference)
                            'Dim attrLoad As AttributeReference = CType(trans.GetObject(oIDs(loadAttIndex), OpenMode.ForRead), AttributeReference)

                            'MsgBox("3.1" + (attrCirc.TextString))
                            'MsgBox("3.2" + (attrLoad.TextString))

                            'ReDim circuits(1)
                            'MsgBox("Redimmed circuits")
                            'circuits(0) = CShort(attrCirc.TextString)
                            'MsgBox("Appended circuit")
                            'ReDim loads(1)
                            'MsgBox("Redimmed loads")
                            'loads(0) = CShort(attrLoad.TextString)
                            'MsgBox("Appended load")
                            'alreadyHasCircuit = False

                            'MsgBox("4")

                        End If

                        MsgBox("5")

                    End If

                    i = 0

                    If loads IsNot Nothing Then

                        For Each load As Short In loads

                            MsgBox("6")

                            MsgBox("Load: " + load.ToString)

                        Next

                    End If

                    'For Each aaa As ObjectId In ac

                    '    ' Print the attributes' tag and value
                    '    'attList &= "Attribute " + i.ToString + ": " + attr.Tag + " = " + attr.TextString + vbNewLine
                    '    'i += 1

                    '    Dim attr As AttributeReference = CType(trans.GetObject(oid, OpenMode.ForRead), AttributeReference)
                    '    Dim circuits() As Int16 = Nothing
                    '    Dim loads() As Int16 = Nothing

                    '    If attr.Tag Is "CIRCUITO" Then



                    '        For Each circuit As Int16 In circuits

                    '            If attr.TextString Is circuit.ToString Then

                    '                circuit = Convert.ToInt16(attr.TextString)

                    '            End If

                    '        Next

                    '    End If

                    'Next

                    'MsgBox(attList)

                Next

            Else

                MsgBox("No objects selected")

            End If

            trans.Commit()

        End Using

    End Sub

End Class
