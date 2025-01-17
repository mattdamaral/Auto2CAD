﻿Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Imports System.IO

Public Class Commands

    <CommandMethod("QCDesenho")> 'Creates a load table from the drawing, based on the blocks withs attributes
    Public Sub QCDesenho()

        'Gets the info from the drawing
        'Then calls for the drawLT function
        GetAttributesFromDrawing()

    End Sub

    <CommandMethod("zz")>
    Public Sub CreateLoadTable()

        GetList()

    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub GetAttributesFromDrawing()

        Dim circList As New List(Of CircuitClass)()
        Dim hasEnoughData As Boolean = False

        'Accessess the database -------------------------------------------------------------------------------------------
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        'Creates a transaction for the database ---------------------------------------------------------------------------
        Using trans As Transaction = doc.TransactionManager.StartTransaction()

            Dim peo As PromptEntityOptions = New PromptEntityOptions("Selecione a POLYLINE que delimita o perímetro dos circuitos do QD:")

            peo.SetRejectMessage("Objeto selecionado não é do tipo POLYLINE")
            peo.AddAllowedClass(GetType(Polyline), True)

            Dim objSelected = ed.GetEntity(peo)

            'Checks if selection is ok
            If objSelected.Status = PromptStatus.OK Then

                Dim pl As Polyline = CType(trans.GetObject(objSelected.ObjectId, OpenMode.ForRead), Polyline)
                Dim vertices(pl.NumberOfVertices) As Point3d 'Gets the vertices of the perimeter
                Dim pointCollection As Point3dCollection = New Point3dCollection()

                'Passes the vertices of the perimeter to a point3dCollection
                For index = 0 To pl.NumberOfVertices - 1

                    pointCollection.Add(pl.GetPoint3dAt(index))

                Next

                Dim tv As TypedValue() = New TypedValue(0) {}
                tv.SetValue(New TypedValue(CInt(DxfCode.Start), "INSERT"), 0)
                Dim filter As SelectionFilter = New SelectionFilter(tv)

                Dim ss As SelectionSet = ed.SelectCrossingPolygon(pointCollection, filter).Value 'Creates a selection using the perimeter as reference

                If ss IsNot Nothing Then

                    'Tries to get the info of every selected object
                    For Each so As SelectedObject In ss

                        'Blocks' parameters
                        Dim br As BlockReference = CType(trans.GetObject(so.ObjectId, OpenMode.ForRead), BlockReference)
                        Dim ac As AttributeCollection = br.AttributeCollection

                        'Attribute related parameters
                        Dim circAux As Integer = Nothing
                        Dim loadAux As Integer = Nothing
                        Dim connAux As String = ""
                        Dim wireAux As Double = Nothing

                        'Fills the attribute list (attList) with attribute references
                        For Each objID As ObjectId In ac

                            Dim attR As AttributeReference = CType(trans.GetObject(objID, OpenMode.ForRead), AttributeReference)  'Stores the attribute

                            'Checks if the attribute's tag is one of the below
                            Select Case attR.Tag
                                Case "CIRCUITO"
                                    Integer.TryParse(attR.TextString, circAux) 'If the attribute's value is an intenger, save it to circAux
                                Case "POTÊNCIA"
                                    Integer.TryParse(attR.TextString, loadAux) 'If the attribute's value is an intenger, save it to loadAux
                                Case "CONEXÃO"
                                    connAux = attR.TextString.ToString 'Saves the type of connection (mono/bi/tri)
                                Case "SEÇÃO"
                                    Double.TryParse(attR.TextString, wireAux) 'If the attribute's value is a double, save it to wireAux
                            End Select

                        Next

                        'Checks if circAux and loadAux were assigned a value
                        If circAux > 0 And loadAux > 0 Then

                            hasEnoughData = True

                            If circList.Count > 0 Then 'If circList is not empty, either adds the load to an existing circuit or create a new one

                                Dim foundSimiliarCirc = False

                                For index = 0 To (circList.Count - 1) 'Checks if there's a circuit in circList with the same value as the attribute's value

                                    If circAux = circList(index).circ Then

                                        foundSimiliarCirc = True
                                        circList(index).load += loadAux

                                        Select Case wireAux
                                            Case 2.5, 4.0, 6.0

                                                If circList(index).wire < wireAux Then
                                                    circList(index).wire = wireAux
                                                End If

                                                Exit Select
                                            Case Else
                                                wireAux = 2.5
                                        End Select

                                        GoTo goto_01

                                    End If

                                Next

                                If foundSimiliarCirc = False Then      'If there's no circAux in circList, adds a new one to the end of the list

                                    Select Case wireAux
                                        Case 2.5, 4.0, 6.0
                                            Exit Select
                                        Case Else
                                            wireAux = 2.5
                                    End Select

                                    circList.Add(New CircuitClass(circAux.ToString, loadAux, connAux, wireAux))

                                End If

                            Else                                       'If circList is empty, adds a new circuit and assigns the load to it

                                Select Case wireAux
                                    Case 2.5, 4.0, 6.0
                                        Exit Select
                                    Case Else
                                        wireAux = 2.5
                                End Select

                                circList.Add(New CircuitClass(circAux.ToString, loadAux, connAux, wireAux))

                            End If

goto_01:

                        End If

                    Next

                End If

            Else    'Displays this message in case no objects were selected in the first place

                MsgBox("No objects selected")

            End If

            trans.Commit()

        End Using

        If hasEnoughData = True Then

            circList.Sort(Function(x, y) x.circ.CompareTo(y.circ))
            Dim loadTableAux As LoadTableClass = New LoadTableClass(circList)
            'Call LoadTableClass.CreateTable(loadTableAux)
            loadTableAux.CreateTable()
            'Call LoadTableClass.CreateDiagram(loadTableAux)
            loadTableAux.CreateDiagram()

        End If

    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------

    ' Gets the circuits and loads -------------------------------------------------------------------------------------------------------------------
    Public Sub GetList()

        'Dim circList As New List(Of Integer)()
        'Dim loadList As New List(Of Integer)()
        Dim circLoadList As New List(Of CircuitClass)()
        Dim hasEnoughData As Boolean = False

        ' Accessess the database -------------------------------------------------------------------------------------------
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Creates a transaction for the database ---------------------------------------------------------------------------
        Using trans As Transaction = doc.TransactionManager.StartTransaction()

            Dim peo As PromptEntityOptions = New PromptEntityOptions("Select the Polyline: ")

            peo.SetRejectMessage("Objeto selecionado não é do tipo POLYLINE")
            peo.AddAllowedClass(GetType(Polyline), True)

            Dim objSelected = ed.GetEntity(peo)

            ' Checks if selection is ok
            If objSelected.Status = PromptStatus.OK Then

                Dim pl As Polyline = CType(trans.GetObject(objSelected.ObjectId, OpenMode.ForRead), Polyline)
                Dim vertices(pl.NumberOfVertices) As Point3d
                Dim pointCollection As Point3dCollection = New Point3dCollection()

                For index = 0 To pl.NumberOfVertices - 1

                    pointCollection.Add(pl.GetPoint3dAt(index))

                Next

                Dim tv As TypedValue() = New TypedValue(0) {}
                tv.SetValue(New TypedValue(CInt(DxfCode.Start), "INSERT"), 0)
                Dim filter As SelectionFilter = New SelectionFilter(tv)

                Dim ss As SelectionSet = ed.SelectCrossingPolygon(pointCollection, filter).Value

                If ss IsNot Nothing Then

                    ' Repeats code for every item selected
                    For Each so As SelectedObject In ss

                        ' Blocks parameters
                        Dim br As BlockReference = CType(trans.GetObject(so.ObjectId, OpenMode.ForRead), BlockReference)
                        Dim ac As AttributeCollection = br.AttributeCollection

                        ' Attribute related parameters
                        Dim circAux As Integer = Nothing
                        Dim loadAux As Integer = Nothing
                        Dim connAux As String = ""
                        Dim wireAux As Double = Nothing

                        ' Fills the attribute list (attList) with attribute references
                        For Each oid As ObjectId In ac

                            Dim attR As AttributeReference = CType(trans.GetObject(oid, OpenMode.ForRead), AttributeReference)  ' Stores the attribute

                            ' Checks if the attribute's tag is either "CIRCUITO" or "POTÊNCIA"
                            If attR.Tag.Contains("CIRCUITO") Then

                                Integer.TryParse(attR.TextString, circAux)      ' If the attribute's value is an intenger, save it to circAux

                            ElseIf attR.Tag.Contains("POTÊNCIA") Then

                                Integer.TryParse(attR.TextString, loadAux)      ' If the attribute's value is an intenger, save it to loadAux

                            ElseIf attR.Tag.Contains("CONEXÃO") Then

                                connAux = attR.TextString.ToString                     ' Save the type of connection (mono/bi/tri)

                            ElseIf attR.Tag.Contains("SEÇÃO") Then

                                Double.TryParse(attR.TextString, wireAux)      ' If the attribute's value is a double, save it to wireAux

                            End If

                        Next

                        ' Checks if circAux and loadAux were assigned a value
                        If circAux > 0 And loadAux > 0 Then

                            hasEnoughData = True

                            If circLoadList.Count > 0 Then                  ' If circuits list is not empty, either adds the load to an existing circuit or create a new one

                                Dim foundSimiliarCirc = False

                                For index = 0 To (circLoadList.Count - 1)   ' Checks if there's a circuit in circList with the same value as the attribute's value

                                    If circAux = circLoadList(index).circ Then

                                        foundSimiliarCirc = True
                                        circLoadList(index).load += loadAux

                                        Select Case wireAux
                                            Case 2.5, 4.0, 6.0

                                                If circLoadList(index).wire < wireAux Then
                                                    circLoadList(index).wire = wireAux
                                                End If

                                                Exit Select
                                            Case Else
                                                wireAux = 2.5
                                        End Select

                                        GoTo goto_01

                                    End If

                                Next

                                If foundSimiliarCirc = False Then      ' If there's no circAux in circList, adds a new one to the end of the list

                                    Select Case wireAux
                                        Case 2.5, 4.0, 6.0
                                            Exit Select
                                        Case Else
                                            wireAux = 2.5
                                    End Select

                                    circLoadList.Add(New CircuitClass(circAux.ToString, loadAux, connAux, wireAux))

                                End If

                            Else                                       ' If circList is empty, adds a new circuit and assigns the load to it

                                Select Case wireAux
                                    Case 2.5, 4.0, 6.0
                                        Exit Select
                                    Case Else
                                        wireAux = 2.5
                                End Select

                                circLoadList.Add(New CircuitClass(circAux.ToString, loadAux, connAux, wireAux))

                            End If

goto_01:

                        End If

                    Next

                End If

            Else    ' Displays this message in case no objects were selected in the first place

                MsgBox("No objects selected")

            End If

            trans.Commit()

        End Using

        If hasEnoughData = True Then

            circLoadList.Sort(Function(x, y) x.circ.CompareTo(y.circ)) 'To sort out the circuits in the list based on their numbers

            'Displays the circuits data
            Dim circuitsDataHeader As String = "Circuito / Descrição / Pot. Total / Fases / R / S / T / Seção / Disjuntor"
            Dim circuitsDataList As New List(Of String)
            circuitsDataList.Add(circuitsDataHeader)
            'MsgBox("Circuito / Descrição / Pot. Total / Fases / R / S / T / Seção / Disjuntor")
            For index = 0 To circLoadList.Count - 1
                Dim currentCircuitData As String = (circLoadList(index).circ.ToString + " / " +
                    circLoadList(index).desc.ToString + " / " +
                    circLoadList(index).load.ToString + " / " +
                    circLoadList(index).phases.ToString + " / " +
                    circLoadList(index).r.ToString + " / " +
                    circLoadList(index).s.ToString + " / " +
                    circLoadList(index).t.ToString + " / " +
                    circLoadList(index).wire.ToString + " / " +
                    circLoadList(index).breaker.ToString)
                circuitsDataList.Add(currentCircuitData)
                'MsgBox(circLoadList(index).circ.ToString + " / " +
                '    circLoadList(index).desc.ToString + " / " +
                '    circLoadList(index).load.ToString + " / " +
                '    circLoadList(index).phases.ToString + " / " +
                '    circLoadList(index).r.ToString + " / " +
                '    circLoadList(index).s.ToString + " / " +
                '    circLoadList(index).t.ToString + " / " +
                '    circLoadList(index).wire.ToString + " / " +
                '    circLoadList(index).breaker.ToString)
            Next
            Dim circuitData As String
            circuitData = String.Join(vbNewLine, circuitsDataList.ToArray())
            'For index = 0 To circuitsDataList.Count - 1
            '    circuitData = String.Concat(circuitsDataList(index))
            'Next
            MsgBox(circuitData)

            Dim loadTableAux As LoadTableClass = New LoadTableClass(circLoadList)
            'Call LoadTableClass.CreateTable(loadTableAux)
            loadTableAux.CreateTable()
            'Call LoadTableClass.CreateDiagram(loadTableAux)
            loadTableAux.CreateDiagram()

        End If

    End Sub

    ' -----------------------------------------------------------------------------------------------------------------------------------------------

    <CommandMethod("readcsv")>
    Public Sub CSVReader()

        Dim csvPath As String = "C:\Users\mathg\OneDrive\Área de Trabalho\AutoCAD Automation Stuff\Planilhas\Projetos Elétricos - CSV2.csv"
        Dim csvContent As String = File.ReadAllText(csvPath)
        'Console.WriteLine(csvContent)
        'Console.Read()
        MsgBox(csvContent)

        ' Accessess the database -------------------------------------------------------------------------------------------
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim circList As New List(Of CircuitClass)()
        Dim totalDesc As String
        Dim totalLoad As Double
        Dim totalPhases As String
        Dim totalR As Double
        Dim totalS As Double
        Dim totalT As Double
        Dim totalWire As Double
        Dim totalBreaker As String

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            ' Opens the Block Table for reading
            Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            ' Opens the Block Table Record model space for writing
            Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

            Dim inputRecord As String = Nothing
            Dim inReader As StreamReader = File.OpenText(csvPath)
            inputRecord = inReader.ReadLine
            Dim lastLine As Integer = inputRecord.Length             ' Amount of lines in the .csv file
            Dim currentLine As Integer = 0


            ' Goes through every line of the .csv file
            While inputRecord IsNot Nothing
                Dim circ As String = "-"       ' Circuit number
                Dim desc As String = "-"       ' Description
                Dim load As Double = 0       ' Circuit load
                Dim r As Double = 0           ' Phase R
                Dim s As Double = 0           ' Phase S
                Dim t As Double = 0           ' Phase T
                Dim phases As String = "-"     ' Phases used by the circuit (e.g R+S)
                Dim wire As Double = 0        ' Wire section
                Dim breaker As String = "-"    ' Breaker value
                Dim dr As String = "-"    ' Breaker value

                ' Checks if the line contains ","
                If inputRecord.Contains(",") Then
                    currentLine += 1
                    ' Splits the line at every "," it finds
                    Dim circuitData(12) As String
                    circuitData = inputRecord.Split(CChar(","))

                    ' Checks if it's the Total line
                    If circuitData(0).Contains("Total") Then
                        For index = 0 To (circuitData.Length - 1)
                            If circuitData(index) IsNot Nothing Then
                                Select Case index
                                    Case 4
                                        Double.TryParse(circuitData(index), totalLoad)
                                        Exit Select
                                    Case 5
                                        totalPhases = circuitData(index)
                                        Exit Select
                                    Case 6
                                        Double.TryParse(circuitData(index), totalR)
                                        Exit Select
                                    Case 7
                                        Double.TryParse(circuitData(index), totalS)
                                        Exit Select
                                    Case 8
                                        Double.TryParse(circuitData(index), totalT)
                                        Exit Select
                                    Case 9
                                        Double.TryParse(circuitData(index), totalWire)
                                        Exit Select
                                    Case 10
                                        totalBreaker = circuitData(index)
                                        Exit Select

                                End Select
                            End If
                        Next
                        'Checks if it's the first line (table title)
                    ElseIf currentLine = 1 Then
                        totalDesc = circuitData(0)

                        ' Checks if it's a circuit line
                    ElseIf currentLine <> 1 And currentLine <> 2 Then
                        For index = 0 To circuitData.Length - 1
                            'MsgBox(circuitData(index).ToString)
                            If circuitData(index) IsNot Nothing Then
                                Select Case index
                                    Case 0
                                        circ = circuitData(index)
                                        Exit Select
                                    Case 1
                                        desc = circuitData(index)
                                        Exit Select
                                    Case 4
                                        Double.TryParse(circuitData(index), load)
                                        Exit Select
                                    Case 5
                                        phases = circuitData(index)
                                        Exit Select
                                    Case 6
                                        Double.TryParse(circuitData(index), r)
                                        Exit Select
                                    Case 7
                                        Double.TryParse(circuitData(index), s)
                                        Exit Select
                                    Case 8
                                        Double.TryParse(circuitData(index), t)
                                        Exit Select
                                    Case 9
                                        Double.TryParse(circuitData(index), wire)
                                        Exit Select
                                    Case 10
                                        breaker = circuitData(index)
                                        Exit Select
                                    Case 11
                                        dr = circuitData(index)
                                        'MsgBox(circuitData(index).ToString)
                                        Exit Select

                                End Select
                            End If
                        Next
                        circList.Add(New CircuitClass(circ, desc, load, phases, r, s, t, wire, breaker, dr))

                    End If
                End If
                ' Reads the next line
                inputRecord = inReader.ReadLine()

            End While
        End Using
        Dim loadTable As LoadTableClass = New LoadTableClass(circList, totalDesc, totalLoad, totalPhases, totalR, totalS, totalT, totalWire, totalBreaker)
        'Call LoadTableClass.CreateDiagram(loadTable)
        loadTable.CreateDiagram()
    End Sub
End Class

