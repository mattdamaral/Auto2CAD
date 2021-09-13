Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class CopperBusbarClass

    Public busbarValue As String
    Public busbarPos As Point3d
    Public totalBreaker As Double

    Public Sub New(totalBreakerAux As Double, busbarPosAux As Point3d)

        Integer.TryParse(totalBreakerAux, totalBreaker)
        busbarPos = busbarPosAux

        If totalBreaker <> Nothing Then

            Select Case (totalBreaker * 1.3)
                Case < 140
                    busbarValue = "15x2"
                    Exit Select
                Case < 170
                    busbarValue = "15x3"
                    Exit Select
                Case < 185
                    busbarValue = "20x2"
                    Exit Select
                Case < 220
                    busbarValue = "20x3"
                    Exit Select
                Case < 270
                    busbarValue = "25x3"
                    Exit Select
                Case < 295
                    busbarValue = "20x5"
                    Exit Select
                Case < 315
                    busbarValue = "30x3"
                    Exit Select
                Case < 350
                    busbarValue = "25x5"
                    Exit Select
                Case < 400
                    busbarValue = "30x5"
                    Exit Select
                Case < 420
                    busbarValue = "40x3"
                    Exit Select
                Case < 520
                    busbarValue = "40x5"
                    Exit Select
                Case < 630
                    busbarValue = "50x5"
                    Exit Select
                Case < 760
                    busbarValue = "40x10"
                    Exit Select
                Case < 820
                    busbarValue = "50x10"
                    Exit Select
                Case < 970
                    busbarValue = "80x5"
                    Exit Select
                Case < 1060
                    busbarValue = "60x10"
                    Exit Select
                Case < 1200
                    busbarValue = "100x5"
                    Exit Select
                Case < 1380
                    busbarValue = "80x10"
                    Exit Select
                Case < 1700
                    busbarValue = "100x10"
                    Exit Select
                Case < 2000
                    busbarValue = "120x10"
                    Exit Select
                Case < 2500
                    busbarValue = "160x10"
                    Exit Select
                Case < 3000
                    busbarValue = "200x10"
                    Exit Select

            End Select

        End If

    End Sub

    ' Draws the busbar (dimensions for the phase/neutral/ground bars)
    Public Sub DrawBusbar()

        Call CommonFunctions.ChangeLayer("MD - Diagrama Unifilar")

        Dim wireColor As Integer = 7 ' White


        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database

        ' Starts a transaction
        Using trans As Transaction = db.TransactionManager.StartTransaction()

            ' Opens the Block table for read
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            Dim blkID As ObjectId = ObjectId.Null ' The ID of the DR block

            ' If there's no block called "MD..." in the block table, get it from the block .dwg, else use it from the block table
            If Not bt.Has("MD - DU Barramento") Then
                ' Should add block from a file path
            Else
                blkID = bt("MD - DU Barramento")
            End If

            ' Creates and inserts the new block reference
            If blkID <> ObjectId.Null Then

                Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)

                    Using blkRef As New BlockReference(busbarPos, btr.Id)

                        Dim curBtr As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                        curBtr.AppendEntity(blkRef)
                        trans.AddNewlyCreatedDBObject(blkRef, True)

                        ' Verify block table record has attribute definitions associated with it
                        If btr.HasAttributeDefinitions Then
                            For Each objID As ObjectId In btr

                                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)

                                If TypeOf dbObj Is AttributeDefinition Then

                                    Dim attDef As AttributeDefinition = dbObj

                                    If Not attDef.Constant Then

                                        Using attRef As New AttributeReference

                                            attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                            attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                            'Checks If the attribute's tag is one of the below
                                            Select Case attRef.Tag
                                                Case "F"
                                                    attRef.TextString = "F - " + busbarValue + " mm"
                                                Case "N"
                                                    attRef.TextString = "N - " + busbarValue + " mm"
                                                Case "PE"
                                                    attRef.TextString = "PE - " + busbarValue + " mm"
                                            End Select

                                            ' Add DU block to the block table record and to the transaction
                                            blkRef.AttributeCollection.AppendAttribute(attRef)
                                            trans.AddNewlyCreatedDBObject(attRef, True)

                                        End Using
                                    End If
                                End If
                            Next
                            'Fills the attribute list (attList) with attribute references
                            'For Each objID As ObjectId In btr

                            '    Dim attR As AttributeReference = CType(trans.GetObject(objID, OpenMode.ForRead), AttributeReference)  'Stores the attribute

                            '    Checks if the attribute's tag is one of the below
                            '    Select Case attR.Tag
                            '        Case "F"
                            '            attR.TextString = busbarValue
                            '        Case "N"
                            '            attR.TextString = busbarValue
                            '        Case "PE"
                            '            attR.TextString = busbarValue
                            '    End Select

                            '     Add DU block to the block table record and to the transaction
                            '    blkRef.AttributeCollection.AppendAttribute(attR)
                            '    trans.AddNewlyCreatedDBObject(attR, True)

                            'Next
                        End If
                    End Using
                End Using
            End If

            trans.Commit()

        End Using
    End Sub

End Class
