Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class DRClass

    'Inherits LoadTableClass

    Public DRValue As String ' The value of the DR
    Public DRType As String ' Either DDR or IDR
    ' Implement a function in loadtable that can get the type of DR through the header
    Public DRPos As Point3d ' Where the DR should be drawn

    'Public Sub New()

    'End Sub

    Public Sub New(circListAux As List(Of CircuitClass), DRValueAux As String, DRTypeAux As String, DRPosAux As Point3d)

        'MyBase.New(circListAux)

        DRValue = DRValueAux
        DRType = DRTypeAux
        DRPos = DRPosAux

    End Sub

    ' Draws the DR using a block from autocad
    Public Sub DrawDR()

        Call CommonFunctions.ChangeLayer("MD - Diagrama Unifilar")

        Dim wireColor As Integer = 7 ' White

        If DRValue.Contains("-") Then
            Call CommonFunctions.DrawLine(DRPos, New Point3d(DRPos.X - 55, DRPos.Y, DRPos.Z), wireColor)
        Else

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database

            ' Starts a transaction
            Using trans As Transaction = db.TransactionManager.StartTransaction()

                ' Opens the Block table for read
                Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                Dim blkID As ObjectId = ObjectId.Null ' The ID of the DR block

                ' If there's no block called "MD..." in the block table, get it from the block .dwg, else use it from the block table
                If Not bt.Has("MD - DU DR") Then
                    ' Should add block from a file path
                Else
                    blkID = bt("MD - DU DR")
                End If

                ' Creates and inserts the new block reference
                If blkID <> ObjectId.Null Then

                    Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)

                        Using blkRef As New BlockReference(New Point3d(DRPos.X - 45, DRPos.Y, DRPos.Z), btr.Id)

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

                                                'If attRef.Tag = "DR_VALUE" Then

                                                ' Draws the lines to the left and to the right of the DR Block (the ones that connect to the buses)
                                                Call CommonFunctions.DrawLine(DRPos, New Point3d(DRPos.X - 10, DRPos.Y, DRPos.Z), wireColor)
                                                Call CommonFunctions.DrawLine(New Point3d(DRPos.X - 45, DRPos.Y, DRPos.Z), New Point3d(DRPos.X - 55, DRPos.Y, DRPos.Z), wireColor)

                                                attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                                attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                                attRef.TextString = DRValue & " A"

                                                ' Add DU block to the block table record and to the transaction
                                                blkRef.AttributeCollection.AppendAttribute(attRef)
                                                trans.AddNewlyCreatedDBObject(attRef, True)

                                                'End If
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
        End If
    End Sub
End Class
