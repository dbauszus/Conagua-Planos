Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Conagua_Elements

    Public Sub Conagua_OverlapTxt()
        Try
            EmptyList("lOverlapTxt")
            Dim x, x1, x2, y, y1, y2, z As Double
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=808")
            Dim lResponse As Integer
            Do
                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ESCAPE
                        Exit Do

                    Case SIS_ARG_POSITION
                        Try
                            Loader.SIS.SplitPos(x, y, z, (Loader.SIS.Snap2D(x, y, 5, True, "L", "fProperty", "")))
                            Dim overlapName = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "planoID$")
                            Dim planoNUMtotal = Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "planoNUMtotal&")
                            Dim a = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetGeomLengthUpto(0, 0, x, y, z))
                            Dim aDeg = Math.Round(a * (180 / Math.PI))
                            If aDeg = 180 Or aDeg = 90 Or aDeg = -90 Then aDeg += 180
                            If SetDatasetOverlay() = False Then Exit Sub

                            'draw flecha point
                            Loader.SIS.CreatePoint(x, y, 0, "CONAGUA.FLECHALIGA", a - 1.57079, 1)
                            Loader.SIS.DeselectAll()
                            Loader.SIS.SelectItem()
                            Loader.SIS.DoCommand("AComExplodeShape")
                            Loader.SIS.OpenSel(0)
                            Loader.SIS.AddToList("lOverlapTxt")

                            'draw text
                            Loader.SIS.CreateText(x + (12.5 * Math.Cos(a + 1.57079)), y + (12.5 * Math.Sin(a + 1.57079)), 0, String.Format("Liga con el plano {0} de {1:00}", overlapName, planoNUMtotal))
                            Loader.SIS.UpdateItem()
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CARTA")
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 946)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_RIGHT)
                            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", aDeg)
                            Loader.SIS.UpdateItem()
                            Loader.SIS.DeselectAll()
                            Loader.SIS.SelectItem()
                            Loader.SIS.DoCommand("AComTextToBox")
                            Loader.SIS.OpenSel(0)
                            Loader.SIS.AddToList("lOverlapTxt")

                            'draw flecha line
                            Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetExtent())
                            Select Case aDeg

                                Case 0
                                    Loader.SIS.MoveTo(x1, y1, 0)
                                    Loader.SIS.LineTo(x2, y1, 0)
                                    Loader.SIS.LineTo(x, y, 0)

                                Case 90
                                    Loader.SIS.MoveTo(x1, y1, 0)
                                    Loader.SIS.LineTo(x1, y2, 0)
                                    Loader.SIS.LineTo(x, y, 0)

                                Case 270
                                    Loader.SIS.MoveTo(x2, y2, 0)
                                    Loader.SIS.LineTo(x2, y1, 0)
                                    Loader.SIS.LineTo(x, y, 0)

                                Case 360
                                    Loader.SIS.MoveTo(x1, y2, 0)
                                    Loader.SIS.LineTo(x2, y2, 0)
                                    Loader.SIS.LineTo(x, y, 0)

                            End Select

                            Loader.SIS.UpdateItem()
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CARTA")
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 946)
                            Loader.SIS.AddToList("lOverlapTxt")
                            Loader.SIS.CloseItem()
                            Loader.SIS.DeselectAll()
                            Loader.SIS.SelectList("lOverlapTxt")
                            Exit Do
                        Catch ex As Exception
                            MsgBox("No L Snap", MsgBoxStyle.Exclamation, "")
                            Exit Do
                        End Try
                End Select
            Loop
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub Conagua_CurvaTxt()
        Try
            Dim x, y, z As Double
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=120")
            Loader.SIS.CreatePropertyFilter("fAdd", "_FC&=121")
            Loader.SIS.CombineFilter("fProperty", "fProperty", "fAdd", SIS_BOOLEAN_OR)
            Loader.SIS.CreatePropertyFilter("fAdd", "_FC&=130")
            Loader.SIS.CombineFilter("fProperty", "fProperty", "fAdd", SIS_BOOLEAN_OR)
            Loader.SIS.CreatePropertyFilter("fAdd", "_FC&=131")
            Loader.SIS.CombineFilter("fProperty", "fProperty", "fAdd", SIS_BOOLEAN_OR)
            Dim lResponse As Integer
            Do
                lResponse = Loader.SIS.GetPosEx(x, y, z)
                Select Case lResponse

                    Case SIS_ARG_ESCAPE
                        Exit Do

                    Case SIS_ARG_POSITION
                        Try
                            Loader.SIS.SplitPos(x, y, z, (Loader.SIS.Snap2D(x, y, 5, True, "L", "fProperty", "")))
                            Dim lineangle = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetGeomLengthUpto(0, 0, x, y, z)) * 180 / Math.PI
                            If (lineangle > 90 And lineangle < 270) Or (lineangle < -90 And lineangle > -270) Then lineangle += 180
                            If SetDatasetOverlay() = False Then Exit Sub
                            Loader.SIS.CreateText(x, y, 0, String.Format("{0:#.0}", Loader.SIS.Evaluate(SIS_OT_CURITEM, 0, "round(_oz#,1)")))
                            Loader.SIS.UpdateItem()
                            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_ALTIMETRIA")
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 142)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", lineangle)
                            Loader.SIS.UpdateItem()
                        Catch ex As Exception
                            MsgBox("No L Snap", MsgBoxStyle.Exclamation, "")
                        End Try

                End Select
            Loop
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub Conagua_FlechaDestino()
        Try
            Loader.SIS.OpenSel(0)
            If SetDatasetOverlay() = False Then Exit Sub
            Loader.SIS.DeselectAll()
            Loader.APP.AddTrigger("AComLine::End", New SisTriggerHandler(AddressOf DrawLine_End))
            Loader.APP.AddTrigger("AComLine::Snap", New SisTriggerHandler(AddressOf DrawLine_Snap))
            Loader.SIS.DoCommand("AComLine")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub DrawLine_End(ByVal sender As Object, ByVal e As SisTriggerArgs)
        Loader.APP.RemoveTrigger("AComLine::End", AddressOf DrawLine_End)
        Loader.APP.RemoveTrigger("AComLine::Snap", AddressOf DrawLine_Snap)
    End Sub

    Private DrawLine_1stSnap = False
    Private Sub DrawLine_Snap(ByVal sender As Object, ByVal e As SisTriggerArgs)
        Try
            If DrawLine_1stSnap = False Then
                DrawLine_1stSnap = True
            Else
                Loader.APP.RemoveTrigger("AComLine::End", AddressOf DrawLine_End)
                Loader.APP.RemoveTrigger("AComLine::Snap", AddressOf DrawLine_Snap)
                SendKeys.SendWait("{ENTER}")
                SendKeys.SendWait("{ESC}")
                Loader.SIS.OpenSel(0)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_PLANIMETRIA")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 225)
                Loader.SIS.UpdateItem()
                Dim x, y, z, a As Double
                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(0, 1))
                a = Loader.SIS.GetGeomAngleFromLength(0, 0.1)
                Loader.SIS.CreatePoint(x, y, 0, "CONAGUA.FLECHA", a, 1)
                Loader.SIS.AddToList("lExplode")
                Loader.SIS.ExplodeShape("lExplode")
                Loader.SIS.EmptyList("lExplode")
                Loader.SIS.OpenSel(0)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_PLANIMETRIA")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 225)
                Loader.SIS.CloseItem()
                Loader.SIS.DeselectAll()
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub Conagua_FlechaTxt(ByVal featureTable As String, ByVal FC As Integer)
        Try
            Loader.SIS.CreateListFromSelection("lTexts")
            Loader.SIS.OpenSel(0)
            If SetDatasetOverlay() = False Then Exit Sub
            Dim x1, x2, x3, y1, y2, y3, yO, z, a As Double
            Loader.SIS.DeselectAll()
            Loader.SIS.SelectList("lTexts")
            Loader.SIS.DoCommand("AComTextToBox")
            Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lTexts"))
            Loader.SIS.DoCommand("AComBoxToText")
            If Loader.SIS.GetListSize("lTexts") = 1 Then
                yO = y1
            Else
                yO = (y1 + y2) / 2
            End If
            Dim lResponse As Integer
            Do
                lResponse = Loader.SIS.GetPosEx(x3, y3, z)
                Select Case lResponse

                    Case SIS_ARG_ESCAPE
                        Exit Do

                    Case SIS_ARG_POSITION
                        If (x3 - x1) < 0 Then
                            Loader.SIS.MoveTo(x2, yO, 0)
                            Loader.SIS.LineTo(x1, yO, 0)
                        Else
                            Loader.SIS.MoveTo(x1, yO, 0)
                            Loader.SIS.LineTo(x2, yO, 0)
                        End If
                        Loader.SIS.LineTo(x3, y3, 0)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", featureTable)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", FC)
                        Loader.SIS.UpdateItem()
                        a = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetGeomLength(0) - 0.001)
                        Loader.SIS.SplitPos(x1, y1, z, Loader.SIS.GetGeomPt(0, Loader.SIS.GetGeomNumPt(0) - 1))
                        Loader.SIS.CreatePoint(x1, y1, 0, "CONAGUA.FLECHAPOINT", a, 1)
                        Loader.SIS.AddToList("lExplode")
                        Loader.SIS.ExplodeShape("lExplode")
                        Loader.SIS.EmptyList("lExplode")
                        Loader.SIS.OpenSel(0)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", featureTable)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", FC)
                        Loader.SIS.CloseItem()
                        Loader.SIS.DeselectAll()
                        Exit Do
                End Select
            Loop
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub Conagua_Control()
        Try
            Loader.SIS.CreateListFromSelection("lControl")
            For i = 0 To Loader.SIS.GetListSize("lControl") - 1
                Loader.SIS.OpenList("lControl", i)
                If SetDatasetOverlay() = True Then
                    Dim scaleFactor As Double = Loader.SIS.GetFlt(SIS_OT_DATASET, Loader.SIS.GetDataset(), "_scale#") / 1000
                    Loader.SIS.CreateText(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"), Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#") + 2.5 * scaleFactor, 0, Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "control_id$"))
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CONTROL")
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 330)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                    Loader.SIS.CloseItem()
                    Loader.SIS.OpenList("lControl", i)
                    Loader.SIS.CreateText(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"), Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#") - 2.5 * scaleFactor, 0, Loader.SIS.Evaluate(SIS_OT_CURITEM, 0, "'Z=' + str(round(_oz#,3))"))
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CONTROL")
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 330)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                    Loader.SIS.CloseItem()
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub Conagua_ControlFlecha()
        Try
            Loader.SIS.CreateListFromSelection("lControl")
            Dim x0, x1, x2, x3, y0, y1, y2, y3, z As Double
            Dim lResponse As Integer
            Do
                lResponse = Loader.SIS.GetPosEx(x0, y0, z)
                Select Case lResponse

                    Case SIS_ARG_ESCAPE
                        Exit Do

                    Case SIS_ARG_POSITION
                        Loader.SIS.DoCommand("AComTextToBox")
                        Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lControl"))
                        x3 = (x1 + x2) / 2
                        y3 = (y1 + y2) / 2
                        Loader.SIS.MoveList("lControl", x0 - x3, y0 - y3, 0, 0, 1)
                        Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lControl"))
                        Loader.SIS.DoCommand("AComBoxToText")
                        Loader.SIS.OpenList("lControl", 0)
                        Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#") + ((y0 - Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")) / 3))
                        Loader.SIS.CloseItem()
                        Loader.SIS.OpenList("lControl", 1)
                        Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#") + ((y0 - Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")) / 3))
                        Loader.SIS.CloseItem()
                        If x0 > x3 Then
                            Loader.SIS.MoveTo(x2, y0, 0)
                            Loader.SIS.LineTo(x1, y0, 0)
                        Else
                            Loader.SIS.MoveTo(x1, y0, 0)
                            Loader.SIS.LineTo(x2, y0, 0)
                        End If
                        Loader.SIS.LineTo(x3, y3, 0)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CONTROL")
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 330)
                        Loader.SIS.CloseItem()
                        Exit Do
                End Select
            Loop
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub Conagua_CotaTxt()
        Try
            Loader.SIS.CreateListFromSelection("lCota")
            For i = 0 To Loader.SIS.GetListSize("lCota") - 1
                Loader.SIS.OpenList("lCota", i)
                If SetDatasetOverlay() = False Then Exit Sub
                Dim scaleFactor = Loader.SIS.GetFlt(SIS_OT_DATASET, Loader.SIS.GetDataset, "_scale#") / 1000
                If Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "_FC&") = 101 Then
                    'bati
                    Loader.SIS.CreateText(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"), Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#") + 2.5 * scaleFactor, 0, Loader.SIS.Evaluate(SIS_OT_CURITEM, 0, "round(_oz#,2)"))
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_ALTIMETRIA")
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 140)
                Else
                    'foto
                    Loader.SIS.CreateText(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"), Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#"), 0, Loader.SIS.Evaluate(SIS_OT_CURITEM, 0, "round(_oz#,2)"))
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_ALTIMETRIA")
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 141)
                End If
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                Loader.SIS.CloseItem()
            Next
            Loader.SIS.DeselectAll()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    'Public Sub Conagua_CotaBati()
    '    Try
    '        Dim x, y, z As Double
    '        Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=311")
    '        Dim lResponse As Integer
    '        Do
    '            lResponse = Loader.SIS.GetPosEx(x, y, z)
    '            Select Case lResponse
    '                Case SIS_ARG_ESCAPE
    '                    Exit Do
    '                Case SIS_ARG_POSITION
    '                    Try
    '                        Loader.SIS.SplitPos(x, y, z, (Loader.SIS.Snap2D(x, y, 5, True, "P", "fProperty", "")))
    '                        Loader.SIS.CreateText(x, y, 0, Loader.SIS.Evaluate(SIS_OT_CURITEM, 0, "round(_oz#,2)"))
    '                        Loader.SIS.UpdateItem()
    '                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA")
    '                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 313)
    '                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
    '                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
    '                        Loader.SIS.UpdateItem()
    '                        Loader.SIS.SelectItem()
    '                        Loader.SIS.DoCommand("AComTextToBox")
    '                    Catch ex As Exception
    '                        MsgBox("No P Snap", MsgBoxStyle.Exclamation, "")
    '                    End Try
    '            End Select
    '        Loop
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '    End Try
    'End Sub

    'Public Sub Conagua_KMTxt()
    '    Try
    '        Dim x, y, z As Double
    '        Dim txt As String
    '        Dim alter As Integer = 0
    '        Loader.SIS.OpenSel(0)
    '        If SetDatasetOverlay() = False Then Exit Sub
    '        If Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "_class$") = "Line" Then
    '            Loader.SIS.CreateListFromSelection("list")
    '            For i = 0 To Loader.SIS.GetGeomLength(0) Step 500
    '                Loader.SIS.OpenList("list", 0)
    '                Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, i))
    '                Loader.SIS.CreatePoint(x, y, 0, "", 0, 1)
    '                If alter = 0 Then
    '                    txt = "km" & Str(Int(i / 1000)) & " +000"
    '                    alter = 1
    '                Else
    '                    txt = "km" & Str(Int(i / 1000)) & " +500"
    '                    alter = 0
    '                End If
    '                Loader.SIS.CreateText(x, y, 0, txt)
    '                Loader.SIS.UpdateItem()
    '                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA")
    '                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 608)
    '                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
    '                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
    '                Loader.SIS.UpdateItem()
    '                Loader.SIS.SelectItem()
    '                Loader.SIS.DoCommand("AComTextToBox")
    '            Next
    '        End If
    '        Loader.SIS.DeselectAll()
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '    End Try
    'End Sub

    'Public Sub Conagua_Flecha()
    '    Try
    '        Dim x, y, z, l, a As Double
    '        Loader.SIS.OpenSel(0)
    '        If Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "_class$") = "Line" Then
    '            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA")
    '            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 980)
    '            l = Loader.SIS.GetGeomLength(0)
    '            a = Loader.SIS.GetGeomAngleFromLength(0, l - 0.001)
    '            Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPt(0, Loader.SIS.GetGeomNumPt(0) - 1))
    '            Loader.SIS.CreatePoint(x, y, 0, "FLECHA", a, 1)
    '            Loader.SIS.UpdateItem()
    '            Loader.SIS.DeselectAll()
    '            Loader.SIS.SelectItem()
    '            Loader.SIS.DoCommand("AComExplodeShape")
    '            Loader.SIS.DeselectAll()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '    End Try
    'End Sub

    Private Function SetDatasetOverlay()
        Try
            For i = 0 To Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1
                If Loader.SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&") = Loader.SIS.GetDataset() Then
                    Loader.SIS.SetInt(SIS_OT_OVERLAY, i, "_status&", 3)
                    Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", i)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"))
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox("Layer locked")
            Return False
        End Try
    End Function

    Private Sub EmptyList(ByVal List As String)
        Try
            Loader.SIS.EmptyList(List)
        Catch
        End Try
    End Sub

End Class
