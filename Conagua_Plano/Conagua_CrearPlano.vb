Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Conagua_CrearPlano

    Private scaleFactor As Double
    Private seccion_pri As Double
    Private seccion_ult As Double
    Private xPlan As Double
    Private yPlan As Double

    Public Sub CrearPlano()
        Try
            'open marco and create "lMarco"
            Loader.SIS.OpenSel(0)
            Loader.SIS.CreateListFromSelection("lMarco")

            'store Marco properties, create locus
            scaleFactor = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "scaleFactor#")
            xPlan = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#") - (Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#") / 2)
            yPlan = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#") - (Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#") / 2)
            Loader.SIS.CreateLocusFromItem("locusIntersect", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
            Loader.SIS.CreateLocusFromItem("locusCrossby", SIS_GT_CROSSBY, SIS_GM_GEOMETRY)
            Loader.SIS.CreateLocusFromItem("locusContain", SIS_GT_CONTAIN, SIS_GM_GEOMETRY)
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=999")

            lDistribuccion()

            'create plan overlay
            Loader.SIS.CreateInternalOverlay(Loader.SIS.GetListItemStr("lMarco", 0, "planoID$"), 0)
            Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", 0)
            Loader.SIS.SetFlt(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, 0, "_nDataset&"), "_scale#", 1000 * scaleFactor)
            Loader.SIS.SetStr(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, 0, "_nDataset&"), "_featureTable$", "CONAGUA_CARTA")

            lPlanoContent()

            If lTemplate() = False Then Exit Sub

            Bloque(Loader.SIS.GetListItemStr("lMarco", 0, "bloque_registro_sgt$"), 902)
            Bloque(Loader.SIS.GetListItemStr("lMarco", 0, "bloque_registro_oc$"), 901)
            Bloque(Loader.SIS.GetListItemStr("lMarco", 0, "bloque_responsable$"), 903)
            Bloque(Loader.SIS.GetListItemStr("lMarco", 0, "bloque_titulo$"), 904)

            MarcoCoords()

            MarcoDistribuccion()

            MarcoCroquis()

            Simbologia()

            SimbologiaText()

            Nota()

            Referencia()

            SeccionTramo()

            ScaleTxt()

            PlanoTxt("planoID", Loader.SIS.GetListItemStr("lMarco", 0, "planoID$"))
            PlanoTxt("planoIDconNUM", String.Format("{0} de {1:00}", Loader.SIS.GetListItemStr("lMarco", 0, "planoID$"), Loader.SIS.GetListItemInt("lMarco", 0, "planoNUMtotal&")))
            PlanoTxt("fecha", Loader.SIS.GetListItemStr("lMarco", 0, "fecha$"))
            PlanoTxt("entidad", Loader.SIS.GetListItemStr("lMarco", 0, "entidad$"))
            PlanoTxt("localidad", Loader.SIS.GetListItemStr("lMarco", 0, "localidad$"))
            PlanoTxt("municipio", Loader.SIS.GetListItemStr("lMarco", 0, "municipio$"))
            PlanoTxt("proyecto", Loader.SIS.GetListItemStr("lMarco", 0, "proyecto$"))

            'set vertices
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=42")
            If Loader.SIS.ScanList("lPoints", "lPlanoContent", "fProperty", "") > 0 Then
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectList("lPoints")
                Loader.SIS.DoCommand("AComExplodeShape")
                Loader.SIS.CreateListFromSelection("lPoints")
                Loader.SIS.SetListInt("lPoints", "_FC&", 42)
            End If
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=21")
            If Loader.SIS.ScanList("lPoints", "lPlanoContent", "fProperty", "") > 0 Then
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectList("lPoints")
                Loader.SIS.DoCommand("AComExplodeShape")
                Loader.SIS.CreateListFromSelection("lPoints")
                Loader.SIS.SetListInt("lPoints", "_FC&", 21)
            End If

            'decompose points
            Loader.SIS.CreateClassTreeFilter("fPoints", "-Item +Point")
            If Loader.SIS.ScanOverlay("lPoints", 0, "fPoints", "") > 0 Then
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectList("lPoints")
                Loader.SIS.DoCommand("AComExplodeShape")
                Loader.SIS.DeselectAll()
                If Loader.SIS.ScanOverlay("lPoints", 0, "fPoints", "") > 0 Then Loader.SIS.Delete("lPoints")
            End If

            'text2box
            Loader.SIS.CreateClassTreeFilter("fTexts", "-Item +Text")
            If Loader.SIS.ScanOverlay("lTexts", 0, "fTexts", "") > 0 Then
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectList("lTexts")
                Loader.SIS.DoCommand("AComTextToBox")
                Loader.SIS.DeselectAll()
            End If

            'change featureTable
            If Loader.SIS.ScanOverlay("lOverlay", 0, "", "") > 0 Then
                Loader.SIS.SetListStr("lOverlay", "_featureTable$", "CONAGUA_CARTA")
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub lDistribuccion()
        Try
            For i = 0 To Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1
                If Loader.SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&") = Loader.SIS.GetDataset() Then
                    Loader.SIS.CreateListFromOverlay(i, "lDistribuccion")
                    Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=808")
                    Loader.SIS.ScanList("lOverlap", "lDistribuccion", "fProperty", "locusCrossby")
                    Loader.SIS.SetInt(SIS_OT_OVERLAY, i, "_status&", 0)
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub lPlanoContent()
        Try

            'copy and cut map elements
            If Loader.SIS.Scan("lPlanoContent", "V", "", "locusIntersect") > 0 Then
                Try
                    Loader.SIS.CombineLists("lPlanoContent", "lOverlap", "lPlanoContent", SIS_BOOLEAN_OR)
                Catch
                End Try
                Loader.SIS.CopyListItems("lPlanoContent")

                'lAreaZF
                Loader.SIS.CreatePropertyFilter("fProperty", "tramo$='" & Loader.SIS.GetListItemStr("lMarco", 0, "tramo$") & "'")
                Loader.SIS.CreatePropertyFilter("fCombine", "_FC& = 40")
                Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_AND)
                If Loader.SIS.ScanList("lAreaZF", "lPlanoContent", "fProperty", "locusContain") > 0 Then ZonaFederal()

                'snip lPlanoContent
                Loader.SIS.OpenList("lMarco", 0)
                If Loader.SIS.ScanList("lSnip", "lPlanoContent", "", "locusCrossby") > 0 Then Loader.SIS.SnipGeometry("lSnip", False)

                'decompose multilines after cutting lines
                Loader.SIS.CreateClassTreeFilter("fMultiLine", "-Item +MultiLine")
                If Loader.SIS.ScanList("lMultiLine", "lPlanoContent", "fMultiLine", "") > 0 Then
                    Loader.SIS.DeselectAll()
                    Loader.SIS.SelectList("lMultiLine")
                    Loader.SIS.DoCommand("AComDecompose")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub ZonaFederal()
        Try
            Loader.SIS.OpenList("lAreaZF", 0)
            seccion_pri = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "seccion_pri#")
            seccion_ult = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "seccion_ult#")
            For i = 0 To Loader.SIS.GetListSize("lAreaZF") - 1
                Loader.SIS.OpenList("lAreaZF", i)
                If Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "seccion_pri#") < seccion_pri Then seccion_pri = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "seccion_pri#")
                If Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "seccion_ult#") > seccion_ult Then seccion_ult = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "seccion_ult#")
            Next
            'Loader.SIS.CreatePropertyFilter("fProperty", "_FC& = 21")
            'Loader.SIS.CreatePropertyFilter("fCombine", "_FC& = 25")
            'Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_OR)
            'Loader.SIS.CreatePropertyFilter("fCombine", "_FC& = 26")
            'Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_OR)
            'Loader.SIS.CreatePropertyFilter("fCombine", "_FC& = 30")
            'Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_OR)
            'Loader.SIS.CreatePropertyFilter("fCombine", "_FC& = 42")
            'Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_OR)
            'If Loader.SIS.ScanList("lZF", "lPlanoContent", "fProperty", "") > 0 Then
            '    Loader.SIS.CreatePropertyFilter("fProperty", "seccion# <= " & seccion_ult.ToString)
            '    Loader.SIS.CreatePropertyFilter("fCombine", "seccion# >= " & seccion_pri.ToString)
            '    Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_AND)
            '    If Loader.SIS.ScanList("lDelete", "lZF", "fProperty", "") > 0 Then
            '        Loader.SIS.CombineLists("lDelete", "lZF", "lDelete", SIS_BOOLEAN_XOR)
            '        Try
            '            Loader.SIS.Delete("lDelete")
            '        Catch
            '        End Try
            '    End If
            'End If
            'Loader.SIS.CreatePropertyFilter("fProperty", "_FC& = 40")
            'Loader.SIS.ScanList("lDelete", "lPlanoContent", "fProperty", "")
            'Loader.SIS.CombineLists("lDelete", "lDelete", "lAreaZF", SIS_BOOLEAN_DIFF)
            'Try
            '    Loader.SIS.Delete("lDelete")
            'Catch
            'End Try
            Loader.SIS.DeselectAll()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Function lTemplate()
        Try
            EmptyList("lTemplate")
            Try
                Loader.SIS.RecallNolItem(Loader.SIS.GetListItemStr("lMarco", 0, "template$"))
            Catch ex As Exception
                MsgBox("Template no encontrado", MsgBoxStyle.Exclamation)
                Loader.SIS.RemoveOverlay(0)
                Return False
            End Try
            Loader.SIS.AddToList("lTemplate")
            Loader.SIS.MoveList("lTemplate", 0, 0, 0, 0, scaleFactor)
            Loader.SIS.DeselectAll()
            Loader.SIS.SelectList("lTemplate")
            Loader.SIS.DoCommand("AComExplodeGroup")
            Loader.SIS.CreateListFromSelection("lTemplate")
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=944")
            If Not Loader.SIS.ScanList("lHook", "lTemplate", "fProperty", "") = 1 Then
                MsgBox("El template no tiene un hookPoint", MsgBoxStyle.Exclamation)
                Loader.SIS.RemoveOverlay(0)
                Return False
            End If
            Loader.SIS.OpenList("lHook", 0)
            Loader.SIS.MoveList("lTemplate", xPlan - Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"), yPlan - Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#"), 0, 0, 1)
            For i = 1 To Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1
                Loader.SIS.SetInt(SIS_OT_OVERLAY, i, "_status&", 0)
            Next
            Loader.SIS.DeselectAll()
            Loader.SIS.Delete("lHook")
            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try
    End Function

    Private Sub Bloque(ByVal bloque As String, ByVal FC As Integer)
        Try
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=" & FC)
            If Loader.SIS.ScanList("lBloque", "lTemplate", "fProperty", "") = 1 Then
                Loader.SIS.OpenList("lBloque", 0)
                Dim x = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim y = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Loader.SIS.CreatePoint(x, y, 0, bloque, 0, 1)
                Loader.SIS.UpdateItem()
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComExplodeShape")
                Loader.SIS.CloseItem()
                Loader.SIS.DeselectAll()
                Loader.SIS.Delete("lBloque")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub MarcoCoords()
        Try
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=948")
            If Loader.SIS.ScanList("lMarcoCoord", "lTemplate", "fProperty", "") > 0 Then
                Dim coordShift = 100 * scaleFactor
                For i = 0 To Loader.SIS.GetListSize("lMarcoCoord") - 1
                    Loader.SIS.OpenList("lMarcoCoord", i)
                    If Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "axis$") = "y" Then
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", "Y = " + (yPlan + coordShift * Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "pos&")).ToString("N0"))
                    ElseIf Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "axis$") = "x" Then
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", "X = " + (xPlan + coordShift * Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "pos&")).ToString("N0"))
                    End If
                    Loader.SIS.CloseItem()
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub MarcoDistribuccion()
        Try
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=930")
            If Loader.SIS.ScanList("lMarcoDistribuccion", "lTemplate", "fProperty", "") = 1 Then

                'get extent of lMarcoDistribuccion
                Loader.SIS.OpenList("lMarcoDistribuccion", 0)
                Dim xS = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                Dim yS = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
                Dim xO = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim yO = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Dim x1, x2, y1, y2, z, scale As Double

                'get extent of lDistribuccion
                Dim xM As Double = 0
                Dim yM As Double = 0
                Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lDistribuccion"))
                If (xS / (x2 - x1)) < (yS / (y2 - y1)) Then
                    scale = xS / (x2 - x1)
                Else
                    scale = yS / (y2 - y1)
                End If
                Loader.SIS.CopyListItems("lDistribuccion")
                Loader.SIS.MoveList("lDistribuccion", 0, 0, 0, 0, scale)
                Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lDistribuccion"))
                Loader.SIS.MoveList("lDistribuccion", xO - ((x1 + x2) / 2), yO - ((y1 + y2) / 2), 0, 0, 1)

                'delete lMarcoDistribuccion
                Loader.SIS.Delete("lMarcoDistribuccion")

                'delete Marco Distribuccion
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=809")
                If Loader.SIS.ScanList("lMarcoDistribuccion", "lDistribuccion", "fProperty", "") = 1 Then
                    Loader.SIS.Delete("lMarcoDistribuccion")
                End If

                'create planoNUM
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=808")
                If Loader.SIS.ScanList("lMarcoPlantilla", "lDistribuccion", "fProperty", "") > 0 Then
                    For i = 0 To Loader.SIS.GetListSize("lMarcoPlantilla") - 1
                        Loader.SIS.OpenList("lMarcoPlantilla", i)
                        Dim planoNUMsuffix As String = ""
                        Try
                            planoNUMsuffix = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "planoNUMsuffix$")
                        Catch
                        End Try
                        Loader.SIS.CreateBoxText(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"), Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#"), 0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#") / 2, Loader.SIS.GetInt(SIS_OT_CURITEM, 0, "planoNUM&").ToString() + planoNUMsuffix)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_DISTRIBUCION")
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 820)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                        Loader.SIS.CloseItem()
                    Next i
                End If

                'set Marco Plantilla Actual
                Loader.SIS.CreatePropertyFilter("fProperty", "planoID$='" & Loader.SIS.GetListItemStr("lMarco", 0, "planoID$") & "'")
                If Loader.SIS.ScanList("lDistribuccionActual", "lDistribuccion", "fProperty", "") = 1 Then
                    Loader.SIS.OpenList("lDistribuccionActual", 0)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_DISTRIBUCION")
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 810)
                End If

                'set Marco Plantilla Actual Txt
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=818")
                If Loader.SIS.ScanList("lDistribuccionActualTxt", "lDistribuccion", "fProperty", "") = 1 Then
                    Loader.SIS.OpenList("lDistribuccionActualTxt", 0)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Loader.SIS.GetListItemStr("lMarco", 0, "planoID$"))
                    Loader.SIS.CloseItem()
                End If

                'draw line between plantilla actual (810) and plantilla actual zoom (812)
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=810")
                If Loader.SIS.ScanList("lMarcoPlantillaActual", "lDistribuccion", "fProperty", "") = 1 Then
                    Loader.SIS.OpenList("lMarcoPlantillaActual", 0)
                    Loader.SIS.MoveTo(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"), Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#"), 0)
                    Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=812")
                    If Loader.SIS.ScanList("lMarcoPlantillaActualZoom", "lDistribuccion", "fProperty", "") = 1 Then
                        Loader.SIS.OpenList("lMarcoPlantillaActualZoom", 0)
                        Loader.SIS.LineTo(Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"), Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#"), 0)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_DISTRIBUCION")
                        Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 806)
                        Loader.SIS.UpdateItem()
                        Loader.SIS.DeselectAll()
                        Loader.SIS.SelectItem()
                        Loader.SIS.CreateListFromSelection("lSnip")
                        Loader.SIS.OpenList("lMarcoPlantillaActual", 0)
                        Loader.SIS.SnipGeometry("lSnip", True)
                        Loader.SIS.OpenList("lMarcoPlantillaActualZoom", 0)
                        Loader.SIS.SnipGeometry("lSnip", True)
                        Loader.SIS.CloseItem()
                        Loader.SIS.DeselectAll()
                    End If
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub MarcoCroquis()
        Try

            'exit if lMarcoCroquis is empty
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=928")
            If Not Loader.SIS.ScanList("lMarcoCroquis", "lTemplate", "fProperty", "") = 1 Then Exit Sub

            Dim x1, x2, xS, xO, y1, y2, yS, yO, z, scale, scaleInv As Double

            'get extent of croquis frame
            Loader.SIS.OpenList("lMarcoCroquis", 0)
            xS = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
            yS = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
            xO = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
            yO = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")

            'get extent of marco frame
            Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lMarco"))

            'set scale
            If (xS / (x2 - x1)) < (yS / (y2 - y1)) Then
                scale = ((x2 - x1) / xS) * 1.2
                scaleInv = xS / (xS * scale) / 2
            Else
                scale = ((y2 - y1) / yS) * 1.2
                scaleInv = yS / (yS * scale) / 2
            End If

            'create croquis extent
            Loader.SIS.CreateInternalOverlay("tmpCroquis", 0)
            Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", 0)
            Loader.SIS.CreateRectangle(((x1 + x2) / 2) - xS * scale, ((y1 + y2) / 2) - yS * scale, ((x1 + x2) / 2) + xS * scale, ((y1 + y2) / 2) + yS * scale)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_TEMPLATE")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 928)
            Loader.SIS.UpdateItem()
            Loader.SIS.CreateLocusFromItem("locusCrossby", SIS_GT_CROSSBY, SIS_GM_GEOMETRY)
            Loader.SIS.CreateLocusFromItem("locusIntersect", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)

            'get lCroquis
            EmptyList("lCroquis")
            Loader.SIS.CreateClassTreeFilter("fLine", "-Item +Line")
            For i = 0 To Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1
                If Loader.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$") = Loader.SIS.GetListItemStr("lMarco", 0, "croquisLayer$") Then
                    If Loader.SIS.ScanOverlay("lCroquis", i, "fLine", "locusIntersect") > 0 Then
                        Exit For
                    End If
                End If
            Next
            Loader.SIS.AddToList("lCroquis")

            'copy and cut lCroquis
            Loader.SIS.CopyListItems("lCroquis")
            If Loader.SIS.ScanList("lSnip", "lCroquis", "", "locusCrossby") Then Loader.SIS.SnipGeometry("lSnip", False)

            'create
            Loader.SIS.CreateRectangle(x1, y1, x2, y2)
            Loader.SIS.UpdateItem()
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CROQUIS")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 602)
            Loader.SIS.UpdateItem()
            Loader.SIS.AddToList("lCroquis")

            'decompose multilines after cutting lines
            Loader.SIS.CreateClassTreeFilter("fMultiLine", "-Item +MultiLine")
            If Loader.SIS.ScanList("lMultiLine", "lCroquis", "fMultiLine", "") > 0 Then
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectList("lMultiLine")
                Loader.SIS.DoCommand("AComDecompose")
            End If

            'move and copy lCroquis
            Loader.SIS.MoveList("lCroquis", 0, 0, 0, 0, scaleInv)
            Loader.SIS.SplitExtent(x1, y1, z, x2, y2, z, Loader.SIS.GetListExtent("lCroquis"))
            Loader.SIS.MoveList("lCroquis", xO - ((x1 + x2) / 2), yO - ((y1 + y2) / 2), 0, 0, 1)
            Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", 1)
            Loader.SIS.CopyListItems("lCroquis")
            Loader.SIS.RemoveOverlay(0)
            Loader.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", 0)

            'proyecto text
            Dim x, y, a As Double
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=603")
            If Loader.SIS.ScanList("lEje", "lCroquis", "fProperty", "") > 0 Then
                For i = 0 To Loader.SIS.GetListSize("lEje") - 1
                    Loader.SIS.OpenList("lEje", i)
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2))
                    a = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2)
                    Dim aDeg = a * (180 / Math.PI)
                    If aDeg > 90 Or aDeg < -90 Then
                        aDeg += 180
                    End If
                    Loader.SIS.CreateText(x, y, 0, Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "tramo$"))
                    Loader.SIS.UpdateItem()
                    Loader.SIS.DeselectAll()
                    Loader.SIS.SelectItem()
                    Loader.SIS.DoCommand("AComTextToBox")
                    Loader.SIS.OpenSel(0)
                    Dim shift As Double = (Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#") / 2) + (7 * scaleFactor)
                    Loader.SIS.DoCommand("AComBoxToText")
                    Loader.SIS.OpenSel(0)
                    Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", aDeg)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CROQUIS")
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 620)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.DoCommand("AComTextToBox")
                    Loader.SIS.CreatePoint(x + (shift * Math.Cos(a)), y + (shift * Math.Sin(a)), 0, "", a, 1)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CROQUIS")
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 617)
                    Loader.SIS.CloseItem()
                Next
                Loader.SIS.Delete("lEje")
            End If

            'municipio text
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=615")
            Dim municipio_izq As String = "municipio_izq"
            Dim municipio_der As String = "municipio_der"
            If Loader.SIS.ScanList("lMunicipio", "lCroquis", "fProperty", "") > 0 Then
                For i = 0 To Loader.SIS.GetListSize("lMunicipio") - 1
                    Loader.SIS.OpenList("lMunicipio", i)
                    Try
                        municipio_izq = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "municipio_izq$")
                        municipio_der = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "municipio_der$")
                    Catch
                    End Try
                    Loader.SIS.SplitPos(x, y, z, Loader.SIS.GetGeomPosFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2))
                    a = Loader.SIS.GetGeomAngleFromLength(0, Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_length#") / 2)
                    Loader.SIS.CreateText(x + ((scaleFactor * 20) * Math.Cos(a + 1.57079)), y + ((scaleFactor * 20) * Math.Sin(a + 1.57079)), 0, "Municipio" + Environment.NewLine + municipio_izq)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CROQUIS")
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 618)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                    Loader.SIS.CloseItem()
                    Loader.SIS.CreateText(x + ((scaleFactor * 20) * Math.Cos(a - 1.57079)), y + ((scaleFactor * 20) * Math.Sin(a - 1.57079)), 0, "Municipio" + Environment.NewLine + municipio_der)
                    Loader.SIS.UpdateItem()
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CROQUIS")
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 618)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                    Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                    Loader.SIS.CloseItem()
                Next
            Else
                Loader.SIS.CreateText(xO, yO - (yS / 4), 0, "Municipio" + Environment.NewLine + Loader.SIS.GetListItemStr("lMarco", 0, "municipio_croquis$"))
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_CROQUIS")
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 618)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_MIDDLE)
                Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                Loader.SIS.CloseItem()
            End If

            'remove marcos
            Loader.SIS.Delete("lMarcoCroquis")
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=928")
            Loader.SIS.ScanList("lMarcoCroquis", "lCroquis", "fProperty", "")
            Loader.SIS.Delete("lMarcoCroquis")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Simbologia()
        Try
            EmptyList("lSimbologiaItems")
            Dim asimbologia() As String = Split(Loader.SIS.GetListItemStr("lMarco", 0, "simbologia$"), ",")
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=940")
            If Loader.SIS.ScanList("lSimbologia", "lTemplate", "fProperty", "") > 0 Then
                Loader.SIS.OpenListCursor("cursor", "lSimbologia", "elementID$")
                Loader.SIS.OpenSortedCursor("scursor", "cursor", 0, True)
                Dim x, y As Double
                Dim n As Integer = 0
                Do
                    For i = n To asimbologia.Length - 1
                        Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=" & asimbologia(i))
                        If Loader.SIS.ScanList("lS", "lPlanoContent", "fProperty", "") > 0 Then
                            Loader.SIS.OpenCursorItem("scursor")
                            x = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                            y = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                            Loader.SIS.CreatePoint(x, y, 0, "SIMBOLOGIA." & asimbologia(i), 0, 1)
                            Loader.SIS.UpdateItem()
                            Loader.SIS.DeselectAll()
                            Loader.SIS.SelectItem()
                            Loader.SIS.DoCommand("AComExplodeShape")
                            Loader.SIS.CreateListFromSelection("lAdd")
                            Loader.SIS.CombineLists("lSimbologiaItems", "lSimbologiaItems", "lAdd", SIS_BOOLEAN_OR)
                            Loader.SIS.CloseItem()
                            n += 1
                            Exit For
                        End If
                        n += 1
                    Next i
                    If n = asimbologia.Length Then Exit Do
                Loop Until Loader.SIS.MoveCursor("scursor", 1) = 0
                Loader.SIS.Delete("lSimbologia")
                Loader.SIS.DeselectAll()
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub SimbologiaText()
        Try
            'curva maestra
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_curva_maestra'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=142")
                Loader.SIS.CreateClassTreeFilter("fText", "-Item +Text")
                Loader.SIS.CombineFilter("fProperty", "fProperty", "fText", SIS_BOOLEAN_AND)
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "_text$"))
                Else
                    Loader.SIS.Delete("lSimbologiaItem")
                End If
            End If

            'cota foto
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_cota_foto'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=141")
                Loader.SIS.CreateClassTreeFilter("fText", "-Item +Text")
                Loader.SIS.CombineFilter("fProperty", "fProperty", "fText", SIS_BOOLEAN_AND)
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "_text$"))
                End If
            End If

            'cota bati
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_cota_bati'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=140")
                Loader.SIS.CreateClassTreeFilter("fText", "-Item +Text")
                Loader.SIS.CombineFilter("fProperty", "fProperty", "fText", SIS_BOOLEAN_AND)
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "_text$"))
                End If
            End If

            'secciones
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_secciones'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=26")
                Loader.SIS.CreateClassTreeFilter("fText", "-Item +Text")
                Loader.SIS.CombineFilter("fProperty", "fProperty", "fText", SIS_BOOLEAN_AND)
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "_text$"))
                End If
            End If

            'limite namo
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_namo'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 2 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=21")
                Loader.SIS.CreatePropertyFilter("fCombine", "lado$='NI'")
                Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_AND)
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 1 Then
                    Loader.SIS.OpenList("lSimbologiaItem", 0)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "pv$"))
                    Loader.SIS.CloseItem()
                    Loader.SIS.OpenList("lSimbologiaItem", 1)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 1, "pv$"))
                    Loader.SIS.CloseItem()
                Else
                    Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=21")
                    Loader.SIS.CreatePropertyFilter("fCombine", "lado$='ND'")
                    Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_AND)
                    If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 1 Then
                        Loader.SIS.OpenList("lSimbologiaItem", 0)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "pv$"))
                        Loader.SIS.CloseItem()
                        Loader.SIS.OpenList("lSimbologiaItem", 1)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 1, "pv$"))
                        Loader.SIS.CloseItem()
                    End If
                End If
            End If

            'limite zf
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_zf'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 2 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=42")
                Loader.SIS.CreatePropertyFilter("fCombine", "lado$='ZI'")
                Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_AND)
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 1 Then
                    Loader.SIS.OpenList("lSimbologiaItem", 0)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "pv$"))
                    Loader.SIS.CloseItem()
                    Loader.SIS.OpenList("lSimbologiaItem", 1)
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 1, "pv$"))
                    Loader.SIS.CloseItem()
                Else
                    Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=42")
                    Loader.SIS.CreatePropertyFilter("fCombine", "lado$='ZD'")
                    Loader.SIS.CombineFilter("fProperty", "fProperty", "fCombine", SIS_BOOLEAN_AND)
                    If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 1 Then
                        Loader.SIS.OpenList("lSimbologiaItem", 0)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "pv$"))
                        Loader.SIS.CloseItem()
                        Loader.SIS.OpenList("lSimbologiaItem", 1)
                        Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 1, "pv$"))
                        Loader.SIS.CloseItem()
                    End If
                End If
            End If

            'placa
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_placa'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=320")
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "control_id$"))
                End If
            End If
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_placa_z'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=320")
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.OpenList("lPlanoItem", 0)
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.Evaluate(SIS_OT_CURITEM, 0, "'Z=' + str(round(_oz#,3))"))
                End If
            End If

            'mojonera
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_mojonera'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=315")
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "control_id$"))
                End If
            End If
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_mojonera_z'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=315")
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.OpenList("lPlanoItem", 0)
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.Evaluate(SIS_OT_CURITEM, 0, "'Z=' + str(round(_oz#,3))"))
                End If
            End If

            'bn
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_bn'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=305")
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "control_id$"))
                End If
            End If
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_bn_z'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=305")
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.OpenList("lPlanoItem", 0)
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.Evaluate(SIS_OT_CURITEM, 0, "'Z=' + str(round(_oz#,3))"))
                End If
            End If

            'vertice
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_vertice'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=335")
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "control_id$"))
                End If
            End If
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_vertice_z'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=335")
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.OpenList("lPlanoItem", 0)
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.Evaluate(SIS_OT_CURITEM, 0, "'Z=' + str(round(_oz#,3))"))
                End If
            End If

            'ce
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_ce'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=310")
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.GetListItemStr("lPlanoItem", 0, "control_id$"))
                End If
            End If
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='txt_ce_z'")
            If Loader.SIS.ScanList("lSimbologiaItem", "lSimbologiaItems", "fProperty", "") = 1 Then
                Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=310")
                If Loader.SIS.ScanList("lPlanoItem", "lPlanoContent", "fProperty", "") > 0 Then
                    Loader.SIS.OpenList("lPlanoItem", 0)
                    Loader.SIS.SetListStr("lSimbologiaItem", "_text$", Loader.SIS.Evaluate(SIS_OT_CURITEM, 0, "'Z=' + str(round(_oz#,3))"))
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Nota()
        Try
            Dim nota = Loader.SIS.GetListItemStr("lMarco", 0, "nota$")
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=934")
            If Loader.SIS.ScanList("lNota", "lTemplate", "fProperty", "") = 1 Then
                Loader.SIS.OpenList("lNota", 0)
                Dim x = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim y = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Loader.SIS.CreatePoint(x, y, 0, nota, 0, 1)
                Loader.SIS.UpdateItem()
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComExplodeShape")
                Loader.SIS.CloseItem()
                Loader.SIS.DeselectAll()
                Loader.SIS.Delete("lNota")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Referencia()
        Try
            Loader.SIS.CreatePropertyFilter("fProperty", "_FC&=938")
            If Loader.SIS.ScanList("lReferencia", "lTemplate", "fProperty", "") = 1 Then
                Loader.SIS.OpenList("lReferencia", 0)
                Dim x = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim y = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Loader.SIS.CreatePoint(x, y, 0, Loader.SIS.GetListItemStr("lMarco", 0, "referenciacionEstacion$"), 0, 1)
                Loader.SIS.UpdateItem()
                Loader.SIS.DeselectAll()
                Loader.SIS.SelectItem()
                Loader.SIS.DoCommand("AComExplodeShape")
                Loader.SIS.CloseItem()
                Loader.SIS.DeselectAll()
                Loader.SIS.Delete("lReferencia")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub SeccionTramo()
        Try
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='seccion_de_tramo'")
            If Loader.SIS.Scan("lElementID", "E", "fProperty", "") = 1 Then
                Loader.SIS.OpenList("lElementID", 0)
                Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", String.Format("(Tramo del km {0} al km {1})", FormatoCadenamiento(seccion_pri), FormatoCadenamiento(seccion_ult)))
                Loader.SIS.CloseItem()
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Function FormatoCadenamiento(ByVal seccion As Double)
        Try
            Dim first = Math.Truncate((seccion / 1000))
            Dim second = Math.Round((seccion - (first * 1000)), 2)
            Dim cadenamiento = String.Format("{0}+{1:000.##}", first, second)
            Return cadenamiento
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return ""
        End Try
    End Function

    Private Sub ScaleTxt()
        Try
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='scaleTXT'")
            If Loader.SIS.Scan("lScaleTxt", "E", "fProperty", "") > 0 Then
                For i = 0 To Loader.SIS.GetListSize("lScaleTxt") - 1
                    Loader.SIS.OpenList("lScaleTxt", i)
                    Dim txt = Loader.SIS.GetStr(SIS_OT_CURITEM, 0, "_text$")
                    Dim scale As Integer
                    If txt.IndexOf(":") = -1 Then
                        scale = Int(txt) * scaleFactor
                        txt = String.Format("{0:###,###}", scale)
                    Else
                        Dim atxt() As String = txt.Split(":")
                        scale = Int(atxt(1)) * scaleFactor
                        txt = String.Format("{0}:{1:###,###}", atxt(0), scale)
                    End If
                    Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", txt)
                    Loader.SIS.CloseItem()
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub PlanoTxt(ByVal elementID As String, ByVal text As String)
        Try
            Loader.SIS.CreatePropertyFilter("fProperty", "elementID$='" & elementID & "'")
            If Loader.SIS.Scan("lElementID", "E", "fProperty", "") > 0 Then
                Loader.SIS.SetListStr("lElementID", "_text$", text)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub EmptyList(ByVal List As String)
        Try
            Loader.SIS.EmptyList(List)
        Catch
        End Try
    End Sub

End Class
