Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System.Windows.Forms
Imports System.IO

Public Class Conagua_Marco

    Private scaleFactor As Double
    Private template As String

    Public Sub New()
        InitializeComponent()
        Try
            cbEscala.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub btnCrearMarco_Click(sender As System.Object, e As System.EventArgs) Handles btnCrearMarco.Click
        Try
            Dim projection As String = ""
            Try
                projection = Loader.SIS.GetStr(SIS_OT_DATASET, Loader.SIS.GetInt(SIS_OT_OVERLAY, Loader.SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"), "_nDataset&"), "_projection$")
            Catch
                MsgBox("No hay overlay activo")
                Loader.SIS.Dispose()
                Loader.SIS = Nothing
                Exit Sub
            End Try
            If projection.Contains("WGS 84.UTM zone 11N") = True Then
                template = "plantilla_rios_utm11"
            ElseIf projection.Contains("WGS 84.UTM zone 12N") = True Then
                template = "plantilla_rios_utm12"
            ElseIf projection.Contains("WGS 84.UTM zone 13N") = True Then
                template = "plantilla_rios_utm13"
            ElseIf projection.Contains("WGS 84.UTM zone 14N") = True Then
                template = "plantilla_rios_utm14"
            ElseIf projection.Contains("WGS 84.UTM zone 15N") = True Then
                template = "plantilla_rios_utm15"
            ElseIf projection.Contains("WGS 84.UTM zone 16N") = True Then
                template = "plantilla_rios_utm16"
            Else
                MsgBox("Proyeccion erronea")
                Loader.SIS.Dispose()
                Loader.SIS = Nothing
                Exit Sub
            End If
            scaleFactor = CDbl(cbEscala.SelectedItem) / 1000
            Loader.APP.AddTrigger("AComPlaceGroup::End", New SisTriggerHandler(AddressOf PlaceFrame_End))
            Loader.APP.AddTrigger("AComPlaceGroup::Snap", New SisTriggerHandler(AddressOf PlaceFrame_Snap))
            Loader.SIS.CreateGroup("")
            Loader.SIS.CreateRectangle(-500 * scaleFactor, -400 * scaleFactor, 500 * scaleFactor, 400 * scaleFactor)
            Loader.SIS.UpdateItem()
            Loader.SIS.Dispose()
            Loader.SIS = Nothing
            Me.Close()
        Catch ex As Exception
            Loader.SIS.Dispose()
            Loader.SIS = Nothing
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub PlaceFrame_End(ByVal sender As Object, ByVal e As SisTriggerArgs)
        Loader.APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceFrame_End)
        Loader.APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceFrame_Snap)
    End Sub

    Private Sub PlaceFrame_Snap(ByVal sender As Object, ByVal e As SisTriggerArgs)
        Try
            Loader.APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceFrame_End)
            Loader.APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceFrame_Snap)
            SendKeys.SendWait("{ENTER}")
            Loader.SIS.CreateListFromSelection("list")
            Loader.SIS.OpenList("list", 0)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "_featureTable$", "CONAGUA_DISTRIBUCION")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "_FC&", 808)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "proyecto$", "Proyecto")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "planoID$", "PlanoID")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "template$", template)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "croquisLayer$", "###")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "tramo$", "###")
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "planoNUM&", 0)
            Loader.SIS.SetInt(SIS_OT_CURITEM, 0, "planoNUMtotal&", 999)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "planoNUMsuffix$", "")
            Loader.SIS.SetFlt(SIS_OT_CURITEM, 0, "scaleFactor#", scaleFactor)
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "bloque_registro_sgt$", "BLOQUE.REGISTRO_SGT.registro_sgt")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "bloque_registro_oc$", "BLOQUE.REGISTRO_OC.registro_oc")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "bloque_responsable$", "BLOQUE.RESPONSABLE.responsable")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "bloque_titulo$", "BLOQUE.TITULO.titulo")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "nota$", "NOTAS.nota")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "referenciacionEstacion$", "REFERENCIA.referencia")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "entidad$", "Entidad")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "localidad$", "Localidad")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "municipio$", "Municipio")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "municipio_croquis$", "Municipio (Est.)")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "simbologia$", "202,204,206,208,210,212,214,216,218,220,222,224,226,228,230,247,232,234,236,238,240,242,244,245,246,248,290,260,130,131,141,140,041,020,025,040,320,315,305,335,310")
            Loader.SIS.SetStr(SIS_OT_CURITEM, 0, "fecha$", DateTime.Today.ToString("d"))
            Loader.SIS.UpdateItem()
            Dim xO = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
            Dim yO = Loader.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
            Dim xNew = Math.Round(xO / (100 * scaleFactor), 0) * (100 * scaleFactor)
            Dim yNew = Math.Round(yO / (100 * scaleFactor), 0) * (100 * scaleFactor)
            Loader.SIS.MoveList("list", xNew - xO, yNew - yO, 0, 0, 1)
            Loader.SIS.Dispose()
            Loader.SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

End Class