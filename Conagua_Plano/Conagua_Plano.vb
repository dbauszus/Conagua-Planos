Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System.IO
Imports System.Windows.Forms

<GisLinkProgram("Conagua_Plano")> _
Public Class Loader

    Public Shared APP As SisApplication
    Private Shared _sis As MapModeller

    Public Shared Property SIS As MapModeller
        Get
            If _sis Is Nothing Then _sis = APP.TakeoverMapManager
            Return _sis
        End Get
        Set(ByVal value As MapModeller)
            _sis = value
        End Set
    End Property

    Public Sub New(ByVal SISApplication As SisApplication)
        APP = SISApplication

        SIS.CreatePropertyFilter("btnPlano", "_FC& = 808")
        SIS.CreatePropertyFilter("ctxDestino", "_FC& = 250")
        SIS.CreatePropertyFilter("ctxMargen", "_FC& = 28")
        SIS.CreatePropertyFilter("ctxControl", "_FC& = 305")
        SIS.CreatePropertyFilter("fCombine", "_FC& = 310")
        SIS.CombineFilter("ctxControl", "ctxControl", "fCombine", SIS_BOOLEAN_OR)
        SIS.CreatePropertyFilter("fCombine", "_FC& = 315")
        SIS.CombineFilter("ctxControl", "ctxControl", "fCombine", SIS_BOOLEAN_OR)
        SIS.CreatePropertyFilter("fCombine", "_FC& = 320")
        SIS.CombineFilter("ctxControl", "ctxControl", "fCombine", SIS_BOOLEAN_OR)
        SIS.CreatePropertyFilter("fCombine", "_FC& = 335")
        SIS.CombineFilter("ctxControl", "ctxControl", "fCombine", SIS_BOOLEAN_OR)
        SIS.CreatePropertyFilter("ctxControlFlecha", "_FC& = 330")
        SIS.CreatePropertyFilter("ctxCotaTxt", "_FC& = 101")
        SIS.CreatePropertyFilter("fCombine", "_FC& = 102")
        SIS.CombineFilter("ctxCotaTxt", "ctxCotaTxt", "fCombine", SIS_BOOLEAN_OR)
        'SIS.CreateClassTreeFilter("", "-Item +Text")
        SIS.Dispose()

        Dim group As SisRibbonGroup = APP.RibbonGroup
        group.Text = "CONAGUA Planos"

        Dim btnConagua_Marco As SisRibbonButton = New SisRibbonButton("Marco", New SisClickHandler(AddressOf subConagua_Marco))
        btnConagua_Marco.LargeImage = True
        btnConagua_Marco.Icon = My.Resources.MARCO
        btnConagua_Marco.Help = "Marco"
        group.Controls.Add(btnConagua_Marco)

        Dim btnConagua_Plano As SisRibbonButton = New SisRibbonButton("Plano", New SisClickHandler(AddressOf subConagua_CrearPlano))
        btnConagua_Plano.LargeImage = True
        btnConagua_Plano.Icon = My.Resources.PLANO
        btnConagua_Plano.Help = "Plano"
        btnConagua_Plano.MinSelection = 1
        btnConagua_Plano.MaxSelection = 1
        btnConagua_Plano.Filter = "btnPlano"
        group.Controls.Add(btnConagua_Plano)

        Dim btnConagua_OverlapTxt As SisRibbonButton = New SisRibbonButton("Overlap Txt", New SisClickHandler(AddressOf subConagua_OverlapTxt))
        btnConagua_OverlapTxt.Icon = My.Resources.OVERLAP_TXT
        group.Controls.Add(btnConagua_OverlapTxt)

        Dim btnConagua_CurvaTxt As SisRibbonButton = New SisRibbonButton("Curva Txt", New SisClickHandler(AddressOf subConagua_CurvaTxt))
        btnConagua_CurvaTxt.Icon = My.Resources.CURVA_TXT
        group.Controls.Add(btnConagua_CurvaTxt)

        Dim ctxCotaTxt As SisMenuItem = New SisMenuItem("Cota Txt", New SisClickHandler(AddressOf subConagua_CotaTxt))
        ctxCotaTxt.MinSelection = 1
        ctxCotaTxt.Filter = "ctxCotaTxt"
        ctxCotaTxt.Image = My.Resources.TXT_CONTROL
        APP.ContextMenu.MenuItems.Add(ctxCotaTxt)

        Dim ctxFlechaDestino As SisMenuItem = New SisMenuItem("Flecha Destino", New SisClickHandler(AddressOf subConagua_FlechaDestino))
        ctxFlechaDestino.Filter = "ctxDestino"
        ctxFlechaDestino.MinSelection = 1
        ctxFlechaDestino.MaxSelection = 1
        ctxFlechaDestino.Image = My.Resources.FLECHA_DESTINO
        APP.ContextMenu.MenuItems.Add(ctxFlechaDestino)

        Dim ctxFlechaMargen As SisMenuItem = New SisMenuItem("Flecha Margen", New SisClickHandler(AddressOf subConagua_FlechaMargen))
        ctxFlechaMargen.Filter = "ctxMargen"
        ctxFlechaMargen.MinSelection = 1
        ctxFlechaMargen.MaxSelection = 2
        ctxFlechaMargen.Image = My.Resources.FLECHA_DESTINO
        APP.ContextMenu.MenuItems.Add(ctxFlechaMargen)

        Dim ctxControl As SisMenuItem = New SisMenuItem("Txt Control", New SisClickHandler(AddressOf subConagua_Control))
        ctxControl.Filter = "ctxControl"
        ctxControl.MinSelection = 1
        ctxControl.Image = My.Resources.TXT_CONTROL
        APP.ContextMenu.MenuItems.Add(ctxControl)

        Dim ctxControlFlecha As SisMenuItem = New SisMenuItem("Flecha Control", New SisClickHandler(AddressOf subConagua_ControlFlecha))
        ctxControlFlecha.Filter = "ctxControlFlecha"
        ctxControlFlecha.MinSelection = 1
        ctxControlFlecha.MaxSelection = 2
        ctxControlFlecha.Image = My.Resources.FLECHA_DESTINO
        APP.ContextMenu.MenuItems.Add(ctxControlFlecha)

        'Dim btnConagua_CotaFoto As SisRibbonButton = New SisRibbonButton("Cota Foto", New SisClickHandler(AddressOf subConagua_CotaFoto))
        'group.Controls.Add(btnConagua_CotaFoto)

        'Dim btnConagua_CotaBati As SisRibbonButton = New SisRibbonButton("Cota Bati", New SisClickHandler(AddressOf subConagua_CotaBati))
        'group.Controls.Add(btnConagua_CotaBati)

    End Sub

    Private Sub subConagua_Marco(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Conagua_Marco As New Conagua_Marco With {.TopMost = True}
            Conagua_Marco.Show()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subConagua_CrearPlano(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Conagua_CrearPlano As New Conagua_CrearPlano
            Conagua_CrearPlano.CrearPlano()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subConagua_OverlapTxt(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Conagua_Elements As New Conagua_Elements
            Conagua_Elements.Conagua_OverlapTxt()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subConagua_CurvaTxt(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Conagua_Elements As New Conagua_Elements
            Conagua_Elements.Conagua_CurvaTxt()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subConagua_CotaTxt(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Conagua_Elements As New Conagua_Elements
            Conagua_Elements.Conagua_CotaTxt()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subConagua_FlechaDestino(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Conagua_Elements As New Conagua_Elements
            Conagua_Elements.Conagua_FlechaDestino()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subConagua_FlechaMargen(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Conagua_Elements As New Conagua_Elements
            Conagua_Elements.Conagua_FlechaTxt("CONAGUA_ZF", 15)
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subConagua_Control(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Conagua_Elements As New Conagua_Elements
            Conagua_Elements.Conagua_Control()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    Private Sub subConagua_ControlFlecha(ByVal sender As Object, ByVal e As SisClickArgs)
        Try
            SIS = e.MapModeller
            Dim Conagua_Elements As New Conagua_Elements
            Conagua_Elements.Conagua_ControlFlecha()
            SIS.Dispose()
            SIS = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            SIS.Dispose()
            SIS = Nothing
        End Try
    End Sub

    'Private Sub subConagua_CotaFoto(ByVal sender As Object, ByVal e As SisClickArgs)
    '    Try
    '        SIS = e.MapModeller
    '        Dim Conagua_Elements As New Conagua_Elements
    '        Conagua_Elements.Conagua_CotaFoto()
    '        SIS.Dispose()
    '        SIS = Nothing
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '        SIS.Dispose()
    '        SIS = Nothing
    '    End Try
    'End Sub

    'Private Sub subConagua_CotaBati(ByVal sender As Object, ByVal e As SisClickArgs)
    '    Try
    '        SIS = e.MapModeller
    '        Dim Conagua_Elements As New Conagua_Elements
    '        Conagua_Elements.Conagua_CotaBati()
    '        SIS.Dispose()
    '        SIS = Nothing
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '        SIS.Dispose()
    '        SIS = Nothing
    '    End Try
    'End Sub

End Class