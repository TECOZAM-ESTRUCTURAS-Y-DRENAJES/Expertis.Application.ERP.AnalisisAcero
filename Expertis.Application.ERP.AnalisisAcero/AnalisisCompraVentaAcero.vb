Imports System.Windows.Forms
Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Engine.UI
Imports Solmicro.Expertis.Application.ERP.GlobalActions

Public Class AnalisisCompraVentaAcero

    Public Sub creardatatable()
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim dc As New DataColumn("Tipo")
        dt.Columns.Add(dc)
        dc = New DataColumn("Mes")
        dt.Columns.Add(dc)
        dc = New DataColumn("Anio")
        dt.Columns.Add(dc)
        dc = New DataColumn("Obra")
        dt.Columns.Add(dc)
        dc = New DataColumn("D8", System.Type.GetType("System.Double"))
        dt.Columns.Add(dc)
        dc = New DataColumn("PrecD8")
        dt.Columns.Add(dc)
        dc = New DataColumn("D10")
        dt.Columns.Add(dc)
        dc = New DataColumn("PrecD10")
        dt.Columns.Add(dc)
        dc = New DataColumn("D12")
        dt.Columns.Add(dc)
        dc = New DataColumn("PrecD12")
        dt.Columns.Add(dc)
        dc = New DataColumn("D16")
        dt.Columns.Add(dc)
        dc = New DataColumn("PrecD16")
        dt.Columns.Add(dc)
        dc = New DataColumn("D20")
        dt.Columns.Add(dc)
        dc = New DataColumn("PrecD20")
        dt.Columns.Add(dc)
        dc = New DataColumn("D25")
        dt.Columns.Add(dc)
        dc = New DataColumn("PrecD25")
        dt.Columns.Add(dc)
        dc = New DataColumn("D32")
        dt.Columns.Add(dc)
        dc = New DataColumn("Prec32")
        dt.Columns.Add(dc)
        dc = New DataColumn("Totales")
        dt.Columns.Add(dc)
        'dr = dt.NewRow

        Dim MesMin As Integer
        Dim AnioMin As Integer
        Dim MesMax As Integer
        Dim AnioMax As Integer

        Dim FInicio As Date = nz(cbxFInicio.Value, "01/01/2000")
        Dim FFin As Date = Nz(cbxFFin.Value, Now())

        Dim MI As Integer = 0
        If FInicio.Day > 20 Then
            MI = Month(FInicio) + 1
        Else
            MI = Month(FInicio)
        End If


        Dim AI As Integer = Year(FInicio)
        Dim MF As Integer = Month(FFin)
        Dim AF As Integer = Year(FFin)
        Dim TotMes As Integer = DateDiff(DateInterval.Month, FInicio, FFin)

        MessageBox.Show("Vamos a sacar el resultado de " & TotMes & " meses")



        Dim strwhere As String = ""
        If AdvObra.Text = "" Then
        Else
            strwhere = "where NObra ='" & AdvObra.Text & "' "
        End If



        Dim strSelectVenta = "SELECT IdObra, NObra, Mes, Anio, SUM(D8) AS D8, SUM(D10) AS D10, SUM(D12) AS D12, SUM(D16) AS D16, " & _
        "SUM(D20) AS D20, SUM(D25) AS D25, SUM(D32) AS D32 FROM  (SELECT IdObra, NObra, " & _
        "CASE WHEN day(fecha) > 20 THEN month(fecha) + 1 ELSE month(fecha) END AS Mes, CASE WHEN day(fecha) > 20 AND month(fecha) = 12 " & _
        "THEN year(fecha) + 1 ELSE year(fecha) END AS Anio, " & _
        "SUM(D8) AS D8, SUM(D10) AS D10, SUM(D12) AS D12, SUM(D16) AS D16, SUM(D20) AS D20, SUM(D25) AS D25, SUM(D32) AS D32" & _
        " FROM dbo.vFrmMedicionesObraAcero WHERE (IdObra IN (SELECT IDObra FROM dbo.tbObraCabecera " & _
        "WHERE (NObra NOT LIKE '%C0'))) AND (Estructura NOT IN ('BASCULA', 'BASCULA+TTE+ALAMBRE', 'ALAMBRE')) and fecha between '" & FInicio & "' and '" & FFin & "' " & _
        " GROUP BY IdObra, NObra, Fecha) AS a " & strwhere & " GROUP BY IdObra, NObra, Mes, Anio"

        Dim strSelectCompra = "select DescArticulo,case when day(FechaAlbaran)>20 then month(FechaAlbaran)+1 else month (FechaAlbaran) end as Mes, " & _
        "case when day(FechaAlbaran)>20 and month(FechaAlbaran)=12 then year(FechaAlbaran)+1 else year(FechaAlbaran) end as Anio, " & _
        "case when descarticulo like '%20 mm%' then 'd20' when descarticulo like '%8 mm%' then 'd8' " & _
        "when descarticulo like '%10 mm%' then 'd10' when descarticulo like '%12 mm%' then 'd12' " & _
        "when descarticulo like '%16 mm%' then 'd16' when descarticulo like '%25 mm%' then 'd25' " & _
        "when descarticulo like '%32 mm%' then 'd32' end as Diametro," & _
        " QServida,Precio from VFrmEstadisticaACDetalladaA where FechaAlbaran between '" & FInicio & "' and '" & FFin & "' and " & _
        "IDTipo='06' and DescProveedor not like 'T___%' and  DescProveedor not like 'DP%' and DescProveedor not like 'F%C0%' " & _
        "and (DescArticulo like 'barra%mm%' or DescArticulo like 'encarr%')"



        'MessageBox.Show(strSelectVenta)

        Dim filtro As New Filter
        filtro.Add("Nobra", AdvObra.Text)

        Dim BD As New BE.DataEngine



        'Dim dtVenta As DataTable = AdminData.GetData(strSelectVenta)
        Dim dtVenta As DataTable = BD.RetrieveData(strSelectVenta)



        'Dim dvDistM As New DataView(dtVenta)
        'dvDistM.RowFilter() = "distinct mes"
        'Dim dtDistM As DataTable = dvDistM.Table

        'If dtDistM.Rows.Count > 0 Then
        '    MessageBox.Show("Hay " & dtDistM.Rows.Count & " Meses Distintos")
        'Else
        '    MessageBox.Show("Error No funciona", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        'End If

        'MessageBox.Show(strSelectCompra)

        ' sacamos los kg y precios por diametros

        'Dim dtCompra As DataTable = AdminData.GetData(strSelectCompra, , , "mes,anio")
        Dim dtCompra As DataTable = BD.RetrieveData(strSelectCompra, , , "mes,anio")
        If dtCompra.Rows.Count > 0 Then
            'Dim drDistinct() As DataRow = dtCompra.)
            'MessageBox.Show("Se han obtenido " & drDistinct.Length & " Registros Diferentes")
            'Diametro 20
            For j As Integer = 0 To 2

                Dim drDiametro8() As DataRow = dtCompra.Select("Diametro in ('d8')")
                Dim precioMed8 As Decimal = 0
                Dim kg8 As Decimal = 0
                If drDiametro8.Length > 0 Then
                    'MessageBox.Show("Se han obtenido " & dr20.Length & " Compras de diametro 20")
                    Dim kgpre As Decimal = 0
                    Dim kg As Decimal = 0


                    For i As Integer = 0 To drDiametro8.Length - 1
                        kgpre += drDiametro8(i)("QServida") * drDiametro8(i)("precio")
                        kg += drDiametro8(i)("QServida")

                    Next
                    precioMed8 = Math.Round((kgpre / kg), 4)
                    kg8 = kg

                    'MessageBox.Show("El precio medio es" & precioMed8 & " eur / kg de diametro 20")
                Else

                End If

                'Diametro 10
                Dim drDiametro10() As DataRow = dtCompra.Select("Diametro in ('d10')")
                Dim precioMed10 As Decimal = 0
                Dim kg10 As Decimal = 0
                If drDiametro10.Length > 0 Then
                    'MessageBox.Show("Se han obtenido " & dr20.Length & " Compras de diametro 20")
                    Dim kgpre As Decimal = 0
                    Dim kg As Decimal = 0


                    For i As Integer = 0 To drDiametro10.Length - 1
                        kgpre += drDiametro10(i)("QServida") * drDiametro10(i)("precio")
                        kg += drDiametro10(i)("QServida")

                    Next
                    precioMed10 = Math.Round((kgpre / kg), 4)
                    kg10 = kg

                    'MessageBox.Show("El precio medio es" & precioMed10 & " eur / kg de diametro 20")
                Else

                End If
                'Diametro 12
                Dim drDiametro12() As DataRow = dtCompra.Select("Diametro in ('d12')")
                Dim precioMed12 As Decimal = 0
                Dim kg12 As Decimal = 0
                If drDiametro12.Length > 0 Then
                    'MessageBox.Show("Se han obtenido " & dr20.Length & " Compras de diametro 20")
                    Dim kgpre As Decimal = 0
                    Dim kg As Decimal = 0


                    For i As Integer = 0 To drDiametro12.Length - 1
                        kgpre += drDiametro12(i)("QServida") * drDiametro12(i)("precio")
                        kg += drDiametro12(i)("QServida")

                    Next
                    precioMed12 = Math.Round((kgpre / kg), 4)
                    kg12 = kg

                    'MessageBox.Show("El precio medio es" & precioMed12 & " eur / kg de diametro 20")
                Else

                End If
                'Diametro 16
                Dim drDiametro16() As DataRow = dtCompra.Select("Diametro in ('d16')")
                Dim precioMed16 As Decimal = 0
                Dim kg16 As Decimal = 0
                If drDiametro16.Length > 0 Then
                    'MessageBox.Show("Se han obtenido " & dr20.Length & " Compras de diametro 20")
                    Dim kgpre As Decimal = 0
                    Dim kg As Decimal = 0


                    For i As Integer = 0 To drDiametro16.Length - 1
                        kgpre += drDiametro16(i)("QServida") * drDiametro16(i)("precio")
                        kg += drDiametro16(i)("QServida")

                    Next
                    precioMed16 = Math.Round((kgpre / kg), 4)
                    kg16 = kg

                    'MessageBox.Show("El precio medio es" & precioMed16 & " eur / kg de diametro 20")
                Else

                End If

                'Diametro 20
                Dim drDiametro20() As DataRow = dtCompra.Select("Diametro in ('d20')")
                Dim precioMed20 As Decimal = 0
                Dim kg20 As Decimal = 0
                If drDiametro20.Length > 0 Then
                    'MessageBox.Show("Se han obtenido " & dr20.Length & " Compras de diametro 20")
                    Dim kgpre As Decimal = 0
                    Dim kg As Decimal = 0


                    For i As Integer = 0 To drDiametro20.Length - 1
                        kgpre += drDiametro20(i)("QServida") * drDiametro20(i)("precio")
                        kg += drDiametro20(i)("QServida")

                    Next
                    precioMed20 = Math.Round((kgpre / kg), 4)
                    kg20 = kg

                    'MessageBox.Show("El precio medio es" & precioMed20 & " eur / kg de diametro 20")
                Else

                End If


                'Diametro 25
                Dim drDiametro25() As DataRow = dtCompra.Select("Diametro in ('d25')")
                Dim precioMed25 As Decimal = 0
                Dim kg25 As Decimal = 0
                If drDiametro25.Length > 0 Then
                    'MessageBox.Show("Se han obtenido " & dr20.Length & " Compras de diametro 20")
                    Dim kgpre As Decimal = 0
                    Dim kg As Decimal = 0


                    For i As Integer = 0 To drDiametro25.Length - 1
                        kgpre += drDiametro25(i)("QServida") * drDiametro25(i)("precio")
                        kg += drDiametro25(i)("QServida")

                    Next
                    precioMed25 = Math.Round((kgpre / kg), 4)
                    kg25 = kg

                    'MessageBox.Show("El precio medio es" & precioMed25 & " eur / kg de diametro 20")
                Else

                End If

                'Diametro 32
                Dim drDiametro32() As DataRow = dtCompra.Select("Diametro in ('d32')")
                Dim precioMed32 As Decimal = 0
                Dim kg32
                If drDiametro32.Length > 0 Then
                    'MessageBox.Show("Se han obtenido " & dr20.Length & " Compras de diametro 20")
                    Dim kgpre As Decimal = 0
                    Dim kg As Decimal = 0


                    For i As Integer = 0 To drDiametro32.Length - 1
                        kgpre += drDiametro32(i)("QServida") * drDiametro32(i)("precio")
                        kg += drDiametro32(i)("QServida")

                    Next
                    precioMed32 = Math.Round((kgpre / kg), 4)
                    kg32 = kg

                    'MessageBox.Show("El precio medio es" & precioMed25 & " eur / kg de diametro 20")
                Else

                End If


                'dr = dt.NewRow
                'dr("Tipo") = "Compra"
                'dr("Mes") = drVenta("Mes")
                'dr("Anio") = drVenta("Anio")
                'dr("Obra") = drVenta("NObra")
                'dr("D8") = drVenta("D8")
                'dr("PrecD8") = 0
                'dr("D10") = drVenta("D10")
                'dr("PrecD10") = 0
                'dr("D12") = drVenta("D12")
                'dr("PrecD12") = 0
                'dr("D16") = drVenta("D16")
                'dr("PrecD16") = 0
                'dr("D20") = drVenta("D20")
                'dr("PrecD20") = 0
                'dr("D25") = drVenta("D25")
                'dr("PrecD25") = 0
                'dr("D32") = drVenta("D32")
                'dr("Prec32") = 0
                'dr("Totales") = 0
                'dt.Rows.Add(dr)
            Next
        Else
            MessageBox.Show("No existen registros de ventas")
        End If


        If dtVenta.Rows.Count > 0 Then
            'MessageBox.Show("Se han obtenido " & dtVenta.Rows.Count & " Registros")
            For Each drVenta As DataRow In dtVenta.Rows
                dr = dt.NewRow
                dr("Tipo") = "Venta"
                dr("Mes") = drVenta("Mes")
                dr("Anio") = drVenta("Anio")
                dr("Obra") = drVenta("NObra")
                dr("D8") = drVenta("D8")
                dr("PrecD8") = 0
                dr("D10") = drVenta("D10")
                dr("PrecD10") = 0
                dr("D12") = drVenta("D12")
                dr("PrecD12") = 0
                dr("D16") = drVenta("D16")
                dr("PrecD16") = 0
                dr("D20") = drVenta("D20")
                dr("PrecD20") = 0
                dr("D25") = drVenta("D25")
                dr("PrecD25") = 0
                dr("D32") = drVenta("D32")
                dr("Prec32") = 0
                dr("Totales") = 0
                dt.Rows.Add(dr)



            Next
            Grid.DataSource = dt
        Else
            MessageBox.Show("No existen registros")
        End If




        'dr("Tipo") = "Compra"
        'dr("Mes") = "01"
        'dr("Anio") = "2019"
        'dr("Obra") = "t200"
        'dr("D8") = "0"
        'dr("PrecD8") = "0"
        'dt.Rows.Add(dr)


        'MsgBox("El datatable tiene " & dt.Columns.Count & " columnas")






    End Sub


#Region "ventas"

    Public Sub obtenerVentaAcero()

    End Sub



#End Region



    Private Sub AnalisisCompraVentaAcero_QueryExecuted(ByVal sender As Object, ByRef e As Solmicro.Expertis.Engine.UI.QueryExecutedEventArgs) Handles MyBase.QueryExecuted
        creardatatable()
    End Sub

End Class
