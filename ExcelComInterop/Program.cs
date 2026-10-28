using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelComInterop
{
    internal static class Program
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        // Ruta local predefinida (se puede sobrescribir pasando otra ruta como primer argumento).
        private static readonly string RutaPredefinida = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "Practica2_ExcelCOM",
            "Reporte_Ventas_Trimestral.xlsx");

        private const string NombreHoja = "Ventas";
        private const int PrimeraFilaDatos = 5;
        private const int UltimaFilaDatos = 14;
        private const int FilaTotales = UltimaFilaDatos + 1;

        // Cada objeto COM obtenido de Excel se registra aquí para liberarlo al final.
        // Evita las expresiones encadenadas (hoja.Range[..].Font.Bold = ...) que crean
        // RCW temporales imposibles de liberar.
        private static readonly List<object> ReferenciasCom = new List<object>();

        private static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            string ruta = args.Length > 0 ? args[0] : RutaPredefinida;

            Console.WriteLine("=== Práctica 2: Automatización de Excel con COM Interop ===");

            Excel.Application excel = null;
            uint pidExcel = 0;
            int codigoSalida = 0;

            try
            {
                // a) Instanciación del servidor COM: Excel.Application invisible
                Console.WriteLine("[a] Creando instancia invisible de Excel.Application...");
                excel = Registrar(new Excel.Application());
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.ScreenUpdating = false;

                GetWindowThreadProcessId(new IntPtr(excel.Hwnd), out pidExcel);
                Console.WriteLine("    Excel versión {0} - proceso EXCEL.EXE PID {1}", excel.Version, pidExcel);

                CrearReporte(excel, ruta);
            }
            catch (COMException ex)
            {
                Console.Error.WriteLine("Error COM (0x{0:X8}): {1}", ex.ErrorCode, ex.Message);
                codigoSalida = 1;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                codigoSalida = 1;
            }
            finally
            {
                // f) Liberación de recursos COM
                Console.WriteLine("[f] Cerrando Excel y liberando referencias COM...");
                CerrarExcel(excel);
                int liberadas = LiberarReferencias();
                Console.WriteLine("    Referencias COM liberadas con Marshal.ReleaseComObject: {0}", liberadas);

                // Fuerza al GC a finalizar los RCW que pudieran quedar sin referencia.
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (pidExcel != 0)
                {
                    // Excel puede tardar varios segundos en terminar tras Quit() (add-ins, telemetría).
                    Console.WriteLine("    Esperando el cierre de EXCEL.EXE (PID {0})...", pidExcel);
                    bool cerrado = EsperarCierreProceso(pidExcel, 60000);
                    Console.WriteLine(cerrado
                        ? "    Proceso EXCEL.EXE (PID {0}) finalizado: no quedan procesos huérfanos."
                        : "    ATENCIÓN: el proceso EXCEL.EXE (PID {0}) sigue activo.", pidExcel);
                    if (!cerrado) codigoSalida = 2;
                }
            }

            // Pausa para poder capturar la evidencia (se omite si la entrada está redirigida)
            if (!Console.IsInputRedirected)
            {
                Console.WriteLine();
                Console.Write("Presione una tecla para salir...");
                Console.ReadKey(true);
            }

            return codigoSalida;
        }

        private static void CrearReporte(Excel.Application excel, string ruta)
        {
            // b) Creación del libro y selección de la hoja activa
            Console.WriteLine("[b] Creando libro de trabajo y hoja activa...");
            Excel.Workbooks libros = Registrar(excel.Workbooks);
            Excel.Workbook libro = Registrar(libros.Add());
            Excel.Worksheet hoja = Registrar((Excel.Worksheet)libro.ActiveSheet);
            hoja.Name = NombreHoja;

            try
            {
                // c) Inserción y formato de datos
                Console.WriteLine("[c] Escribiendo datos, fórmulas y formato...");
                EscribirEncabezadoReporte(hoja);
                EscribirTabla(hoja);
                EscribirFormulas(hoja);
                AplicarFormato(hoja);

                // d) Gráfico estadístico embebido en la misma hoja
                Console.WriteLine("[d] Generando gráfico de columnas vinculado a la tabla...");
                CrearGrafico(hoja);

                // e) Exportación y guardado
                Console.WriteLine("[e] Guardando archivo XLSX...");
                string carpeta = Path.GetDirectoryName(ruta);
                if (!string.IsNullOrEmpty(carpeta)) Directory.CreateDirectory(carpeta);
                if (File.Exists(ruta)) File.Delete(ruta);

                libro.SaveAs(ruta, Excel.XlFileFormat.xlOpenXMLWorkbook);
                Console.WriteLine("    Archivo guardado en: " + ruta);
            }
            finally
            {
                libro.Close(false);
            }
        }

        // ---------------------------------------------------------------- contenido

        private static void EscribirEncabezadoReporte(Excel.Worksheet hoja)
        {
            Excel.Range titulo = Rango(hoja, "A1:H1");
            titulo.Merge();
            titulo.Value2 = "Reporte de Ventas Trimestral";

            Excel.Range subtitulo = Rango(hoja, "A2:H2");
            subtitulo.Merge();
            subtitulo.Value2 = "Periodo: 01/07/2026 - 30/09/2026  |  Generado el "
                               + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + " mediante COM Interop";
        }

        private static void EscribirTabla(Excel.Worksheet hoja)
        {
            // Encabezados de la tabla (fila 4)
            Rango(hoja, "A4:H4").Value2 = new object[,]
            {
                { "ID_Prod", "Nom_Producto", "Unidades existentes", "Unidades vendidas",
                  "Fecha Venta", "Precio Unitario", "Total Ventas", "Valor del Inventario" }
            };

            // 10 filas de datos (columnas A-F); G y H son fórmulas.
            object[,] datos =
            {
                { "P-001", "Laptop 15\"",        25,  18, new DateTime(2026, 7, 3),  14500.00 },
                { "P-002", "Mouse inalámbrico", 150, 120, new DateTime(2026, 7, 14),   250.00 },
                { "P-003", "Teclado mecánico",   80,  45, new DateTime(2026, 7, 29),  1200.00 },
                { "P-004", "Monitor 24\"",       40,  32, new DateTime(2026, 8, 5),   3800.00 },
                { "P-005", "Impresora láser",    15,   9, new DateTime(2026, 8, 18),  5600.00 },
                { "P-006", "Disco SSD 1 TB",     60,  51, new DateTime(2026, 8, 27),  1750.00 },
                { "P-007", "Memoria USB 64 GB", 200, 175, new DateTime(2026, 9, 2),    180.00 },
                { "P-008", "Webcam HD",          55,  38, new DateTime(2026, 9, 11),   650.00 },
                { "P-009", "Audífonos Bluetooth",90,  66, new DateTime(2026, 9, 19),   890.00 },
                { "P-010", "Router Wi-Fi 6",     35,  27, new DateTime(2026, 9, 28),  2100.00 },
            };

            // Las fechas se escriben como número de serie OLE; el formato visual se aplica después.
            int filas = datos.GetLength(0);
            for (int i = 0; i < filas; i++)
                datos[i, 4] = ((DateTime)datos[i, 4]).ToOADate();

            Rango(hoja, string.Format("A{0}:F{1}", PrimeraFilaDatos, UltimaFilaDatos)).Value2 = datos;
        }

        private static void EscribirFormulas(Excel.Worksheet hoja)
        {
            // Se usa Formula (sintaxis en inglés, independiente de la configuración regional).
            // Al asignar una fórmula relativa a un rango, Excel la ajusta para cada fila.
            Rango(hoja, string.Format("G{0}:G{1}", PrimeraFilaDatos, UltimaFilaDatos))
                .Formula = "=D" + PrimeraFilaDatos + "*F" + PrimeraFilaDatos;      // Unidades vendidas * Precio
            Rango(hoja, string.Format("H{0}:H{1}", PrimeraFilaDatos, UltimaFilaDatos))
                .Formula = "=C" + PrimeraFilaDatos + "*F" + PrimeraFilaDatos;      // Unidades existentes * Precio

            // Fila de totales
            Rango(hoja, "A" + FilaTotales + ":B" + FilaTotales).Merge();
            Rango(hoja, "A" + FilaTotales).Value2 = "TOTAL";
            foreach (string col in new[] { "C", "D", "G", "H" })
            {
                Rango(hoja, col + FilaTotales).Formula =
                    string.Format("=SUM({0}{1}:{0}{2})", col, PrimeraFilaDatos, UltimaFilaDatos);
            }
        }

        // ---------------------------------------------------------------- formato

        private static void AplicarFormato(Excel.Worksheet hoja)
        {
            // Título y subtítulo
            Excel.Range titulo = Rango(hoja, "A1:H1");
            Fuente(titulo, negrita: true, tamano: 18, color: Rgb(31, 56, 100));
            titulo.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titulo.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            titulo.RowHeight = 32;

            Excel.Range subtitulo = Rango(hoja, "A2:H2");
            Fuente(subtitulo, cursiva: true, tamano: 10, color: Rgb(89, 89, 89));
            subtitulo.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Encabezados de la tabla: negritas, texto blanco y fondo azul
            Excel.Range encabezados = Rango(hoja, "A4:H4");
            Fuente(encabezados, negrita: true, color: Rgb(255, 255, 255));
            Fondo(encabezados, Rgb(31, 78, 121));
            encabezados.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            encabezados.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            encabezados.WrapText = true;
            encabezados.RowHeight = 32;

            // Fila de totales
            Excel.Range totales = Rango(hoja, "A" + FilaTotales + ":H" + FilaTotales);
            Fuente(totales, negrita: true);
            Fondo(totales, Rgb(221, 235, 247));
            Rango(hoja, "A" + FilaTotales).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Bordes en toda la tabla (encabezado + datos + totales)
            Excel.Range tabla = Rango(hoja, "A4:H" + FilaTotales);
            Excel.Borders bordes = Registrar(tabla.Borders);
            bordes.LineStyle = Excel.XlLineStyle.xlContinuous;
            bordes.Weight = Excel.XlBorderWeight.xlThin;

            // Formatos numéricos (códigos en sintaxis invariante)
            const string moneda = "$#,##0.00";
            Rango(hoja, string.Format("F{0}:H{1}", PrimeraFilaDatos, FilaTotales)).NumberFormat = moneda;
            Rango(hoja, string.Format("C{0}:D{1}", PrimeraFilaDatos, FilaTotales)).NumberFormat = "#,##0";
            Rango(hoja, string.Format("E{0}:E{1}", PrimeraFilaDatos, UltimaFilaDatos)).NumberFormat = "dd/mm/yyyy";

            // Alineación de columnas de texto/fecha
            Rango(hoja, string.Format("A{0}:A{1}", PrimeraFilaDatos, UltimaFilaDatos))
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Rango(hoja, string.Format("E{0}:E{1}", PrimeraFilaDatos, UltimaFilaDatos))
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Ancho de columnas
            double[] anchos = { 10, 24, 14, 14, 14, 16, 16, 18 };
            for (int i = 0; i < anchos.Length; i++)
            {
                string letra = ((char)('A' + i)).ToString();
                Rango(hoja, letra + ":" + letra).ColumnWidth = anchos[i];
            }
        }

        private static void Fuente(Excel.Range rango, bool negrita = false, bool cursiva = false,
                                   int tamano = 0, int color = -1)
        {
            Excel.Font fuente = Registrar(rango.Font);
            if (negrita) fuente.Bold = true;
            if (cursiva) fuente.Italic = true;
            if (tamano > 0) fuente.Size = tamano;
            if (color >= 0) fuente.Color = color;
        }

        private static void Fondo(Excel.Range rango, int color)
        {
            Excel.Interior interior = Registrar(rango.Interior);
            interior.Color = color;
        }

        // El color en Excel es BGR: R + G*256 + B*65536
        private static int Rgb(int r, int g, int b)
        {
            return r + (g << 8) + (b << 16);
        }

        // ---------------------------------------------------------------- gráfico

        private static void CrearGrafico(Excel.Worksheet hoja)
        {
            // El gráfico ocupa el área A18:H38, debajo de la tabla
            Excel.Range zona = Rango(hoja, "A18:H38");
            double izquierda = (double)zona.Left;
            double arriba = (double)zona.Top;
            double ancho = (double)zona.Width;
            double alto = (double)zona.Height;

            Excel.ChartObjects objetos = Registrar((Excel.ChartObjects)hoja.ChartObjects());
            Excel.ChartObject objetoGrafico = Registrar(objetos.Add(izquierda, arriba, ancho, alto));
            objetoGrafico.Name = "GraficoVentas";

            Excel.Chart grafico = Registrar(objetoGrafico.Chart);
            grafico.ChartType = Excel.XlChartType.xlColumnClustered;

            // Las series se enlazan directamente a los rangos de la tabla: si cambian los
            // datos (o las fórmulas), el gráfico se actualiza solo.
            Excel.SeriesCollection series = Registrar((Excel.SeriesCollection)grafico.SeriesCollection());

            Excel.Series ventas = Registrar(series.NewSeries());
            ventas.Name = "=" + NombreHoja + "!$G$4";
            ventas.Values = Rango(hoja, string.Format("G{0}:G{1}", PrimeraFilaDatos, UltimaFilaDatos));
            ventas.XValues = Rango(hoja, string.Format("B{0}:B{1}", PrimeraFilaDatos, UltimaFilaDatos));

            Excel.Series inventario = Registrar(series.NewSeries());
            inventario.Name = "=" + NombreHoja + "!$H$4";
            inventario.Values = Rango(hoja, string.Format("H{0}:H{1}", PrimeraFilaDatos, UltimaFilaDatos));
            inventario.XValues = Rango(hoja, string.Format("B{0}:B{1}", PrimeraFilaDatos, UltimaFilaDatos));

            grafico.HasTitle = true;
            Excel.ChartTitle titulo = Registrar(grafico.ChartTitle);
            titulo.Text = "Total de ventas y valor del inventario por producto";

            grafico.HasLegend = true;
            Excel.Legend leyenda = Registrar(grafico.Legend);
            leyenda.Position = Excel.XlLegendPosition.xlLegendPositionBottom;

            Excel.Axis ejeX = Registrar((Excel.Axis)grafico.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary));
            ejeX.HasTitle = true;
            Excel.AxisTitle tituloX = Registrar(ejeX.AxisTitle);
            tituloX.Text = "Producto";

            Excel.Axis ejeY = Registrar((Excel.Axis)grafico.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary));
            ejeY.HasTitle = true;
            Excel.AxisTitle tituloY = Registrar(ejeY.AxisTitle);
            tituloY.Text = "Importe ($)";
        }

        // ---------------------------------------------------------------- gestión COM

        private static Excel.Range Rango(Excel.Worksheet hoja, string direccion)
        {
            return Registrar(hoja.Range[direccion]);
        }

        private static T Registrar<T>(T objetoCom) where T : class
        {
            if (objetoCom != null) ReferenciasCom.Add(objetoCom);
            return objetoCom;
        }

        private static void CerrarExcel(Excel.Application excel)
        {
            if (excel == null) return;
            try
            {
                excel.Quit();
            }
            catch (COMException ex)
            {
                Console.Error.WriteLine("No se pudo cerrar Excel limpiamente: " + ex.Message);
            }
        }

        // Libera en orden inverso al de adquisición (hijos antes que padres).
        private static int LiberarReferencias()
        {
            int liberadas = 0;
            for (int i = ReferenciasCom.Count - 1; i >= 0; i--)
            {
                try
                {
                    // Un mismo RCW puede haberse registrado varias veces; se repite hasta 0.
                    while (Marshal.ReleaseComObject(ReferenciasCom[i]) > 0) { }
                    liberadas++;
                }
                catch (InvalidComObjectException)
                {
                    // Ya liberado por una llamada anterior sobre el mismo RCW.
                }
                catch (ArgumentException)
                {
                    // El objeto no era un objeto COM.
                }
                ReferenciasCom[i] = null;
            }
            ReferenciasCom.Clear();
            return liberadas;
        }

        private static bool EsperarCierreProceso(uint pid, int milisegundos)
        {
            try
            {
                using (Process proceso = Process.GetProcessById((int)pid))
                {
                    return proceso.WaitForExit(milisegundos);
                }
            }
            catch (ArgumentException)
            {
                return true; // el proceso ya no existe
            }
        }
    }
}
