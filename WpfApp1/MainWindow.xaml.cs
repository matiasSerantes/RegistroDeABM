using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using ExcelDataReader;
using ExcelDataReader.Exceptions;
using ExcelDataReader.Log;
using ClosedXML.Excel; // Asegúrate de tener esta directiva 'using' para ClosedXML
using OfficeOpenXml;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        private string originalFilePath = @"C:\users\matias.serantes\OneDrive - Toyota Argentina S.A\registro ticket.xlsx";
        private DataTable originalDataTable; // Almacena los datos originales del Excel
        private System.Timers.Timer timer;
        private string temporalFilePath = Path.Combine(Path.GetTempPath(), "registro_ticket_temp.xlsx");

        public MainWindow()
        {
            InitializeComponent();
            // Llama a la función para cargar los datos de Excel en el DataGrid
            LoadExcelData();
            
        }

        private void excelDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void LoadExcelData()
        {
            try
            {
                // Carga del archivo Excel usando ExcelDataReader
                using (var stream = File.Open(originalFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Configura el lector para usar la primera fila como cabecera
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        // Obtener la tabla del DataSet
                        originalDataTable = result.Tables[0];
                    }
                }
                // Eliminar la primera columna del DataTable
                if (originalDataTable.Columns.Count > 0)
                {
                    originalDataTable.Columns.RemoveAt(0);
                }

                // Asignar el DataTable original al DataGrid
                excelDataGrid.ItemsSource = originalDataTable.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar el archivo Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                // Copia el archivo original a la ubicación temporal
                File.Copy(originalFilePath, temporalFilePath, true);

                // Recarga los datos en el DataGrid
                Dispatcher.Invoke(() => LoadExcelData());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al verificar la modificación del archivo Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void Buscar_Click(object sender, RoutedEventArgs e)
        {
            // Verifica si todos los campos de búsqueda están vacíos
            if (string.IsNullOrWhiteSpace(TicketBox.Text) &&
                string.IsNullOrWhiteSpace(ActivoBox.Text) &&
                string.IsNullOrWhiteSpace(UsuarioBox.Text) &&
                string.IsNullOrWhiteSpace(MarcaBox.Text) &&
                string.IsNullOrWhiteSpace(ModeloBox.Text) &&
                string.IsNullOrWhiteSpace(SerieBox.Text) &&
                string.IsNullOrWhiteSpace(FechaBox.Text) &&
                string.IsNullOrWhiteSpace(ABMBox.Text))
            {
                // Si todos los campos están vacíos, carga todos los datos originales
                LoadExcelData();
                return;
            }

            // Crea un filtro basado en los campos que contienen información
            string filterExpression = "";

            if (!string.IsNullOrWhiteSpace(TicketBox.Text))
            {
                filterExpression += $"CONVERT([Numero De Ticket], 'System.String') LIKE '%{TicketBox.Text}%' AND ";
            }

            if (!string.IsNullOrWhiteSpace(ActivoBox.Text))
            {
                filterExpression += $"CONVERT([Numero De Activo], 'System.String') LIKE '%{ActivoBox.Text}%' AND ";
            }

            if (!string.IsNullOrWhiteSpace(UsuarioBox.Text))
            {
                filterExpression += $"[Usuario] LIKE '%{UsuarioBox.Text}%' AND ";
            }

            if (!string.IsNullOrWhiteSpace(MarcaBox.Text))
            {
                filterExpression += $"[Marca] LIKE '%{MarcaBox.Text}%' AND ";
            }

            if (!string.IsNullOrWhiteSpace(ModeloBox.Text))
            {
                filterExpression += $"CONVERT([Modelo], 'System.String') LIKE '%{ModeloBox.Text}%' AND ";
            }

            if (!string.IsNullOrWhiteSpace(SerieBox.Text))
            {
                filterExpression += $"[IME o Serial] LIKE '%{SerieBox.Text}%' AND ";
            }

            if (!string.IsNullOrWhiteSpace(FechaBox.Text))
            {
                filterExpression += $"[Fecha de Entrega] = '{FechaBox.Text}' AND ";
            }

            if (!string.IsNullOrWhiteSpace(ABMBox.Text))
            {
                filterExpression += $"[ABM] LIKE '%{ABMBox.Text}%' AND ";
            }

            // Elimina el "AND" adicional al final del filtro
            if (!string.IsNullOrWhiteSpace(filterExpression))
            {
                filterExpression = filterExpression.Substring(0, filterExpression.Length - 5);
            }

            // Aplica el filtro al DataTable original y actualiza el DataGrid
            if (!string.IsNullOrWhiteSpace(filterExpression))
            {
                DataView dv = originalDataTable.DefaultView;
                dv.RowFilter = filterExpression;
                excelDataGrid.ItemsSource = dv;
            }
            else
            {
                // Si no se ingresaron criterios de búsqueda, se restauran todos los datos originales
                LoadExcelData();
            }
        }


        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Eliminar el archivo temporal si existe al cerrar la aplicación
            if (File.Exists(temporalFilePath))
            {
                try
                {
                    File.Delete(temporalFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al eliminar el archivo temporal: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        private void Actualizar_Click(object sender, RoutedEventArgs e)
        {
            LoadExcelData();
        }
        private void Vaciar_Click(object sender, RoutedEventArgs e)
        {
            // Establecer el contenido de todos los TextBox como una cadena vacía
            TicketBox.Text = string.Empty;
            ActivoBox.Text = string.Empty;
            UsuarioBox.Text = string.Empty;
            MarcaBox.Text = string.Empty;
            ModeloBox.Text = string.Empty;
            SerieBox.Text = string.Empty;
            FechaBox.Text = string.Empty;
            ABMBox.Text = string.Empty;
            LoadExcelData();
        }
        // Función para verificar si una fila está vacía
        private bool EsFilaVacia(IXLRow row)
        {
            foreach (var cell in row.CellsUsed())
            {
                if (!string.IsNullOrWhiteSpace(cell.Value.ToString()))
                {
                    return false;
                }
            }
            return true;
        }

        private void Guardar_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Mostrar mensaje mientras se guardan los datos
                MessageBox.Show("Guardando los datos, por favor espere...", "Guardando", MessageBoxButton.OK, MessageBoxImage.Information);

                if (!File.Exists(originalFilePath))
                {
                    MessageBox.Show($"El archivo Excel especificado no existe en la ruta: {originalFilePath}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                using (var workbook = new XLWorkbook(originalFilePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    int lastUsedRow = worksheet.LastRowUsed().RowNumber();

                    // Copiar los estilos de la fila anterior a la nueva fila
                    var newRow = worksheet.Row(lastUsedRow + 1);
                    var previousRow = worksheet.Row(lastUsedRow);

                    foreach (var cell in previousRow.CellsUsed())
                    {
                        var newCell = newRow.Cell(cell.Address.ColumnNumber);
                        newCell.Style = cell.Style; // Copiar el estilo de la celda
                    }

                    // Asignar los valores a la nueva fila respetando los formatos de las celdas originales
                    worksheet.Cell(lastUsedRow + 1, 2).SetValue(TicketBox.Text); // Columna B
                    worksheet.Cell(lastUsedRow + 1, 3).SetValue(ActivoBox.Text); // Columna C
                    worksheet.Cell(lastUsedRow + 1, 4).SetValue(UsuarioBox.Text); // Columna D
                    worksheet.Cell(lastUsedRow + 1, 5).SetValue(MarcaBox.Text);  // Columna E
                    worksheet.Cell(lastUsedRow + 1, 6).SetValue(ModeloBox.Text);  // Columna F
                    worksheet.Cell(lastUsedRow + 1, 7).SetValue(SerieBox.Text);  // Columna G
                    worksheet.Cell(lastUsedRow + 1, 8).SetValue(FechaBox.Text);  // Columna H
                    worksheet.Cell(lastUsedRow + 1, 9).SetValue(ABMBox.Text);    // Columna I

                    workbook.Save();
                }

                // Mostrar mensaje de éxito
                MessageBox.Show("Datos guardados exitosamente en el archivo Excel.", "Éxito", MessageBoxButton.OK, MessageBoxImage.Information);

                // Actualizar el DataGrid para reflejar los nuevos datos
                LoadExcelData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar los datos en el archivo Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
