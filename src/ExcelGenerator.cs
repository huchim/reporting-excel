// -----------------------------------------------------------------------
// <copyright file="ExcelGenerator.cs" company="Carlos Huchim Ahumada">
// Este código se libera bajo los términos de licencia especificados.
// </copyright>
// -----------------------------------------------------------------------
namespace Jaguar.Reporting.Generators
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using Jaguar.Reporting;
    using Jaguar.Reporting.Common;
    using OfficeOpenXml;

    /// <summary>
    /// Genera la información en formato CSV.
    /// </summary>
    public class ExcelGenerator : IGeneratorEngine
    {
        private ReportHandler report;
        private Dictionary<string, object> variables;

        /// <inheritdoc/>
        public string FileExtension => ".xlsx";

        /// <inheritdoc/>
        public Guid Id => new Guid("c8df5fce-0681-40e6-9ec9-e651f3669a47");

        /// <inheritdoc/>
        public string MimeType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        /// <inheritdoc/>
        public string Name => "Libro de excel";

        /// <inheritdoc/>
        public bool IsEmbed => false;

        private bool IsTemplating
        {
            get
            {
                if (!this.report.Options.ContainsKey("excel.template"))
                {
                    return false;
                }

                if (string.IsNullOrEmpty(this.report.Options["excel.template"].ToString()))
                {
                    return false;
                }

                return true;
            }
        }

        private bool EnableHeader
        {
            get
            {
                if (!this.report.Options.ContainsKey("excel.defaults.header"))
                {
                    return true;
                }

                if (string.IsNullOrEmpty(this.report.Options["excel.defaults.header"].ToString()))
                {
                    return true;
                }

                if (this.report.Options["excel.defaults.header"] is bool)
                {
                    return (bool)this.report.Options["excel.defaults.header"];
                }

                return true;
            }
        }

        private bool IsReadOnly
        {
            get
            {
                if (!this.report.Options.ContainsKey("excel.defaults.protect"))
                {
                    return false;
                }

                if (string.IsNullOrEmpty(this.report.Options["excel.defaults.protect"].ToString()))
                {
                    return false;
                }

                if (this.report.Options["excel.defaults.protect"] is bool)
                {
                    return (bool)this.report.Options["excel.defaults.protect"];
                }

                return false;
            }
        }

        private string DefaultSheetName
        {
            get
            {
                if (!this.report.Options.ContainsKey("excel.defaults.sheetname"))
                {
                    return "Hoja";
                }

                return this.report.Options["excel.defaults.sheetname"].ToString();
            }
        }

        private string TemplateFile
        {
            get
            {
                if (this.IsTemplating && string.IsNullOrEmpty(this.report.Options["excel.template"] as string))
                {
                    throw new ArgumentNullException(nameof(this.TemplateFile));
                }

                return Path.Combine(this.report.WorkDirectory, this.report.Options["excel.template"].ToString());
            }
        }

        /// <inheritdoc/>
        public string GetString(ReportHandler report, List<DataTable> data, Dictionary<string, object> variables)
        {
            throw new NotImplementedException();
        }

        /// <inheritdoc/>
        public byte[] GetAllBytes(ReportHandler report, List<DataTable> data, Dictionary<string, object> variables)
        {
            this.variables = variables;
            this.report = report;

            var results = this.IsTemplating ? this.CreateCustomFile(data) : this.CreateDefaultFile(data);

            return results;
        }

        private void ReplaceWorkBookVariables(ExcelWorkbook wb)
        {
            // Acceder a las celdas que requieren se sustituídas.
            var cells = wb.Names.Where(x => x.Name.StartsWith("_") && x.Columns == 1 && x.Rows == 1);

            foreach (var cell in cells)
            {
                if (cell.Value == null)
                {
                    break;
                }

                var value = cell.Value.ToString();

                foreach (var k in this.variables)
                {
                    value = value.Replace($"%{k.Key}%".ToString(), k.Value.ToString());
                }

                cell.Value = value;
            }
        }

        private void ReplaceDataVariables(ExcelWorkbook wb, List<DataTable> data)
        {
            // Acceder a las celdas que requieren se sustituídas.
            var cells = wb.Names.Where(x => x.Name.Contains(".") && x.Columns == 1 && x.Rows == 1);

            foreach (var cell in cells)
            {
                var cellName = cell.Name.Split(".".ToArray(), StringSplitOptions.RemoveEmptyEntries);

                if (cellName.Count() != 2)
                {
                    // Si el nombre no corresponde a una referencia a la tabla y columna, lo salto.
                    break;
                }

                var tableName = cellName[0];
                var columnName = cellName[1];

                // Buscar la tabla que tenga la columna buscada.
                var table = data.FirstOrDefault(x => x.TableName == tableName && x.Columns.Any(c => c.Name == columnName));

                if (table == null)
                {
                    // La tabla no existe. ¿Debo devolver un error?
                    cell.Value = $"##{tableName}.{columnName}##NotFound";
                    break;
                }

                if (table.HasRows)
                {
                    // Tomar el valor de la columna requerida del primer registro encontrado.
                    cell.Value = table.Rows.First().Columns.Single(x => x.Name == columnName).Value;
                }
                else
                {
                    // No se encontró registros en la tabla.
                    cell.Value = null;
                }
            }
        }

        private string ParseColumnName(string columnName)
        {
            if (!this.report.Options.ContainsKey($"excel.headings.{columnName}"))
            {
                return columnName;
            }

            return this.report.Options[$"excel.headings.{columnName}"] as string;
        }

        private void ConfigureWorkBookMetada(ExcelWorkbook wb)
        {
            wb.Properties.Title = this.report.Label;
            wb.Properties.Subject = this.report.Description;
            wb.Properties.Comments = $"Reporte v{this.report.Version}. Privado: {this.report.IsPrivate}";
            wb.Properties.Author = string.Join(", ", this.report.Authors);
            wb.Properties.Keywords = string.Join(",", this.report.Keywords);
    }

        private byte[] CreateCustomFile(List<DataTable> data)
        {
            var templateFile = new FileInfo(Path.Combine(this.report.WorkDirectory, this.TemplateFile));

            if (!templateFile.Exists)
            {
                throw new FileNotFoundException("No se pudo encontrar la plantilla que se desea usar.", templateFile.FullName);
            }

            using (ExcelPackage p = new ExcelPackage(templateFile, true))
            {
                // Metadatos del archivo.
                this.ConfigureWorkBookMetada(p.Workbook);

                var calcMode = p.Workbook.CalcMode;

                // Deshabilitar temporalmente el cálculo de las formulas.
                if (p.Workbook.CalcMode != ExcelCalcMode.Manual)
                {
                    p.Workbook.CalcMode = ExcelCalcMode.Manual;
                }

                // Se puede asignar un nombre a una celda y agregar en la celda
                // una sustitución por variable o el valor de una columna en el
                // primer registro de la tabla.
                // Ejemplo: _GENERATED = %system.now%
                //          data.No_Cheque = x;
                this.ReplaceWorkBookVariables(p.Workbook);
                this.ReplaceDataVariables(p.Workbook, data);

                var listOfTables = p.Workbook.Worksheets[1].Tables;

                foreach (var tableInfo in listOfTables)
                {
                    // Buscar si la tabla existe entre los datos.
                    var currentData = data.FirstOrDefault(x => x.TableName == tableInfo.Name);
                    var sheet = tableInfo.WorkSheet;
                    var tableBeginning = tableInfo.Address.Start;
                    var rowIndex = (tableInfo.ShowHeader ? 1 : 0) + tableBeginning.Row;
                    var colIndex = tableBeginning.Column;
                    var maxColIndex = tableInfo.Address.End.Column;
                    var templateRowCount = tableInfo.Address.End.Row - tableInfo.Address.Start.Row;

                    // Verificar protección de la hoja.
                    if (this.IsReadOnly)
                    {
                        sheet.Protection.AllowSort = true;
                        sheet.Protection.AllowFormatRows = true;
                        sheet.Protection.AllowFormatCells = true;
                        sheet.Protection.AllowFormatColumns = true;
                        sheet.Protection.AllowAutoFilter = true;
                        sheet.Protection.IsProtected = true;
                        sheet.Protection.SetPassword("HuchimIsAlive");
                    }

                    if (currentData != null)
                    {
                        // Obtener la lista de columnas que son requeridas.
                        // Se omiten aquellas que comiencen con un guión bajo.
                        var requiredColumnList = tableInfo.Columns.Where(x => !x.Name.StartsWith("_")).Select(x => x.Name).ToArray();

                        // Validar que las columnas existan en la tabla de datos.
                        var allRequiredColumnsExists = currentData.Columns.Count(x => requiredColumnList.Contains(x.Name)) >= requiredColumnList.Count();

                        if (!allRequiredColumnsExists)
                        {
                            sheet.SetValue(rowIndex, colIndex, "Se requieren campos que no se encuentran en la consulta. Use un guión bajo como prefijo en las columnas que no sean obligatorias.");
                            break;
                        }

                        if (currentData.Rows.Count != 0)
                        {
                            foreach (var currentRow in currentData.Rows)
                            {
                                // Agregar una fila para este registro.
                                sheet.InsertRow(rowIndex, 1, rowIndex + 1);

                                // Copiar el contenido de la fila de referencia que quedaría abajo.
                                // sheet.Cells[rowIndex + 1, colIndex, rowIndex + 1, maxColIndex].Copy(sheet.Cells[rowIndex, colIndex, rowIndex, maxColIndex]);
                                // workSheet.Cells[1, 1, 1, maxColumnIndex].Copy(workSheet.Cells[i * totalRows + 1, 1]);

                                // Agregar los datos a la fila.
                                for (var tableColumnIndex = 0; tableColumnIndex < tableInfo.Columns.Count; tableColumnIndex++)
                                {
                                    // Recuperar el nombre de la columna.
                                    var columnName = tableInfo.Columns[tableColumnIndex].Name;

                                    // Incluir únicamente las columnas que no inicien con un guión bajo.
                                    var validColumnName = !columnName.StartsWith("_");

                                    if (validColumnName)
                                    {
                                        // Recuperar la información de la columna.
                                        var columnInfo = currentRow.Columns.Single(x => x.Name == columnName);

                                        sheet.SetValue(rowIndex, colIndex + tableColumnIndex, columnInfo.Value);
                                    }
                                }

                                rowIndex++;
                            }
                        }

                        // Eliminar todas las filas que son de la plantilla y no corresponden a los datos.
                        for (var c = 1; c <= templateRowCount; c++)
                        {
                            sheet.DeleteRow(rowIndex);
                        }
                    }

                    // Actualizar el nombre de las columnas en caso de ser necesario.
                    if (tableInfo.ShowHeader)
                    {
                        // Asignar el nombre que se desea para la columna.
                        for (var i = colIndex; i <= maxColIndex; i++)
                        {
                            var columnName = sheet.Cells[tableBeginning.Row, i].Value.ToString();
                            var alias = this.ParseColumnName(columnName);

                            if (!string.IsNullOrEmpty(alias) && alias != columnName)
                            {
                                sheet.Cells[tableBeginning.Row, i].Value = alias;
                            }
                        }
                    }
                }

                p.Workbook.CalcMode = calcMode;

                return p.GetAsByteArray();
            }
        }

        private string GetNewSheetName(ExcelWorksheets sheets)
        {
            if (sheets.Count == 0)
            {
                return this.DefaultSheetName;
            }

            return $"{this.DefaultSheetName}{sheets.Count + 1}";
        }

        private byte[] CreateDefaultFile(List<DataTable> data)
        {
            using (ExcelPackage p = new ExcelPackage())
            {
                // Metadatos del archivo.
                this.ConfigureWorkBookMetada(p.Workbook);

                foreach (var currentData in data)
                {
                    // Crear una hoja en el libro nuevo.
                    var sheet = p.Workbook.Worksheets.Add(this.GetNewSheetName(p.Workbook.Worksheets));
                    var colIndex = 1;
                    var rowIndex = 1;

                    // Generar las columnas para los datos.
                    if (currentData.Columns.Count != 0 && this.EnableHeader)
                    {
                        foreach (var c in currentData.Columns)
                        {
                            // Asignar el nombre de la columna.
                            sheet.SetValue(rowIndex, colIndex, c.Name);

                            // Dar formato a toda la columna.
                            if (c.Type == typeof(DateTime))
                            {
                                // ShortDatePattern no es del todo compatible con NumberFormat.
                                // Referencia: http://stackoverflow.com/questions/32459905/is-excel-epplus-number-format-compatible-with-datetimeformat-shortdatepattern
                                sheet.Column(colIndex).Style.Numberformat.Format = "yyyy-MM-dd";
                            }

                            colIndex++;
                        }

                        // Poner en negritas el encabezado.
                        sheet.Row(rowIndex).Style.Font.Bold = true;

                        // Establecer los valores de los contadores.
                        colIndex = 1;
                        rowIndex++;
                    }

                    if (currentData.Rows.Count != 0)
                    {
                        foreach (var c in currentData.Rows)
                        {
                            foreach (var m in c.Columns)
                            {
                                sheet.SetValue(rowIndex, colIndex, m.Value);
                                colIndex++;
                            }

                            colIndex = 1;
                            rowIndex++;
                        }
                    }
                }

                return p.GetAsByteArray();
            }
        }
    }
}