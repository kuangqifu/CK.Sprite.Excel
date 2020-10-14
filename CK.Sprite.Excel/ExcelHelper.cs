using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace CK.Sprite.Excel
{
    public class ExcelHelper
    {
        public static string Company { get; set; } = "东宸科技";
        public static string ApplicationName { get; set; } = "内部管理系统";
        public static string Author { get; set; } = "kuangqifu";

        #region 导出

        /// <summary>
        /// 导出list数据到Excel
        /// </summary>
        /// <typeparam name="T">实体</typeparam>
        /// <param name="exportDatas">导出的list数据</param>
        /// <param name="templateModels">Excel模版信息</param>
        /// <param name="title">Sheet名称</param>
        /// <param name="multiHeaderInfos">多表头定义信息</param>
        /// <returns></returns>
        public static HSSFWorkbook Export<T>(List<T> exportDatas, List<ExcelTemplate> templateModels, string title, List<List<MultiHeaderInfo>> multiHeaderInfos = null)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            workbook.SetSheetName(0, title);

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = Company;
            workbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Author = Author;
            si.ApplicationName = ApplicationName;
            si.Title = title;

            si.CreateDateTime = DateTime.Now;
            workbook.SummaryInformation = si;

            //取得列宽
            int[] arrColWidth = new int[templateModels.Count];
            int columnIndex = 0;
            foreach (var templateModel in templateModels)
            {
                arrColWidth[columnIndex] = templateModel.CellLength > 0 ? templateModel.CellLength * 2 : Encoding.UTF8.GetBytes(templateModel.Name.ToString()).Length;
                columnIndex++;
            }

            int rowIndex = 0;
            foreach (var exportData in exportDatas)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = workbook.CreateSheet();
                    }

                    if (multiHeaderInfos != null && multiHeaderInfos.Count > 0) // 复杂表头合并等
                    {
                        List<int>[] usedCellIndexs = new List<int>[multiHeaderInfos.Count];
                        for (var i = 0; i < multiHeaderInfos.Count; i++)
                        {
                            usedCellIndexs[i] = new List<int>();
                        }
                        for (var i = 0; i < multiHeaderInfos.Count; i++)
                        {
                            var colIndex = 0;
                            var headerRow = sheet.CreateRow(i);
                            var headStyle = workbook.CreateCellStyle();
                            headStyle.Alignment = HorizontalAlignment.Center;
                            headStyle.VerticalAlignment = VerticalAlignment.Center;
                            var font = workbook.CreateFont();
                            font.FontHeightInPoints = 10;
                            font.IsBold = true;
                            headStyle.SetFont(font);
                            foreach (var multiHeaderInfo in multiHeaderInfos[i])
                            {
                                while (true) // 找未使用的第一个单元格
                                {
                                    if (!usedCellIndexs[i].Contains(colIndex))
                                    {
                                        break;
                                    }
                                    colIndex++;
                                }
                                headerRow.CreateCell(colIndex).SetCellValue(multiHeaderInfo.Name);
                                var oldColIndex = colIndex;
                                if (multiHeaderInfo.ColSpan > 1 || multiHeaderInfo.RowSpan > 1)
                                {
                                    sheet.AddMergedRegion(new CellRangeAddress(i, i + multiHeaderInfo.RowSpan - 1, colIndex, colIndex + multiHeaderInfo.ColSpan - 1));
                                    if (multiHeaderInfo.RowSpan > 1)
                                    {
                                        for (var j = 1; j < multiHeaderInfo.RowSpan; j++)
                                        {
                                            for (var k = colIndex; k < colIndex + multiHeaderInfo.ColSpan; k++)
                                            {
                                                usedCellIndexs[i + j].Add(k);
                                            }
                                        }
                                    }
                                    colIndex = colIndex + multiHeaderInfo.ColSpan;
                                }
                                else
                                {
                                    colIndex++;
                                }
                                headerRow.GetCell(oldColIndex).CellStyle = headStyle;
                            }
                        }
                        rowIndex = multiHeaderInfos.Count;
                    }
                    else
                    {
                        #region 列头及样式
                        {
                            var headerRow = sheet.CreateRow(0);
                            var headStyle = workbook.CreateCellStyle();
                            headStyle.Alignment = HorizontalAlignment.Center;
                            var font = workbook.CreateFont();
                            font.FontHeightInPoints = 10;
                            font.IsBold = true;
                            headStyle.SetFont(font);
                            columnIndex = 0;
                            foreach (var templateModel in templateModels)
                            {
                                headerRow.CreateCell(columnIndex).SetCellValue(templateModel.Name);
                                headerRow.GetCell(columnIndex).CellStyle = headStyle;

                                //设置列宽
                                sheet.SetColumnWidth(columnIndex, (arrColWidth[columnIndex] + 1) * 256);
                                columnIndex++;
                            }
                        }
                        #endregion

                        rowIndex = 1;
                    }
                }


                #endregion


                #region 填充内容
                var dataRow = sheet.CreateRow(rowIndex);
                columnIndex = 0;
                foreach (var templateModel in templateModels)
                {
                    var newCell = dataRow.CreateCell(columnIndex);

                    var objValue = typeof(T).GetProperty(templateModel.Field).GetValue(exportData)?.ToString();

                    switch (templateModel.FieldType)
                    {
                        case EFieldType.Int:
                            int intV = 0;
                            int.TryParse(objValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case EFieldType.Double:
                            double doubV = 0;
                            double.TryParse(objValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case EFieldType.Guid:
                            newCell.SetCellValue(objValue);
                            break;
                        case EFieldType.Bool:
                            bool boolV = false;
                            bool.TryParse(objValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case EFieldType.Date:
                            DateTime.TryParse(objValue, out var dateV);
                            if (dateV != DateTime.MinValue)
                            {
                                newCell.SetCellValue(dateV.ToString("yyyy-MM-dd"));
                            }
                            else
                            {
                                newCell.SetCellValue("");
                            }
                            break;
                        case EFieldType.DateTime:
                            DateTime.TryParse(objValue, out var datetimeV);
                            if (datetimeV != DateTime.MinValue)
                            {
                                newCell.SetCellValue(datetimeV.ToString("yyyy-MM-dd HH:mm:ss"));
                            }
                            else
                            {
                                newCell.SetCellValue("");
                            }
                            //newCell.SetCellType(CellType.String);
                            //newCell.SetCellValue(datetimeV.ToString("yyyy-MM-dd HH:mm:ss"));
                            break;
                        case EFieldType.String:
                            newCell.SetCellValue(objValue);
                            break;
                        default:
                            newCell.SetCellValue(objValue);
                            break;
                    }

                    columnIndex++;

                }
                #endregion

                rowIndex++;
            }

            for (int i = 0; i < templateModels.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            return workbook;
        }

        /// <summary>
        /// 导出Datatable数据到Excel
        /// </summary>
        /// <param name="exportTable">导出的Table信息</param>
        /// <param name="templateModels">Excel模版信息</param>
        /// <param name="title">Sheet名称</param>
        /// <param name="multiHeaderInfos">多表头定义信息</param>
        /// <returns></returns>
        public static HSSFWorkbook Export(DataTable exportTable, List<ExcelTemplate> templateModels, string title, List<List<MultiHeaderInfo>> multiHeaderInfos = null)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            workbook.SetSheetName(0, title);

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = Company;
            workbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Author = Author;
            si.ApplicationName = ApplicationName;
            si.Title = title;

            si.CreateDateTime = DateTime.Now;
            workbook.SummaryInformation = si;

            //取得列宽
            int[] arrColWidth = new int[templateModels.Count];
            int columnIndex = 0;
            foreach (var templateModel in templateModels)
            {
                arrColWidth[columnIndex] = templateModel.CellLength > 0 ? templateModel.CellLength * 2 : Encoding.UTF8.GetBytes(templateModel.Name.ToString()).Length;
                columnIndex++;
            }

            int rowIndex = 0;
            foreach (DataRow exportData in exportTable.Rows)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = workbook.CreateSheet();
                    }

                    if (multiHeaderInfos != null && multiHeaderInfos.Count > 0) // 复杂表头合并等
                    {
                        List<int>[] usedCellIndexs = new List<int>[multiHeaderInfos.Count];
                        for (var i = 0; i < multiHeaderInfos.Count; i++)
                        {
                            usedCellIndexs[i] = new List<int>();
                        }
                        for (var i = 0; i < multiHeaderInfos.Count; i++)
                        {
                            var colIndex = 0;
                            var headerRow = sheet.CreateRow(i);
                            var headStyle = workbook.CreateCellStyle();
                            headStyle.Alignment = HorizontalAlignment.Center;
                            headStyle.VerticalAlignment = VerticalAlignment.Center;
                            var font = workbook.CreateFont();
                            font.FontHeightInPoints = 10;
                            font.IsBold = true;
                            headStyle.SetFont(font);
                            foreach (var multiHeaderInfo in multiHeaderInfos[i])
                            {
                                while (true) // 找未使用的第一个单元格
                                {
                                    if (!usedCellIndexs[i].Contains(colIndex))
                                    {
                                        break;
                                    }
                                    colIndex++;
                                }
                                headerRow.CreateCell(colIndex).SetCellValue(multiHeaderInfo.Name);
                                var oldColIndex = colIndex;
                                if (multiHeaderInfo.ColSpan > 1 || multiHeaderInfo.RowSpan > 1)
                                {
                                    sheet.AddMergedRegion(new CellRangeAddress(i, i + multiHeaderInfo.RowSpan - 1, colIndex, colIndex + multiHeaderInfo.ColSpan - 1));
                                    if (multiHeaderInfo.RowSpan > 1)
                                    {
                                        for (var j = 1; j < multiHeaderInfo.RowSpan; j++)
                                        {
                                            for (var k = colIndex; k < colIndex + multiHeaderInfo.ColSpan; k++)
                                            {
                                                usedCellIndexs[i + j].Add(k);
                                            }
                                        }
                                    }
                                    colIndex = colIndex + multiHeaderInfo.ColSpan;
                                }
                                else
                                {
                                    colIndex++;
                                }
                                headerRow.GetCell(oldColIndex).CellStyle = headStyle;
                            }
                        }
                        rowIndex = multiHeaderInfos.Count;
                    }
                    else
                    {
                        #region 列头及样式
                        {
                            var headerRow = sheet.CreateRow(0);
                            var headStyle = workbook.CreateCellStyle();
                            headStyle.Alignment = HorizontalAlignment.Center;
                            var font = workbook.CreateFont();
                            font.FontHeightInPoints = 10;
                            font.IsBold = true;
                            headStyle.SetFont(font);
                            columnIndex = 0;
                            foreach (var templateModel in templateModels)
                            {
                                headerRow.CreateCell(columnIndex).SetCellValue(templateModel.Name);
                                headerRow.GetCell(columnIndex).CellStyle = headStyle;

                                //设置列宽
                                sheet.SetColumnWidth(columnIndex, (arrColWidth[columnIndex] + 1) * 256);
                                columnIndex++;
                            }
                        }
                        #endregion

                        rowIndex = 1;
                    }
                }


                #endregion


                #region 填充内容
                var dataRow = sheet.CreateRow(rowIndex);
                columnIndex = 0;
                foreach (var templateModel in templateModels)
                {
                    var newCell = dataRow.CreateCell(columnIndex);

                    var objValue = exportData[templateModel.Field]?.ToString();

                    switch (templateModel.FieldType)
                    {
                        case EFieldType.Int:
                            int intV = 0;
                            int.TryParse(objValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case EFieldType.Double:
                            double doubV = 0;
                            double.TryParse(objValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case EFieldType.Guid:
                            newCell.SetCellValue(objValue);
                            break;
                        case EFieldType.Bool:
                            bool boolV = false;
                            bool.TryParse(objValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case EFieldType.Date:
                            DateTime.TryParse(objValue, out var dateV);
                            if (dateV != DateTime.MinValue)
                            {
                                newCell.SetCellValue(dateV.ToString("yyyy-MM-dd"));
                            }
                            else
                            {
                                newCell.SetCellValue("");
                            }
                            break;
                        case EFieldType.DateTime:
                            DateTime.TryParse(objValue, out var datetimeV);
                            if (datetimeV != DateTime.MinValue)
                            {
                                newCell.SetCellValue(datetimeV.ToString("yyyy-MM-dd HH:mm:ss"));
                            }
                            else
                            {
                                newCell.SetCellValue("");
                            }
                            //newCell.SetCellType(CellType.String);
                            //newCell.SetCellValue(datetimeV.ToString("yyyy-MM-dd HH:mm:ss"));
                            break;
                        case EFieldType.String:
                            newCell.SetCellValue(objValue);
                            break;
                        default:
                            newCell.SetCellValue(objValue);
                            break;
                    }

                    columnIndex++;

                }
                #endregion

                rowIndex++;
            }

            for (int i = 0; i < templateModels.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            return workbook;
        }

        /// <summary>
        /// 直接导出Datatable数据到Excel
        /// </summary>
        /// <param name="exportTable">导出的Table信息</param>
        /// <param name="title">Sheet名称</param>
        /// <returns></returns>
        public static HSSFWorkbook ExportFromTable(DataTable exportTable, string title)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            workbook.SetSheetName(0, title);

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = Company;
            workbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Author = Author;
            si.ApplicationName = ApplicationName;
            si.Title = title;

            si.CreateDateTime = DateTime.Now;
            workbook.SummaryInformation = si;

            IRow headerRow = sheet.CreateRow(0);
            for (int i = 0; i < exportTable.Columns.Count; i++)
            {
                ICell headerCell = headerRow.CreateCell(i);
                headerCell.SetCellValue(exportTable.Columns[i].ColumnName?.ToString());
            }

            for (int i = 0; i < exportTable.Rows.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < exportTable.Columns.Count; j++)
                {
                    ICell dataCell = dataRow.CreateCell(j);
                    dataCell.SetCellValue(exportTable.Rows[i][j]?.ToString());
                }
            }
            for (int i = 0; i < exportTable.Columns.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            return workbook;
        }

        /// <summary>
        /// 生成导入模版Excel信息
        /// </summary>
        /// <param name="templateModels">模板定义信息</param>
        /// <param name="title">Sheet名称</param>
        /// <returns></returns>
        public static HSSFWorkbook ExportTemplate(List<ExcelTemplate> templateModels, string title)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            workbook.SetSheetName(0, title);

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = Company;
            workbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Author = Author;
            si.ApplicationName = ApplicationName;
            si.Title = title;

            si.CreateDateTime = DateTime.Now;
            workbook.SummaryInformation = si;

            //取得列宽
            int[] arrColWidth = new int[templateModels.Count];
            int columnIndex = 0;
            foreach (var templateModel in templateModels)
            {
                arrColWidth[columnIndex] = templateModel.CellLength > 0 ? templateModel.CellLength * 2 : Encoding.UTF8.GetBytes(templateModel.Name.ToString()).Length;
                columnIndex++;
            }

            var headerRow = sheet.CreateRow(0);
            columnIndex = 0;
            foreach (var templateModel in templateModels)
            {
                var cell = headerRow.CreateCell(columnIndex);
                if (!string.IsNullOrEmpty(templateModel.ExportComments))
                {
                    HSSFPatriarch patr = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                    HSSFComment comment = (HSSFComment)patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 1, 2, 4, 16));
                    comment.String = new HSSFRichTextString(templateModel.ExportComments);
                    comment.Author = ApplicationName;
                    cell.CellComment = comment;
                }

                if (templateModel.DictionaryItems != null && templateModel.DictionaryItems.Count > 0)
                {
                    DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(templateModel.DictionaryItems.ToArray());
                    CellRangeAddressList regions = new CellRangeAddressList(1, 65535, columnIndex, columnIndex);
                    IDataValidation validation = new HSSFDataValidation(regions, constraint);
                    sheet.AddValidationData(validation);
                }

                cell.SetCellValue(templateModel.Name);
                if (templateModel.IsRequred)
                {
                    var headStyle = workbook.CreateCellStyle();
                    headStyle.Alignment = HorizontalAlignment.Center;
                    var font = workbook.CreateFont();
                    font.Color = HSSFColor.Red.Index;
                    font.FontHeightInPoints = 10;
                    font.IsBold = true;
                    headStyle.SetFont(font);
                    cell.CellStyle = headStyle;
                }
                else
                {
                    var headStyle = workbook.CreateCellStyle();
                    headStyle.Alignment = HorizontalAlignment.Center;
                    var font = workbook.CreateFont();
                    font.Color = HSSFColor.Black.Index;
                    font.FontHeightInPoints = 10;
                    font.IsBold = true;
                    headStyle.SetFont(font);
                    cell.CellStyle = headStyle;
                }

                //设置列宽
                sheet.SetColumnWidth(columnIndex, (arrColWidth[columnIndex] + 1) * 256);
                columnIndex++;
            }

            return workbook;
        }

        #endregion

        #region 导入

        /// <summary>
        /// 导入Excel
        /// </summary>
        /// <param name="importTable">读取的Excel内容</param>
        /// <param name="templateModels">Excel模板定义</param>
        /// <param name="customerValidateMethod">用户自定义验证方法</param>
        /// <param name="importOtherColumns">添加的额外列信息</param>
        /// <param name="isSameUpdate">相同是否更新，否者输出错误</param>
        /// <param name="objParam">用户自定义验证公共信息</param>
        /// <param name="isShowErrorMsg">是否输出错误列信息</param>
        /// <returns></returns>
        public static ImportResult ImportFromExcels(DataTable importTable, IEnumerable<ExcelTemplate> templateModels, Func<DataRow, DataTable, DataTable, object, CustomerValidateResult> customerValidateMethod, Dictionary<string, object> importOtherColumns, bool isSameUpdate, object objParam, bool isShowErrorMsg = true)
        {
            ImportResult importResult = new ImportResult();
            Dictionary<string, ExcelTemplate> columnModeles;

            DataTable successTable = importTable.Clone();
            DataTable updateTable = importTable.Clone();
            AddOtherColumns(successTable, importOtherColumns);
            AddOtherColumns(updateTable, importOtherColumns);
            updateTable.Columns.Add("SameId");

            DataTable errorTable = CreateErrorTable(importTable);

            var isOwnerTemplate = ValidateExcelTemplate(importTable, templateModels, out columnModeles);
            if (!isOwnerTemplate)
            {
                importResult.IsSuccess = false;
                importResult.ImportResultMsg = "导入数据的模版有错误，请从正确的模版导入";
                return importResult;
            }

            int totalCount = importTable.Rows.Count;
            int totalSuccess = 0;
            int totalUpdate = 0;
            int totalError = 0;
            int rowIndex = 0;
            foreach (DataRow ownerRow in importTable.Rows)
            {
                var errorRow = errorTable.NewRow();
                ChangeRowToRow(ownerRow, errorRow, importTable.Columns);

                var tempError = ValidateTableRow(ownerRow, templateModels, importTable, columnModeles);

                var tempRowValidateResult = new CustomerValidateResult();
                if (customerValidateMethod != null)
                {
                    tempRowValidateResult = customerValidateMethod(ownerRow, importTable, successTable, objParam);
                    if (!tempRowValidateResult.IsSuccess)
                    {
                        tempError += tempRowValidateResult.ErrorMsg;
                    }
                }
                if (string.IsNullOrEmpty(tempError))
                {
                    if (!tempRowValidateResult.IsSame)
                    {
                        var successRow = successTable.NewRow();
                        ChangeRowToRow(ownerRow, successRow, importTable.Columns);
                        if (importOtherColumns != null)
                        {
                            foreach (var otherColumn in importOtherColumns)
                            {
                                successRow[otherColumn.Key] = otherColumn.Value;
                            }
                        }
                        successTable.Rows.Add(successRow);
                        totalSuccess++;
                    }
                    else
                    {
                        if (isSameUpdate)
                        {
                            var updateRow = updateTable.NewRow();
                            ChangeRowToRow(ownerRow, updateRow, importTable.Columns);
                            if (importOtherColumns != null)
                            {
                                foreach (var otherColumn in importOtherColumns)
                                {
                                    updateRow[otherColumn.Key] = otherColumn.Value;
                                }
                            }
                            updateRow["SameId"] = tempRowValidateResult.SameId;
                            updateTable.Rows.Add(updateRow);
                            totalUpdate++;
                        }
                        else
                        {
                            errorRow["错误提示"] = $"数据重复";
                            errorRow["错误行号"] = rowIndex + 2;
                            errorTable.Rows.Add(errorRow);
                            totalError++;
                        }
                    }
                }
                else
                {
                    errorRow["错误提示"] = tempError;
                    errorRow["错误行号"] = rowIndex + 2;
                    errorTable.Rows.Add(errorRow);
                    totalError++;
                }
                rowIndex++;
            }

            if (totalError != 0)
            {
                string strError = string.Empty;
                if (totalUpdate > 0)
                {
                    strError = "统计：导入" + totalCount + "条数据，新增成功" + totalSuccess + "条，更新成功" + totalUpdate + "条，失败" + totalError + "条，请修改后重新导入！";
                }
                else
                {
                    strError = "统计：导入" + totalCount + "条数据，导入成功" + totalSuccess + "条，导入失败" + totalError + "条，请修改后重新导入！";
                }
                if (isShowErrorMsg)
                {
                    var totalErrorRow = errorTable.NewRow();
                    totalErrorRow["错误行号"] = strError;
                    errorTable.Rows.Add(totalErrorRow);
                }
                importResult.ImportResultMsg = strError;
                importResult.IsSuccess = false;
            }
            else
            {
                string strImportResultMsg = string.Empty;
                if (totalUpdate > 0)
                {
                    strImportResultMsg = "导入统计：导入" + totalCount + "条数据，新增成功" + totalSuccess + "条数据，更新成功" + totalUpdate + "条数据！";
                }
                else
                {
                    strImportResultMsg = "导入统计：导入" + totalCount + "条数据！";
                }
                importResult.ImportResultMsg = strImportResultMsg;
                importResult.IsSuccess = true;
            }

            SetTableHeader(templateModels, successTable);
            SetTableHeader(templateModels, updateTable);
            importResult.SuccessTable = successTable;
            importResult.UpdateTable = updateTable;

            importResult.ErrorTable = errorTable;

            return importResult;

            //// 生成错误表格
            //TableToWorkbook(templateModels, errorTable);

            //// 导入数据到Excel
            //BulkInsertToTable(successTable, "SEC_User_Owner");
            //return importResult;
        }

        // 读取Excel内容到datatable
        public static DataTable ReadToDatatable(Stream stream, string fileName)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook = null;
            string fileExt = Path.GetExtension(fileName).ToLower();
            if (fileExt == ".xlsx")
            {
                workbook = new XSSFWorkbook(stream);
            }
            else if (fileExt == ".xls")
            {
                workbook = new HSSFWorkbook(stream);
            }

            ISheet sheet = workbook.GetSheetAt(0);

            //表头  
            IRow header = sheet.GetRow(sheet.FirstRowNum);
            List<int> columns = new List<int>();
            for (int i = 0; i < header.LastCellNum; i++)
            {
                object obj = GetValueType(header.GetCell(i));
                if (obj == null || obj.ToString() == string.Empty)
                {
                    dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                }
                else
                {
                    dt.Columns.Add(new DataColumn(obj.ToString()));
                }
                columns.Add(i);
            }
            //数据  
            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                DataRow dr = dt.NewRow();
                bool hasValue = false;
                foreach (int j in columns)
                {
                    dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                    if (dr[j] != null && dr[j].ToString() != string.Empty)
                    {
                        hasValue = true;
                    }
                }
                if (hasValue)
                {
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        try
                        {
                            return cell.DateCellValue.ToString();
                        }
                        catch (NullReferenceException)
                        {
                            return DateTime.FromOADate(cell.NumericCellValue).ToString();
                        }
                    }
                    return cell.NumericCellValue.ToString();
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }

        private static void AddOtherColumns(DataTable srcTable, Dictionary<string, object> importOtherColumns)
        {
            if (importOtherColumns != null)
            {
                foreach (var otherColumn in importOtherColumns)
                {
                    srcTable.Columns.Add(otherColumn.Key);
                }
            }
        }

        private static bool ValidateExcelTemplate(DataTable importTable, IEnumerable<ExcelTemplate> templateModels, out Dictionary<string, ExcelTemplate> columnModeles)
        {
            columnModeles = new Dictionary<string, ExcelTemplate>();
            var neetImportColumnCount = templateModels.Count();
            if (neetImportColumnCount != importTable.Columns.Count)
            {
                return false;
            }

            foreach (DataColumn tempColumn in importTable.Columns)
            {
                var templateModel = templateModels.SingleOrDefault(r => r.Name == tempColumn.ColumnName);
                if (templateModel == null)
                {
                    return false;
                }
                else
                {
                    columnModeles.Add(tempColumn.ColumnName, templateModel);
                }
            }

            return true;
        }

        private static string ValidateTableRow(DataRow validateRow, IEnumerable<ExcelTemplate> templateModels, DataTable validateTable, Dictionary<string, ExcelTemplate> columnModeles)
        {
            StringBuilder sbErrorValue = new StringBuilder();
            string tempStrError = "";
            foreach (DataColumn tempColumn in validateTable.Columns)
            {
                var columnValue = validateRow[tempColumn];
                var templateModel = columnModeles[tempColumn.ColumnName];

                if (!templateModel.IsRequred && string.IsNullOrEmpty(columnValue.ToString()))
                {
                    if (templateModel.ValidateType == EValidateType.Date || templateModel.ValidateType == EValidateType.Datetime)
                    {
                        validateRow[templateModel.Name] = null;
                    }
                    continue;
                }

                // 验证必填字段
                if (templateModel.IsRequred && string.IsNullOrEmpty(columnValue.ToString()))
                {
                    sbErrorValue.Append(tempColumn.ColumnName + "为必填字段;");
                }

                bool isValidate = ValidateTableColumn(validateRow, columnValue.ToString(), templateModels.FirstOrDefault(r => r.Name == tempColumn.ColumnName), out tempStrError);
                if (!isValidate)
                {
                    sbErrorValue.Append(tempStrError);
                }
            }

            return sbErrorValue.ToString();
        }

        private static bool ValidateTableColumn(DataRow validateRow, string cellValue, ExcelTemplate templateModel, out string strErrorMsg)
        {
            string strRegex = "";
            strErrorMsg = "";

            switch (templateModel.ValidateType)
            {
                case EValidateType.Empty:
                    break;
                case EValidateType.Int:
                    strRegex = @"^[1-9]+\d{0,}$";
                    strErrorMsg = "[" + templateModel.Name + "]填写错误，必须是数字;";
                    break;
                case EValidateType.Decimal:
                    strRegex = @"^([0]|[1-9]\d*|\d+\.\d{1,2})$";
                    strErrorMsg = "[" + templateModel.Name + "]填写错误，必须是数字且小数点后最多两位;";
                    break;
                case EValidateType.Double:
                    strRegex = @"^([1-9]\d*|\d+\.\d{1,2})$";
                    strErrorMsg = "[" + templateModel.Name + "]填写错误，必须是数字且小数点后最多两位;";
                    break;
                case EValidateType.String:
                    if (cellValue.Length > templateModel.CellLength)
                    {
                        strErrorMsg = "[" + templateModel.Name + "]填写错误，超过了最大长度;";
                        return false;
                    }
                    else
                    {
                        strRegex = @"^[^~#^$@%&!*]+$";
                        strErrorMsg = "[" + templateModel.Name + "]填写错误，不能为特殊字符;";
                        break;
                    }
                case EValidateType.Phone:
                    strRegex = @"^1(3|4|5|6|7|8|9)\d{9}$";
                    strErrorMsg = "[" + templateModel.Name + "]填写错误，必须是电话号码;";
                    break;
                case EValidateType.Idcard:
                    strRegex = @"^[1-9]\d{7}((0\d)|(1[0-2]))(([0|1|2]\d)|3[0-1])\d{3}$|^[1-9]\d{5}[1-9]\d{3}((0\d)|(1[0-2]))(([0|1|2]\d)|3[0-1])\d{3}([0-9]|X)$";
                    strErrorMsg = "[" + templateModel.Name + "]填写错误，必须是身份证号码;";
                    break;
                case EValidateType.Email:
                    strRegex = "^[a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(.[a-zA-Z0-9_-]+)+$";
                    strErrorMsg = "[" + templateModel.Name + "]填写错误，必须是邮箱地址;";
                    break;
                case EValidateType.Mobile:
                    strRegex = @"^1(3|4|5|6|7|8|9)\d{9}$";
                    strErrorMsg = "[" + templateModel.Name + "]填写错误，必须是手机号码;";
                    break;
                case EValidateType.Date:
                    DateTime tryDate;
                    var isDate = DateTime.TryParse(cellValue, out tryDate);
                    if (false == isDate)
                    {
                        strErrorMsg = "[" + templateModel.Name + "]填写错误，必须为日期类型(例如:2016-09-27);";
                        return false;
                    }
                    else
                    {
                        validateRow[templateModel.Name] = tryDate.ToString("yyyy-MM-dd");
                        return true;
                    }
                case EValidateType.Datetime:
                    DateTime tryTime;
                    var isDateTime = DateTime.TryParse(cellValue, out tryTime);
                    if (!isDateTime)
                    {
                        strErrorMsg = "[" + templateModel.Name + "]填写错误，必须为日期型;";
                        return false;
                    }
                    else
                    {
                        validateRow[templateModel.Name] = tryTime;
                        return true;
                    }
                case EValidateType.Dict:
                    strErrorMsg = "[" + templateModel.Name + "]填写错误，必须从下拉列表中选择;";
                    var dict = templateModel.DictionaryItems.SingleOrDefault(r => r == cellValue);
                    if (dict == null)
                    {
                        return false;
                    }
                    else
                    {
                        validateRow[templateModel.Name] = dict;
                        return true;
                    }
                case EValidateType.Regular:
                    strRegex = templateModel.ValidateValue;
                    strErrorMsg = "[" + templateModel.Name + "]验证错误";
                    break;
                default:
                    strRegex = templateModel.ValidateValue;
                    strErrorMsg = "[" + templateModel.Name + "]验证错误";
                    break;
            }
            return Regex.IsMatch(cellValue, strRegex);
        }

        /// <summary>
        /// 创建错误输出的Table
        /// </summary>
        /// <param name="srcTable">源Table</param>
        /// <returns>错误输出的Table</returns>
        private static DataTable CreateErrorTable(DataTable srcTable)
        {
            DataTable errorTable = new DataTable();
            errorTable.Columns.Add("错误行号");
            foreach (DataColumn srcColumn in srcTable.Columns)
            {
                errorTable.Columns.Add(srcColumn.Caption);
            }
            errorTable.Columns.Add("错误提示");

            return errorTable;
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 把一个Table行的数据迁移到另外一个Table行
        /// </summary>
        /// <param name="srcRow">源Table行</param>
        /// <param name="tagRow">目标Table行</param>
        /// <param name="srcRowColumns">目标表集合</param>
        public static void ChangeRowToRow(DataRow srcRow, DataRow tagRow, DataColumnCollection srcRowColumns)
        {
            foreach (DataColumn srcColumn in srcRowColumns)
            {
                string strColumnValue = srcRow[srcColumn.ColumnName].ToString().Trim();
                if (!string.IsNullOrEmpty(strColumnValue))
                {
                    srcRow[srcColumn.ColumnName] = strColumnValue;
                }

                if (!string.IsNullOrEmpty(srcRow[srcColumn.ColumnName].ToString()))
                {
                    tagRow[srcColumn.ColumnName] = srcRow[srcColumn.ColumnName];
                }
            }
        }

        // Excel Table表名设置
        public static void SetTableHeader(IEnumerable<ExcelTemplate> templateModels, DataTable dataTable)
        {
            foreach (DataColumn tempColumn in dataTable.Columns)
            {
                var templateModel = templateModels.SingleOrDefault(r => r.Name == tempColumn.ColumnName);
                if (templateModel != null)
                {
                    tempColumn.ColumnName = templateModel.Field;
                }
            }
        }

        /// <summary>
        /// 创建指定实体的更新SQL语句
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateModels"></param>
        /// <returns></returns>
        public static string CreateUpdateSQL<T>(IEnumerable<ExcelTemplate> templateModels, string tableName) where T : class, new()
        {
            T t = new T();
            StringBuilder builder = new StringBuilder();
            builder.Append(" UPDATE " + string.Format("`{0}`", tableName) + " SET ");
            //能导出的属性
            //所有属性
            var properties = t.GetType().GetProperties();
            foreach (var p in properties)
            {
                //获取当前属性
                var _model = templateModels.FirstOrDefault(f => f.Field == p.Name);
                if (null != _model)
                {
                    builder.Append(string.Format("`{0}` = @{0} ,", _model.Field));
                }
            }
            builder.Remove(builder.Length - 1, 1);

            builder.Append(" WHERE `Id` = @Id;");

            return builder.ToString();
        }

        #endregion

        #region Datatable List转换

        public static DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }
            return table;
        }

        /// <summary>
        /// DataTable转成List
        /// </summary>
        public static List<T> ConvertToList<T>(DataTable dt)
        {
            var list = new List<T>();
            var plist = new List<PropertyInfo>(typeof(T).GetProperties());

            if (dt == null || dt.Rows.Count == 0)
            {
                return null;
            }

            foreach (DataRow item in dt.Rows)
            {
                T s = Activator.CreateInstance<T>();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    PropertyInfo info = plist.Find(p => p.Name == dt.Columns[i].ColumnName);
                    if (info != null)
                    {
                        try
                        {
                            if (!Convert.IsDBNull(item[i]))
                            {
                                object v = null;
                                if (info.PropertyType.ToString().Contains("System.Nullable"))
                                {
                                    v = Convert.ChangeType(item[i], Nullable.GetUnderlyingType(info.PropertyType));
                                }
                                else
                                {
                                    v = Convert.ChangeType(item[i], info.PropertyType);
                                }
                                info.SetValue(s, v, null);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("字段[" + info.Name + "]转换出错," + ex.Message);
                        }
                    }
                }
                list.Add(s);
            }
            return list;
        }

        #endregion
    }
}
