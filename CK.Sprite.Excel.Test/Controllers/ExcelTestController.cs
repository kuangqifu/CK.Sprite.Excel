using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;

namespace CK.Sprite.Excel.Test
{
    /// <summary>
    /// Excel导入导出测试
    /// </summary>
    [ApiController]
    [Area("excel")]
    [ControllerName("ExcelTest")]
    [Route("api/excel/ExcelTest")]
    public class ExcelTestController : Controller
    {
        /// <summary>
        /// 导入测试
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("ExcelImportTest")]
        public ActionResult ExcelImportTest(IFormFile file)
        {
            var files = Request.Form.Files;
            var importResultInfo = ExcelHttpHelper.GetUploadInfos(files, GetTemplate(), ValidateStudentColumn, null, true, null);

            return HandleImportResult(importResultInfo, "学生测试数据导入错误", (importResult) =>
            {
                // 处理导入逻辑
            });
        }

        /// <summary>
        /// 导入模版导出
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("ExcelExportTemplateTest")]
        public ActionResult ExcelExportTemplateTest()
        {
            var exportResult = ExcelHelper.ExportTemplate(GetTemplate(), "导出测试模版");

            return ExportResult(exportResult, "导出模版测试.xls");
        }

        /// <summary>
        /// 基础导出
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("ExcelExportDataTest")]
        public ActionResult ExcelExportDataTest()
        {
            List<Student> students = new List<Student>();
            students.Add(new Student()
            {
                Age = 11,
                BirthDay = DateTime.Now,
                Name = "kuangqifu",
                Field1 = "Field11",
                Field2 = "Field21",
                Field3 = "Field31",
            });

            students.Add(new Student()
            {
                Age = 22,
                BirthDay = DateTime.Now.AddDays(2),
                Name = "chenfang",
                Field1 = "Field12",
                Field2 = "Field22",
                Field3 = "Field32",
            });

            students.Add(new Student()
            {
                Age = 33,
                BirthDay = DateTime.Now.AddDays(3),
                Name = "chenchen",
                Field1 = "Field13",
                Field2 = "Field23",
                Field3 = "Field33",
            });

            var dtStudents = ExcelHelper.ConvertToDataTable<Student>(students);

            var exportResult = ExcelHelper.Export<Student>(students, GetTemplate(), "学生测试数据");

            return ExportResult(exportResult, "导出数据测试.xls");
        }

        /// <summary>
        /// 多表头导出
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("MultiHeaderExcelExportDataTest")]
        public ActionResult MultiHeaderExcelExportDataTest()
        {
            List<Student> students = new List<Student>();
            students.Add(new Student()
            {
                Age = 11,
                BirthDay = DateTime.Now,
                Name = "kuangqifu",
                Field1 = "Field11",
                Field2 = "Field21",
                Field3 = "Field31",
            });

            students.Add(new Student()
            {
                Age = 22,
                BirthDay = DateTime.Now.AddDays(2),
                Name = "chenfang",
                Field1 = "Field12",
                Field2 = "Field22",
                Field3 = "Field32",
            });

            students.Add(new Student()
            {
                Age = 33,
                BirthDay = DateTime.Now.AddDays(3),
                Name = "chenchen",
                Field1 = "Field13",
                Field2 = "Field23",
                Field3 = "Field33",
            });

            var dtStudents = ExcelHelper.ConvertToDataTable<Student>(students);

            var exportResult = ExcelHelper.Export<Student>(students, GetTemplate(), "学生测试数据", GetMultiHeaderInfos());

            return ExportResult(exportResult, "多表头导出数据测试.xls");
        }

        private ActionResult HandleImportResult(ImportResult importResultInfo, string errorName, Action<ImportResult> action)
        {
            if (importResultInfo.IsSuccess)
            {
                try
                {
                    action(importResultInfo);
                }
                catch (Exception ex)
                {
                    Response.Headers.Add("ImportResult", HttpUtility.UrlEncode(ex.Message));
                    Response.Headers.Add("HaveErrorTable", "false");
                    return Content(ex.Message);
                }

                Response.Headers.Add("ImportResult", HttpUtility.UrlEncode(importResultInfo.ImportResultMsg));
                Response.Headers.Add("HaveErrorTable", "false");
                return Content(importResultInfo.ImportResultMsg);
            }
            else
            {
                if (importResultInfo.ErrorTable == null)
                {
                    Response.Headers.Add("ImportResult", HttpUtility.UrlEncode(importResultInfo.ImportResultMsg));
                    Response.Headers.Add("HaveErrorTable", "false");
                    return Content(importResultInfo.ImportResultMsg);
                }
                else
                {
                    try
                    {
                        action(importResultInfo);
                    }
                    catch (Exception ex)
                    {
                        Response.Headers.Add("ImportResult", HttpUtility.UrlEncode(ex.Message));
                        Response.Headers.Add("HaveErrorTable", "false");
                        return Content(ex.Message);
                    }

                    var exportResult = ExcelHelper.ExportFromTable(importResultInfo.ErrorTable, errorName);
                    Response.Headers.Add("ImportResult", HttpUtility.UrlEncode(importResultInfo.ImportResultMsg));
                    Response.Headers.Add("HaveErrorTable", "true");
                    Response.Headers.Add("ErrorTableName", HttpUtility.UrlEncode($"{errorName}.xls"));

                    return ExportResult(exportResult, $"{errorName}.xls");
                }
            }
        }

        private ActionResult ExportResult(HSSFWorkbook exportResult, string fileName)
        {
            byte[] buffer;
            using (MemoryStream ms = new MemoryStream())
            {
                exportResult.Write(ms);
                buffer = ms.ToArray();
                ms.Close();
            }

            return File(
                fileContents: buffer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: fileName
            );
        }

        private CustomerValidateResult ValidateStudentColumn(DataRow validateRow, DataTable importTable, DataTable tempSuccessTable, object objParam)
        {
            var result = new CustomerValidateResult();
            List<string> sameName = new List<string>() { "kuangqifu", "chenfang", "chenchen" };
            var name = validateRow["姓名"].ToString();
            if (sameName.Contains(name))
            {
                result.IsSame = true;
                result.SameId = name;
            }

            return result;
        }

        private static List<ExcelTemplate> GetTemplate()
        {
            var excelTemplates = new List<ExcelTemplate>();
            excelTemplates.Add(new ExcelTemplate()
            {
                Field = "Name",
                FieldType = EFieldType.String,
                ExportComments = "姓名，请重下拉中选择",
                Name = "姓名",
                ValidateType = EValidateType.Dict,
                DictionaryItems = new List<string>() { "kuangqifu", "chenfang", "chenchen" },
                IsRequred = true
            });

            excelTemplates.Add(new ExcelTemplate()
            {
                Field = "Age",
                FieldType = EFieldType.Int,
                ExportComments = "年龄",
                Name = "年龄",
                ValidateType = EValidateType.Int,
            });

            excelTemplates.Add(new ExcelTemplate()
            {
                Field = "BirthDay",
                FieldType = EFieldType.DateTime,
                Name = "生日",
                ValidateType = EValidateType.Datetime,
                CellLength = 10
            });

            excelTemplates.Add(new ExcelTemplate()
            {
                Field = "Field1",
                FieldType = EFieldType.String,
                ExportComments = "扩展字段1",
                Name = "扩展字段1",
                ValidateType = EValidateType.String,
                CellLength = 10,
                IsRequred = true
            });

            excelTemplates.Add(new ExcelTemplate()
            {
                Field = "Field2",
                FieldType = EFieldType.String,
                ExportComments = "扩展字段2",
                Name = "扩展字段2",
                ValidateType = EValidateType.String,
                CellLength = 20,
                IsRequred = true
            });

            excelTemplates.Add(new ExcelTemplate()
            {
                Field = "Field3",
                FieldType = EFieldType.String,
                ExportComments = "扩展字段3",
                Name = "扩展字段3",
                ValidateType = EValidateType.String,
                CellLength = 20,
                IsRequred = false
            });

            return excelTemplates;
        }

        private static List<List<MultiHeaderInfo>> GetMultiHeaderInfos()
        {
            List<List<MultiHeaderInfo>> multiHeaderInfos = new List<List<MultiHeaderInfo>>();
            List<MultiHeaderInfo> multiHeaders1 = new List<MultiHeaderInfo>()
            {
                new MultiHeaderInfo()
                {
                     RowSpan = 2,
                      Name = "姓名"
                },
                new MultiHeaderInfo()
                {
                     RowSpan = 2,
                     Name = "年龄"
                },
                new MultiHeaderInfo()
                {
                     RowSpan = 2,
                      Name = "生日"
                },
                new MultiHeaderInfo()
                {
                     ColSpan = 2,
                      Name = "合并两列"
                },
                new MultiHeaderInfo()
                {
                     RowSpan = 2,
                      Name = "扩展列3"
                }
            };
            multiHeaderInfos.Add(multiHeaders1);

            List<MultiHeaderInfo> multiHeaders2 = new List<MultiHeaderInfo>()
            {
                new MultiHeaderInfo()
                {
                      Name = "扩展列1"
                },
                new MultiHeaderInfo()
                {
                      Name = "扩展列2"
                }
            };
            multiHeaderInfos.Add(multiHeaders2);

            return multiHeaderInfos;
        }
    }

    public class Student
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public DateTime BirthDay { get; set; }
        public string Field1 { get; set; }
        public string Field2 { get; set; }
        public string Field3 { get; set; }
    }
}
