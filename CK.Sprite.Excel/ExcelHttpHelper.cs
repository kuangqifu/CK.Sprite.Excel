using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace CK.Sprite.Excel
{
    public class ExcelHttpHelper
    {
        /// <summary>
        /// 获取导入Excel验证结果
        /// </summary>
        /// <param name="files">From上传的文件信息</param>
        /// <param name="excelTemplates">导入模版信息</param>
        /// <param name="customerValidateMethod">客户自定义验证方法</param>
        /// <param name="importOtherColumns">是否增加额外字段</param>
        /// <param name="isSameUpdate">数据相同是否修改</param>
        /// <param name="objParam">验证回调方法参数</param>
        /// <returns></returns>
        public static ImportResult GetUploadInfos(IFormFileCollection files,
            List<ExcelTemplate> excelTemplates,
            Func<DataRow, DataTable, DataTable, object, CustomerValidateResult> customerValidateMethod,
            Dictionary<string, object> importOtherColumns,
            bool isSameUpdate,object objParam)
        {
            if (files == null || files.Count <= 0)
            {
                return new ImportResult()
                {
                    IsSuccess = false,
                    ImportResultMsg = "未找到上传的文件"
                };
            }
            var file = files[0];
            if(!Path.GetExtension(file.FileName).Contains(".xls") && !Path.GetExtension(file.FileName).Contains(".xlsx"))
            {
                return new ImportResult()
                {
                    IsSuccess = false,
                    ImportResultMsg = "上传的文件格式不对"
                };
            }
            var stream = file.OpenReadStream();
            var dt = ExcelHelper.ReadToDatatable(stream, file.FileName);
            var importResult = ExcelHelper.ImportFromExcels(dt, excelTemplates, customerValidateMethod, importOtherColumns, isSameUpdate, objParam);

            return importResult;
        }
    }
}
