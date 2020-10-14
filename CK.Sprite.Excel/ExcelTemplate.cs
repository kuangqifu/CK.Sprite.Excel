using System;
using System.Collections.Generic;
using System.Data;

namespace CK.Sprite.Excel
{
    public class ExcelTemplate
    {
        /// <summary>
        /// 字段名称
        /// </summary>
        public string Field { get; set; }

        /// <summary>
        /// 列称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 字段类型
        /// </summary>
        public EFieldType FieldType { get; set; }

        /// <summary>
        /// 列宽（显示多少个字符）
        /// </summary>
        public int CellLength { get; set; }

        /// <summary>
        /// Excel下拉值（如果是数据字典，读取数据字典内容）
        /// </summary>
        public List<string> DictionaryItems { get; set; }

        /// <summary>
        /// 导出模版备注
        /// </summary>
        public string ExportComments { get; set; }

        /// <summary>
        /// 导入 是否必填
        /// </summary>
        public bool IsRequred { get; set; }

        /// <summary>
        /// 导入 验证类型
        /// </summary>
        public EValidateType ValidateType { get; set; }

        /// <summary>
        /// 导入 验证类型为String时，验证长度，为Regular，为正则表达式
        /// </summary>
        public string ValidateValue { get; set; }
    }

    public class MultiHeaderInfo
    {
        public string Name { get; set; }
        public int ColSpan { get; set; } = 1;
        public int RowSpan { get; set; } = 1;
    }

    /// <summary>
    /// 导入结果
    /// </summary>
    [Serializable]
    public class ImportResult
    {
        /// <summary>
        /// 导入是否成功
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 导入结果
        /// </summary>
        public string ImportResultMsg { get; set; }

        /// <summary>
        /// 错误的Table输出
        /// </summary>
        public DataTable ErrorTable { get; set; }

        /// <summary>
        /// 成功的Table输出
        /// </summary>
        public DataTable SuccessTable { get; set; }

        /// <summary>
        /// 相同数据Table输出（导入时，设置相同数据更新用）
        /// </summary>
        public DataTable UpdateTable { get; set; }
    }

    /// <summary>
    /// 用户自定义验证范型方法验证结果
    /// </summary>
    public class CustomerValidateResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 判断是否重复
        /// </summary>
        public bool IsSame { get; set; }

        /// <summary>
        /// 错误内容
        /// </summary>
        public string ErrorMsg { get; set; }

        /// <summary>
        /// 相同数据Id
        /// </summary>
        public string SameId { get; set; }
    }

    /// <summary>
    /// 导入Excel输出
    /// </summary>
    public class ImportExcelResponse
    {
        /// <summary>
        /// 临时File文件Id
        /// </summary>
        public string TempFileId { get; set; }

        /// <summary>
        /// 是否存在错误
        /// </summary>
        public bool IsError { get; set; }

        /// <summary>
        /// 导出信息提示
        /// </summary>
        public string ImportMessage { get; set; }
    }

    public enum EFieldType
    {
        Int = 1,
        Double = 2,
        Guid = 3,
        Bool = 4,
        String = 5,
        Date = 6,
        DateTime = 7
    }

    public enum EValidateType
    {
        Empty = 0,
        Int = 1,
        Decimal = 2,
        Double = 3,
        String = 4,
        Phone = 5,
        Idcard = 6,
        Email = 7,
        Mobile = 8,
        Date = 9,
        Datetime = 10,
        Dict = 11,
        Regular = 99
    }

    public enum ESpecialColumnType
    {
        EnumInt = 1,
        DateTime = 2
    }
}
