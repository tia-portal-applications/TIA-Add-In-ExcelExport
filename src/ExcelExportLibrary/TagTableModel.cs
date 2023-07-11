namespace ExcelExportLibrary
{
    public class TagTableModel
    {
        public string TagName { get; set; }
        public string DataType { get; set; }
        public string Address { get; set; }
        public bool ExternalAccessible { get; set; }
        public bool ExternalVisible { get; set; }
        public bool ExternalWritable { get; set; }
    }
}