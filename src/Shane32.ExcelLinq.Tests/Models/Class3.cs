namespace Shane32.ExcelLinq.Tests.Models
{
    class Class3
    {
        public int Invalid1 { get; }
        public int Invalid2 { set { } }
        public int Invalid3() => 0;
        public int Valid1;
        public int Valid2 { get; set; }
        public int? Valid3;
        public string Valid4;
    }
}
