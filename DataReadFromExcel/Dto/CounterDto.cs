namespace DataReadFromExcel.Dto
{
    public class CounterDto
    {

        public string? Name { get; set; }
        public decimal? Dmm { get; set; }
        public decimal? dmm1 { get; set; }
        public List<DataDto> Changes { get; set; }
        public decimal Total { get; set; }
      
    }
}
