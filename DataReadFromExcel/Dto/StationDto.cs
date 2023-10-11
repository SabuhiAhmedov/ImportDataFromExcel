namespace DataReadFromExcel.Dto
{
    public class StationDto
    {
        public string Name { get; set; }
        public List<CounterDto> Counters { get; set; }
    }
}
