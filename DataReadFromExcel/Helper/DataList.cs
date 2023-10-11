using DataReadFromExcel.Dto;
using Microsoft.AspNetCore.Http.HttpResults;
using OfficeOpenXml;

namespace DataReadFromExcel.Helper
{
    public static class DataList
    {
      public  static List<StationDto> DataListExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string path = "C:\\Users\\Admin\\Downloads\\2023.07.25-30_saatlıq (2).xlsx";
            var package = new ExcelPackage(new FileInfo(path));
            var workSheet = package.Workbook.Worksheets[0];
            List<DataDto> dataList = new List<DataDto>();
            var endResult = new List<StationDto>();
            List<CounterDto> counterList = new List<CounterDto>();
            CounterDto findCounter = new CounterDto();
            StationDto findStation = new StationDto();
            int compute = 0;
            int helperVariable = 0;
            int helperVariable1 = 0;

            for (int row = 10; row <= workSheet.Dimension.End.Row; row++)
            {
                var date = workSheet.Cells[row, 1].Value?.ToString();
                var weight = workSheet.Cells[row, 2].Value?.ToString();
                var pressureDifference = workSheet.Cells[row, 3].Value?.ToString();
                var pressure = workSheet.Cells[row, 4].Value?.ToString();
                var temprature = workSheet.Cells[row, 5].Value?.ToString();
                var hourlySpend = workSheet.Cells[row, 6].Value?.ToString();
                var spend = workSheet.Cells[row, 7].Value?.ToString();

                if (date != null && weight != null && pressure != null && pressureDifference != null && temprature != null
                    && hourlySpend != null && spend != null && double.TryParse(temprature, out double a))
                {
                    compute++;
                    DateTime dateTimeValue = workSheet.Cells[row, 1].GetValue<DateTime>();
                    decimal.TryParse(weight, out decimal weightD);
                    decimal.TryParse(pressureDifference, out decimal pressureDifferenceD);
                    decimal.TryParse(pressure, out decimal pressureD);
                    decimal.TryParse(temprature, out decimal tempratureD);
                    decimal.TryParse(hourlySpend, out decimal hourlySpendD);
                    decimal.TryParse(spend, out decimal spendD);

                    DataDto newData = new DataDto
                    {
                        Date = dateTimeValue,
                        SpecialWeight = weightD,
                        Pressure = pressureD,
                        PressureDifference = pressureDifferenceD,
                        HourlySpend = hourlySpendD,
                        Spend = spendD,
                        Temperature = tempratureD
                    };
                    dataList.Add(newData);
                    if (compute == 1)
                    {
                        findStation = new StationDto() { Name = workSheet.Cells[row - 5, 1].Value?.ToString() };
                        findCounter = new CounterDto() { Name = workSheet.Cells[row - 4, 1].Value?.ToString() };
                    }
                    if (workSheet.Cells[row + 1, 1].Value?.ToString() == "CƏMi" && workSheet.Cells[row + 3, 1].Value?.ToString() == "D(mm):")

                    {
                        findCounter.Changes = new List<DataDto>();
                        findCounter.Changes.AddRange(dataList);
                        decimal.TryParse(workSheet.Cells[row + 1, 6].Value?.ToString(), out decimal total);
                        findCounter.Total = total;
                        decimal.TryParse(workSheet.Cells[row + 3, 2].Value?.ToString(), out decimal ConvertDmm);
                        findCounter.Dmm = ConvertDmm;
                        decimal.TryParse(workSheet.Cells[row + 3, 4].Value?.ToString(), out decimal Convertdmm);
                        findCounter.dmm1 = Convertdmm;
                        counterList.Add(findCounter);
                        findCounter = new CounterDto() { Name = workSheet.Cells[row + 2, 1].Value?.ToString() };
                        helperVariable = row;
                        dataList = new List<DataDto>();
                    }
                    else if (workSheet.Cells[row + 1, 1].Value?.ToString() == "CƏMi" && workSheet.Cells[row + 4, 1].Value?.ToString() == "D(mm):")
                    {
                        helperVariable1 = row;
                        findCounter.Changes = new List<DataDto>();
                        findCounter.Changes.AddRange(dataList);
                        decimal.TryParse(workSheet.Cells[row - (helperVariable1 - helperVariable) + 3, 2].Value?.ToString(), out decimal ConvertDmm);
                        findCounter.Dmm = ConvertDmm;
                        decimal.TryParse(workSheet.Cells[row - (helperVariable1 - helperVariable) + 3, 4].Value?.ToString(), out decimal Convertdmm);
                        findCounter.dmm1 = Convertdmm;
                        decimal.TryParse(workSheet.Cells[row + 1, 6].Value?.ToString(), out decimal total);
                        findCounter.Total = total;
                        counterList.Add(findCounter);
                        findCounter = new CounterDto() { Name = workSheet.Cells[row + 3, 1].Value?.ToString() };
                        dataList = new List<DataDto>();
                        findStation.Counters = new List<CounterDto>();
                        findStation.Counters.AddRange(counterList);
                        endResult.Add(findStation);
                        counterList = new List<CounterDto>();
                        findStation = new StationDto() { Name = workSheet.Cells[row + 2, 1].Value?.ToString() };
                    }
                }
            }
            return endResult;
        }
        public static List<StationDto> DataListExcelDynamicColumns(ColumnsDto columnsDto)
        {

            #region Read data as dynamic columns
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string path = "C:\\Users\\Admin\\Downloads\\2023.07.25-30_saatlıq (2).xlsx";
            var package = new ExcelPackage(new FileInfo(path));
            var workSheet = package.Workbook.Worksheets[0];
            List<DataDto> dataList = new List<DataDto>();
            var endResult = new List<StationDto>();
            List<CounterDto> counterList = new List<CounterDto>();
            CounterDto findCounter = new CounterDto();
            StationDto findStation = new StationDto();
            int compute = 0;
            int helperVariable = 0;
            int helperVariable1 = 0;



            for (int row = 10; row <= workSheet.Dimension.End.Row; row++)
            {
                var date = workSheet.Cells[row, columnsDto.DateId].Value?.ToString();
                var weight = workSheet.Cells[row, columnsDto.WeightId].Value?.ToString();
                var pressureDifference = workSheet.Cells[row, columnsDto.PressureDifferenceId].Value?.ToString();
                var pressure = workSheet.Cells[row, columnsDto.PressureId].Value?.ToString();
                var temprature = workSheet.Cells[row, columnsDto.TempratureId].Value?.ToString();
                var hourlySpend = workSheet.Cells[row, columnsDto.HourlySpendId].Value?.ToString();
                var spend = workSheet.Cells[row, columnsDto.SpendId].Value?.ToString();

                if (date != null && weight != null && pressure != null && pressureDifference != null && temprature != null
                    && hourlySpend != null && spend != null && double.TryParse(temprature, out double a))
                {
                    compute++;
                    DateTime dateTimeValue = workSheet.Cells[row, 1].GetValue<DateTime>();
                    decimal.TryParse(weight, out decimal weightD);
                    decimal.TryParse(pressureDifference, out decimal pressureDifferenceD);
                    decimal.TryParse(pressure, out decimal pressureD);
                    decimal.TryParse(temprature, out decimal tempratureD);
                    decimal.TryParse(hourlySpend, out decimal hourlySpendD);
                    decimal.TryParse(spend, out decimal spendD);

                    DataDto newData = new DataDto
                    {
                        Date = dateTimeValue,
                        SpecialWeight = weightD,
                        Pressure = pressureD,
                        PressureDifference = pressureDifferenceD,
                        HourlySpend = hourlySpendD,
                        Spend = spendD,
                        Temperature = tempratureD
                    };
                    dataList.Add(newData);
                    if (compute == 1)
                    {
                        findStation = new StationDto() { Name = workSheet.Cells[row - 5, 1].Value?.ToString() };
                        findCounter = new CounterDto() { Name = workSheet.Cells[row - 4, 1].Value?.ToString() };
                    }
                    if (workSheet.Cells[row + 1, 1].Value?.ToString() == "CƏMi" && workSheet.Cells[row + 3, 1].Value?.ToString() == "D(mm):")

                    {
                        findCounter.Changes = new List<DataDto>();
                        findCounter.Changes.AddRange(dataList);
                        decimal.TryParse(workSheet.Cells[row + 1, 6].Value?.ToString(), out decimal total);
                        findCounter.Total = total;
                        decimal.TryParse(workSheet.Cells[row + 3, 2].Value?.ToString(), out decimal ConvertDmm);
                        findCounter.Dmm = ConvertDmm;
                        decimal.TryParse(workSheet.Cells[row + 3, 4].Value?.ToString(), out decimal Convertdmm);
                        findCounter.dmm1 = Convertdmm;
                        counterList.Add(findCounter);
                        findCounter = new CounterDto() { Name = workSheet.Cells[row + 2, 1].Value?.ToString() };
                        helperVariable = row;
                        dataList = new List<DataDto>();
                    }
                    else if (workSheet.Cells[row + 1, 1].Value?.ToString() == "CƏMi" && workSheet.Cells[row + 4, 1].Value?.ToString() == "D(mm):")
                    {
                        helperVariable1 = row;
                        findCounter.Changes = new List<DataDto>();
                        findCounter.Changes.AddRange(dataList);
                        decimal.TryParse(workSheet.Cells[row - (helperVariable1 - helperVariable) + 3, 2].Value?.ToString(), out decimal ConvertDmm);
                        findCounter.Dmm = ConvertDmm;
                        decimal.TryParse(workSheet.Cells[row - (helperVariable1 - helperVariable) + 3, 4].Value?.ToString(), out decimal Convertdmm);
                        findCounter.dmm1 = Convertdmm;
                        decimal.TryParse(workSheet.Cells[row + 1, 6].Value?.ToString(), out decimal total);
                        findCounter.Total = total;
                        counterList.Add(findCounter);
                        findCounter = new CounterDto() { Name = workSheet.Cells[row + 3, 1].Value?.ToString() };
                        dataList = new List<DataDto>();
                        findStation.Counters = new List<CounterDto>();
                        findStation.Counters.AddRange(counterList);
                        endResult.Add(findStation);
                        counterList = new List<CounterDto>();
                        findStation = new StationDto() { Name = workSheet.Cells[row + 2, 1].Value?.ToString() };
                    }
                }
            }

            return endResult;
            #endregion
        }
    }
}
