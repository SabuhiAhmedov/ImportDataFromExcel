using DataReadFromExcel.Dto;
using DataReadFromExcel.Helper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Globalization;
using System.Security.Cryptography.X509Certificates;
using System.Xml.Linq;
using static DataReadFromExcel.Dto.DataDto;

namespace DataReadFromExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TaskController : ControllerBase
    {
        [HttpGet("ReadData")]
        public ActionResult ReadData()
        {
            List<StationDto> Data = DataList.DataListExcel();
            return Ok(Data.Take(1));

        }

        [HttpGet("ReadDataAsDynamicColumns")]
        public ActionResult ReadDataAsDynamicColums( )
        {
            ColumnsDto dto=new ColumnsDto();
            List<StationDto> Data = DataList.DataListExcelDynamicColumns(dto);
          return Ok(Data.Take(1));

        }
    }
}

