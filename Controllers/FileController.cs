using ExcelDemo.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelDemo.Controllers
{
    public class FileController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public IActionResult Upload()
        {
            //用户List
            List<User> users = new List<User>();
            //获取到上传的文件
            var file = Request.Form.Files[0];
            //文件转成流
            using (var fileStream = file.OpenReadStream())
            {
                //excel转实体
                users = ExcelHelper.ParseExcelToList<User>(fileStream, "Sheet1");
            }
            return Content(JsonConvert.SerializeObject(users));

        }
        public IActionResult DownLoad()
        {
            //生成List
            List<User> usrs = new List<User>();
            for(int i=1;i<=10;i++)
            {
                User user = new User()
                {
                    UserId = i,
                    UserName = "用户" + i,
                    Phone = "13800000000",
                    Email = $"user{i}@qq.com"
                };
                usrs.Add(user);
            }
            //List转为Excel文件
            var fileStream = ExcelHelper.ParseListToExcel(usrs);
            return File(fileStream.ToArray(), "application/vnd.ms-excel", "用户信息.xlsx");
        }
    }
}
