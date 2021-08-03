using Npoi.Mapper.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelDemo.Models
{
    public class User
    {
        [Column("用户ID")]
        public int UserId { get; set; }
        [Column("用户名")]
        public string UserName { get; set; }
        [Column("手机号")]
        public string Phone { get; set; }
        [Column("邮箱")]
        public string Email { get; set; }
    }
}
