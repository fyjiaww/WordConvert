using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WordConvert.Models
{
    /// <summary>
    /// 简历信息
    /// </summary>
    public class CurriculumVitae
    {
        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 性别
        /// </summary>
        public string Sex { get; set; }
        /// <summary>
        /// 年龄
        /// </summary>
        public string Age { get; set; }
        /// <summary>
        /// 学历
        /// </summary>
        public string Education { get; set; }
        /// <summary>
        /// 手机
        /// </summary>
        public string Phone { get; set; }
        /// <summary>
        /// 电子邮件
        /// </summary>
        public string Email { get; set; }
        /// <summary>
        /// 英语等级
        /// </summary>
        public string EnglishLevel { get; set; }
        /// <summary>
        /// 求职意向
        /// </summary>
        public string JobIntension { get; set; }
        /// <summary>
        /// 工作地点
        /// </summary>
        public string WorkPlace { get; set; }
        /// <summary>
        ///工作年限
        /// </summary>
        public string WorkLife { get; set; }
        /// <summary>
        /// 期望薪水
        /// </summary>
        public string Salary { get; set; }
        /// <summary>
        /// 毕业院校
        /// </summary>
        public string School { get; set; }
        /// <summary>
        /// 专业
        /// </summary>
        public string Major { get; set; }
        /// <summary>
        /// 最近单位
        /// </summary>
        public string RecentWorkUnits { get; set; }
        /// <summary>
        /// 最近职位
        /// </summary>
        public string RecnetJob { get; set; }
        /// <summary>
        /// 接收时间
        /// </summary>
        public DateTime  ReceiptTime { get; set; }
    }
}