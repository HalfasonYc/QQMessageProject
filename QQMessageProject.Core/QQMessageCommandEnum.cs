using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QQMessageProject.Core
{
    public enum QQMessageCommand
    {
        /// <summary>
        /// 无效命令
        /// </summary>
        None = -1,
        /// <summary>
        /// 取当前系统时间
        /// </summary>
        Time = 0,
        /// <summary>
        /// 结束运行
        /// </summary>
        End = 1,
    }
}
