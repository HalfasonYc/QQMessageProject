using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessibleProject
{
    public struct AccessibleInformation
    {
        [DisplayName("窗口标题"), System.ComponentModel.Category("控件信息"), Description("窗口标题（TitleWindow）")]
        public string TitleWindow { get; set; }

        [DisplayName("窗口类名"), System.ComponentModel.Category("控件信息"), Description("窗口类名（ClassNameWindow）")]
        public string ClassNameWindow { get; set; }

        [DisplayName("窗口句柄"), System.ComponentModel.Category("控件信息"), Description("窗口句柄（HandleWindow）")]
        public IntPtr HandleWindow { get; set; }

        [DisplayName("子窗体数"), System.ComponentModel.Category("控件信息"), Description("子窗体数（AccChildCount）")]
        public int AccChildCount { get; set; }

        [DisplayName("控件ID"), System.ComponentModel.Category("控件信息"), Description("控件ID（AccChildId）")]
        public int AccChildId { get; set; }

        [DisplayName("控件描述"), System.ComponentModel.Category("控件信息"), Description("控件描述（AccDescription）")]
        public string AccDescription { get; set; }

        [DisplayName("控件帮助"), System.ComponentModel.Category("控件信息"), Description("控件帮助（AccHelp）")]
        public string AccHelp { get; set; }

        [DisplayName("控件默认动作"), System.ComponentModel.Category("控件信息"), Description("控件默认动作（AccDefaultAction）")]
        public string AccDefaultAction { get; set; }

        [DisplayName("控件名称"), System.ComponentModel.Category("控件信息"), Description("控件名称（AccName）")]
        public string AccName { get; set; }

        [DisplayName("控件角色"), System.ComponentModel.Category("控件信息"), Description("控件角色（AccRole）")]
        public object AccRole { get; set; }

        [DisplayName("控件状态"), System.ComponentModel.Category("控件信息"), Description("控件状态（AccState）")]
        public object AccState { get; set; }

        [DisplayName("控件值"), System.ComponentModel.Category("控件信息"), Description("控件值（AccValue）")]
        public string AccValue { get; set; }

        [DisplayName("控件快捷键"), System.ComponentModel.Category("控件信息"), Description("控件快捷键（AccKeyboardShortcut）")]
        public string AccKeyboardShortcut { get; set; }

        [DisplayName("控件位置"), System.ComponentModel.Category("控件信息"), Description("控件位置（AccLocation）")]
        public System.Drawing.Rectangle AccLocation { get; set; }
    }
}
