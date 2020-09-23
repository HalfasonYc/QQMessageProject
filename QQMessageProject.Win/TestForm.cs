using QQMessageProject.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QQMessageProject.Win
{
    public partial class TestForm : Form
    {
        QQMessageAssistant information = QQMessageAssistant.Empty;
        public TestForm()
        {
            InitializeComponent();
            this.propertyGrid1.SelectedObject = information;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Task.Run(new Action(() => 
            {
                string errorText = string.Empty;
                QQMessageAssistant assistant = QQMessageAssistant.FromInformation(information.QQ, information.SpecifyQQ, information.SpecifyName,
                    information.Name, out errorText);
                if (assistant != null)
                {
                    this.propertyGrid1.Invoke(new MethodInvoker(() =>
                    {
                        this.propertyGrid1.SelectedObject = assistant;
                    }));
                    assistant.SendMessage(QQMessageAssistant.DefaultTestMsg);
                }
                else
                {
                    this.information.LastError = errorText;
                }
                this.propertyGrid1.Invoke(new MethodInvoker(() =>
                {
                    this.propertyGrid1.Refresh();
                }));
            }));
        }
    }
}
