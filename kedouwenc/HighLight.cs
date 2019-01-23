using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace kedouwenc
{
    public partial class HighLight : Form
    {
        public HighLight()
        {
            InitializeComponent();
            base.Opacity = 0.25;// Set the opacity to 75%.
            Ribbon1.lHwndForm = this.Handle;
           
            
           


        }

        protected override CreateParams CreateParams
        {
            get
            {
                const int WS_CAPTION = 0xC00000;
                const int WS_BORDER = 0x800000;
                const int WS_CHILD = 0x40000000;
                const int WS_EX_TOOLWINDOW = 0x00000080;
                const int WS_EX_TRANSPARENT = 0x00000020;
                const int WS_EX_LAYERED = 0x00080000;
                CreateParams CP;
                CP = base.CreateParams;
                CP.Style &= ~WS_CAPTION & ~WS_BORDER | WS_CHILD;
                CP.ExStyle |= WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW;//   '添加扩展风格：1（图形）层叠的2(对鼠标）透明的 3工具窗口式的（插件一般要设为工具窗口以使其图标不在任务栏中显示）
                return CP;
            }
        }




    }
}
