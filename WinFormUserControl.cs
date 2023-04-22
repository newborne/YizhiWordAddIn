using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace YizhiWordAddIn
{
    public partial class WinFormUserControl : UserControl
    {
        public WinFormUserControl()
        {
            InitializeComponent();
            //新建一个ElementHost在WinForm中引入WPF组件
            ElementHost host = new ElementHost();
            //设置宽度500
            host.Dock = DockStyle.Fill;

            //新建一个WPF组件
            WpfUserControl wpfUserControl = new WpfUserControl();

            //将WPF组件添加到ElementHost中
            host.Child = wpfUserControl;
            //将ElementHost添加到WinForm中
            this.Controls.Add(host);
        }
    }
}
