
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms.Integration;
using Office = Microsoft.Office.Core;

// TODO:   按照以下步骤启用功能区(XML)项:

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意: 如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。


namespace YizhiWordAddIn
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {

        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("YizhiWordAddIn.Ribbon1.xml");
        }

        #endregion

        #region 功能区回调
        //在此处创建回叫方法。有关添加回叫方法的详细信息，请访问 https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public Bitmap GetImage(IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.logo_filled);

        }
        public void ShowTaskPaneButton_Click(Office.IRibbonControl control)
        {
            // 获取当前文档的所有任务窗格
            CustomTaskPaneCollection taskPanes = Globals.ThisAddIn.CustomTaskPanes;

            // 查找名为 "易知" 的任务窗格
            Microsoft.Office.Tools.CustomTaskPane myTaskPane = taskPanes.FirstOrDefault(c => c.Title == "易知");

            // 如果任务窗格未创建，创建新的任务窗格并显示
            if (myTaskPane == null)
            {
                //引入winformUserControl
                WinFormUserControl hello = new WinFormUserControl();
                myTaskPane = taskPanes.Add(hello, "易知");

                myTaskPane.Width = 520;
                myTaskPane.Visible = true;
                
                

            }
            // 如果任务窗格已创建，切换可见性
            else
            {
                myTaskPane.Visible = !myTaskPane.Visible;
            }
        }


        #endregion

        #region 帮助器

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
