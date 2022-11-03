using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace YAMLConvDNA
{
    public class MyAddin : IExcelAddIn
    {
        Application xlApp = (Application)ExcelDnaUtil.Application;

        private Office.CommandBar GetCellContextMenu()
        {
            return this.xlApp.CommandBars["Cell"];
        }

        void exampleMenuItemClick(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            System.Windows.Forms.MessageBox.Show("Example Menu Item clicked");
        }

        void IExcelAddIn.AutoOpen()
        {
            Office.MsoControlType menuItem = Office.MsoControlType.msoControlButton;
            Office.CommandBarButton exampleMenuItem = (Office.CommandBarButton)GetCellContextMenu().Controls.Add(menuItem, System.Reflection.Missing.Value, System.Reflection.Missing.Value, 1, true);

            exampleMenuItem.Style = Office.MsoButtonStyle.msoButtonCaption;
            exampleMenuItem.Caption = "Example Menu Item";
            exampleMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(exampleMenuItemClick);
        }

        private void ResetCellMenu()
        {
            GetCellContextMenu().Reset(); // reset the cell context menu back to the default
        }

        void IExcelAddIn.AutoClose()
        {
            ResetCellMenu();
        }
    }

}
