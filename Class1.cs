using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Text.RegularExpressions;

using System.Windows.Forms;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Application = Microsoft.Office.Interop.Excel.Application;

using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

using Microsoft.Vbe.Interop;
using YamlDotNet.Core;
using YamlDotNet.Core.Tokens;

namespace YAMLConvDNA
{
    public class MyAddin : IExcelAddIn
    {
        Office.CommandBarButton exampleMenuItem;
        Application xlApp = (Application)ExcelDnaUtil.Application;

        // XXX: ここで持っておかないとガベコレされることがある

        Form1 form;

        private Office.CommandBar GetCellContextMenu()
        {
            return this.xlApp.CommandBars["Cell"];
        }

        void exampleMenuItemClick(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var selection = xlApp.Selection;

            if (!(selection is Range))
            {
                return;
            }

            var selectedRange = (Range)selection;
            var values = selectedRange.Value;

            if (values == null || !values.GetType().IsArray)
            {
                return;
            }

            //int firstCol = selectedRange.Column;
            //int firstRow = selectedRange.Row;
            //int numCols = selectedRange.Columns.Count;
            int numRows = selectedRange.Rows.Count;

            if (numRows < 2)
            {
                return;
            }

            var keyValuePairs = tableToKeyValuePairs(values);
            var properties = keyValuePairs[0].Keys;

            if (properties.Count == 0)
            {
                return;
            }

            //string s = "";
            //foreach (var property in properties)
            //{
            //    s += $"{property}\n"; 
            //}
            //MessageBox.Show($"次の列からYAMLを出力します。\n\n{s}");

            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();
            var yaml = serializer.Serialize(keyValuePairs);

            //MessageBox.Show(yaml);

            Clipboard.SetText(yaml);

            form.SetText(yaml);
            form.ShowDialog();

            //selectedRange.Worksheet.Cells[firstRow, firstCol].Value = "foo";
            //MessageBox.Show($"{firstCol}, {firstRow}, {numCols}, {numRows}");
        }

        static List<(int index, int count, string[] identifier)> getPropertiesFromHeader(object[,] values)
        {
            var properties = new List<(int index, int count, string[] identifier)>();
            int i0 = values.GetLowerBound(1);
            int row = values.GetLowerBound(0);

            for (int i = i0, n = i0 + values.GetLength(1); i < n; i++)
            {
                var v = values[row, i];

                if (!(v != null && v is string))
                {
                    continue;
                }

                var s = (string)v;

                int length = 0;

                // 配列は一旦は末尾のみ対応
                const string arrayMark = "[]";

                if (s.EndsWith(arrayMark))
                {
                    // 一旦仮で配列の印をつけておく
                    length = 1;
                    s = s.Substring(0, s.Length - arrayMark.Length);
                }

                string[] identifiers = s.Split('.');

                // \w, \d は全角、日本語とかも含むようなので指定
                if (!identifiers.All(identifier => Regex.IsMatch(identifier, @"^\$?[_a-zA-Z][_a-zA-Z0-9]*$")))
                {
                    continue;
                }

                properties.Add((i, length, identifiers));
            }

            for (int i = 0, n = properties.Count; i < n; i++)
            {
                var property = properties[i];

                if (property.count == 0)
                {
                    continue;
                }

                if (i < n - 1)
                {
                    property.count = properties[i + 1].index - property.index;
                }
                else
                {
                    property.count = 1 + values.GetLength(1) - property.index;
                }

                properties[i] = property;
            }

            return properties;
        }

        static List<Dictionary<string, dynamic>> tableToKeyValuePairs(object[,] values)
        {
            List<Dictionary<string, dynamic>> keyValuePairs = new List<Dictionary<string, dynamic>>();
            var properties = getPropertiesFromHeader(values);
            int i0 = values.GetLowerBound(0);

            for (int i = 1 + i0, n = i0 + values.GetLength(0); i < n; i++)
            {
                Dictionary<string, dynamic> keyValuePair = new Dictionary<string, dynamic>();
                foreach (var property in properties)
                {
                    Dictionary<string, dynamic> kvp = keyValuePair;
                    var pathList = property.identifier.Take(property.identifier.Length - 1);
                    string key = property.identifier.Last();

                    foreach (var path in pathList)
                    {
                        if (!kvp.ContainsKey(path))
                        {
                            kvp.Add(path, new Dictionary<string, dynamic>());
                        }
                        kvp = kvp[path];
                    }

                    if (property.count == 0)
                    {
                        var value = values[i, property.index];

                        kvp.Add(key, value);
                    }
                    else
                    {
                        List<dynamic> valueList = new List<dynamic>();

                        for (int j = property.index; j < property.index + property.count; j++)
                        {
                            valueList.Add(values[i, j]);
                        }
                        kvp.Add(key, valueList.ToList());
                    }
                }
                keyValuePairs.Add(keyValuePair);
            }

            return keyValuePairs;
        }


        void IExcelAddIn.AutoOpen()
        {
            Office.MsoControlType menuItem = Office.MsoControlType.msoControlButton;
            exampleMenuItem = (Office.CommandBarButton)GetCellContextMenu().Controls.Add(menuItem, System.Reflection.Missing.Value, System.Reflection.Missing.Value, 1, true);

            exampleMenuItem.Style = Office.MsoButtonStyle.msoButtonCaption;
            exampleMenuItem.Caption = "to YAML";
            exampleMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(exampleMenuItemClick);

            form = new Form1();
        }

        private void ResetCellMenu()
        {
            GetCellContextMenu().Reset(); // reset the cell context menu back to the default
        }

        void IExcelAddIn.AutoClose()
        {
            ResetCellMenu();

            form.Dispose();
        }
    }

}
