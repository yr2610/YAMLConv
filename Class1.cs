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
using System.Drawing.Drawing2D;

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

            List<Dictionary<string, dynamic>> keyValuePairs;
            try
            {
                keyValuePairs = tableToKeyValuePairs(values);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }
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

        static IEnumerable<(int index, int count, string[] identifier)> getPropertiesFromHeader(IEnumerable<dynamic> values)
        {
            var properties = new List<(int index, int count, string[] identifier)>();

            for (int i = 0, n = values.Count(); i < n; i++)
            {
                var v = values.ElementAt(i);

                if (!(v != null && v is string))
                {
                    continue;
                }

                var s = (string)v;

                int count = 0;

                // 配列は一旦は末尾のみ対応
                const string arrayMark = "[]";

                if (s.EndsWith(arrayMark))
                {
                    // 一旦仮で配列の印をつけておく
                    count = 1;
                    s = s.Substring(0, s.Length - arrayMark.Length);
                }

                string[] identifiers = s.Split('.');

                // \w, \d は全角、日本語とかも含むようなので指定
                if (!identifiers.All(identifier => Regex.IsMatch(identifier, @"^\$?[_a-zA-Z][_a-zA-Z0-9]*$")))
                {
                    continue;
                }

                properties.Add((i, count, identifiers));
            }

            // 番兵（末尾が配列の場合用）
            properties.Add((values.Count(), 0, null));

            for (int i = 0, n = properties.Count - 1; i < n; i++)
            {
                var property = properties[i];

                if (property.count == 0)
                {
                    continue;
                }

                // 次のプロパティの直前まで
                property.count = properties[i + 1].index - property.index;

                properties[i] = property;
            }

            // 番兵削除
            properties.RemoveAt(properties.Count - 1);

            return properties;
        }

        // List な jagged array にする
        // ついでに先頭行と末尾行で連続している空行の trim もしてしまう
        static IEnumerable<List<dynamic>> MultiDimArrayToJaggedArray(object[,] a)
        {
            var i = Enumerable.Range(a.GetLowerBound(0),a.GetLength(0));
            var j = Enumerable.Range(a.GetLowerBound(1),a.GetLength(1));
            var list = i.Select(y => new List<dynamic>(j.Select(x => a[y, x])))
                .SkipWhile(y => y.All(x => x == null))
                .TakeWhile(y => y.Any(x => x != null));

            return new List<List<dynamic>>(list);
        }

        static void TrimValues(ref IEnumerable<List<dynamic>> values, ref IEnumerable<(int index, int count, string[] identifier)> properties)
        {
            if (values.Count() == 0)
            {
                return;
            }

            // ヘッダー行（1行目）の左寄りの空欄の列は不要なので削除
            int x0 = values.First().FindIndex(n => n != null);

            if (x0 >= 1)
            {
                values = new List<List<dynamic>>(values.Select(row => new List<dynamic>(row.Skip(x0))));
                properties = properties.Select(property => (property.index - x0, property.count, property.identifier));
            }
        }

        static IEnumerable<Dictionary<string, dynamic>> tableToKeyValuePairs(object[,] valuesArray)
        {
            var values = MultiDimArrayToJaggedArray(valuesArray);

            if (values.Count() == 0)
            {
                return new List<Dictionary<string, dynamic>>();
            }

            var properties = getPropertiesFromHeader(values.First());

            values = new List<List<dynamic>>(values.Skip(1));

            TrimValues(ref values, ref properties);

            var keyValuePairs = new List<Dictionary<string, dynamic>>();

            foreach (var row in values)
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
                        else
                        {
                            var v = kvp[path];

                            if (v == null || !v.GetType().IsGenericType)
                            {
                                var propertyName = new List<string>(pathList);
                                propertyName.Add(key);
                                string message = $"プロパティ名 {String.Join(".", propertyName)} ( {path} ) が重複しています。";
                                throw new Exception(message);
                            }
                        }
                        kvp = kvp[path];
                    }

                    if (kvp.ContainsKey(key))
                    {
                        var propertyName = new List<string>(pathList);
                        propertyName.Add(key);
                        string message = $"プロパティ名 {String.Join(".", propertyName)} が重複しています。";
                        throw new Exception(message);
                    }

                    if (property.count == 0)
                    {
                        var value = row[property.index];

                        kvp.Add(key, value);
                    }
                    else
                    {
                        // 配列として追加
                        var i = Enumerable.Range(property.index, property.count);
                        var array = i.Select(x => row[x]);

                        // 末尾の連続した null を削除
                        array = array
                            .Reverse()
                            .SkipWhile(x => x == null)
                            .Reverse();
                        kvp.Add(key, array);
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
