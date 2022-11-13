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
using YamlDotNet.Core.Events;
using YamlDotNet.Serialization.EventEmitters;
using System.IO;

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
            var cellValues = selectedRange.Value;

            if (cellValues == null || !cellValues.GetType().IsArray)
            {
                return;
            }

            var values = MultiDimArrayToJaggedArray(cellValues);

            //int firstCol = selectedRange.Column;
            //int firstRow = selectedRange.Row;
            //int numCols = selectedRange.Columns.Count;
            //int numRows = selectedRange.Rows.Count;

            if (values.Count < 2)
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
                .WithEventEmitter(next => new FlowStyleSequences(next))
                .WithEventEmitter(next => new MultilineScalarFlowStyleEmitter(next))
                //.WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();
            var yaml = serializer.Serialize(keyValuePairs);

            //MessageBox.Show(yaml);

            Clipboard.SetText(yaml);

            form.SetText(yaml);
            form.ShowDialog();

            //selectedRange.Worksheet.Cells[firstRow, firstCol].Value = "foo";
            //MessageBox.Show($"{firstCol}, {firstRow}, {numCols}, {numRows}");
        }

        class FlowStyleSequences : ChainedEventEmitter
        {
            public FlowStyleSequences(IEventEmitter nextEmitter)
                : base(nextEmitter) { }

            public override void Emit(SequenceStartEventInfo eventInfo, IEmitter emitter)
            {
//                if (typeof(IEnumerable<int>).IsAssignableFrom(eventInfo.Source.Type) ||
//                    typeof(IEnumerable<double>).IsAssignableFrom(eventInfo.Source.Type))
//                {
//                    eventInfo = new SequenceStartEventInfo(eventInfo.Source)
//                    {
//                        Style = SequenceStyle.Flow
//                    };
//                }
                if (typeof(IEnumerable<dynamic>).IsAssignableFrom(eventInfo.Source.Type))
                {
                    IEnumerable<dynamic> values = (IEnumerable<dynamic>)eventInfo.Source.Value;

                    // 最初に見つかった null じゃないやつで判定
                    // XXX: 一定の長さ以上の文字列が含まれてなければ、とかでも良いか
                    bool IsFlowStyle()
                    {
                        Type[] types = {
                            typeof(double),
                            typeof(bool),
                            typeof(char),
                        };

                        // 全要素が null or types に代入可能なら flow style
                        return values.All(x => x == null || types.Any(t => t.IsAssignableFrom(x.GetType())));
                    }

                    if (IsFlowStyle())
                    {
                        eventInfo = new SequenceStartEventInfo(eventInfo.Source)
                        {
                            Style = SequenceStyle.Flow
                        };
                    }
                }

                nextEmitter.Emit(eventInfo, emitter);
            }
        }
        public class MultilineScalarFlowStyleEmitter : ChainedEventEmitter
        {
            public MultilineScalarFlowStyleEmitter(IEventEmitter nextEmitter)
                : base(nextEmitter) { }

            public override void Emit(ScalarEventInfo eventInfo, IEmitter emitter)
            {

                if (typeof(string).IsAssignableFrom(eventInfo.Source.Type))
                {
                    string value = eventInfo.Source.Value as string;
                    if (!string.IsNullOrEmpty(value))
                    {
                        bool isMultiLine = value.IndexOfAny(new char[] { '\r', '\n', '\x85', '\x2028', '\x2029' }) >= 0;
                        if (isMultiLine)
                            eventInfo = new ScalarEventInfo(eventInfo.Source)
                            {
                                Style = ScalarStyle.Literal
                            };
                    }
                }

                nextEmitter.Emit(eventInfo, emitter);
            }
        }

        const string idIdentifier = "$id";
        const string basePropertyMark = "*";

        static IEnumerable<(int index, int count, string[] identifier)> getPropertiesFromHeader(IEnumerable<dynamic> values, out int idIndex, out int baseIndex)
        {
            int? _baseIndex = null;
            var properties = new List<(int index, int count, string[] identifier)>();

            idIndex = -1;

            // $id 以外の最初のプロパティ
            int? firstPropertyIndex = null;

            for (int i = 0, n = values.Count(); i < n; i++)
            {
                var v = values.ElementAt(i);

                if (!(v != null && v is string))
                {
                    continue;
                }

                var s = (string)v;

                int count = 0;

                if (s == idIdentifier)
                {
                    idIndex = i;
                    properties.Add((i, count, new string[] { s }));
                    continue;
                }

                // 配列は一旦は末尾のみ対応
                const string arrayMark = "[]";

                if (s.EndsWith(arrayMark))
                {
                    // 一旦仮で配列の印をつけておく
                    count = 1;
                    s = s.Substring(0, s.Length - arrayMark.Length);
                }

                bool marked = s.StartsWith(basePropertyMark);

                if (marked)
                {
                    // マーク削除
                    s = s.Substring(1);
                }

                string[] identifiers = s.Split('.');

                // \w, \d は全角、日本語とかも含むようなので指定
                if (!identifiers.All(identifier => Regex.IsMatch(identifier, @"^[_a-zA-Z][_a-zA-Z0-9]*$")))
                {
                    continue;
                }

                if (_baseIndex == null)
                {
                    if (marked)
                    {
                        _baseIndex = i;
                    }
                    else if (firstPropertyIndex == null)
                    {
                        firstPropertyIndex = i;
                    }
                }

                properties.Add((i, count, identifiers));
            }

            // base mark がない場合は先頭
            baseIndex = _baseIndex ?? firstPropertyIndex.Value;

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
                .Reverse()
                .SkipWhile(y => y.All(x => x == null))
                .Reverse();

            return new List<List<dynamic>>(list);
        }

        //static void TrimValues(ref IEnumerable<List<dynamic>> values, ref IEnumerable<(int index, int count, string[] identifier)> properties)
        //{
        //    // ヘッダー行（1行目）の左寄りの空欄の列は不要なので削除
        //    int x0 = values.First().FindIndex(n => n != null);
        //
        //    if (x0 >= 1)
        //    {
        //        values = new List<List<dynamic>>(values.Select(row => new List<dynamic>(row.Skip(x0))));
        //        properties = properties.Select(property => (property.index - x0, property.count, property.identifier));
        //    }
        //}

        // base 列が null の行を削除
        static void DeleteEmptyRow(ref IEnumerable<List<dynamic>> values, int baseIndex)
        {
            values = values.Where(x => x[baseIndex] != null);
        }

        //static int GetIdPropertyIndex(IEnumerable<(int index, int count, string[] identifier)> properties)
        //{
        //    var idProperty = properties.FirstOrDefault(property => property.identifier.Length == 1 && property.identifier[0] == idIdentifier);
        //
        //    return (idProperty.identifier == null) ? -1 : idProperty.index;
        //}

        // $id 以外の先頭のプロパティ
        //static (int index, int count, string[] identifier) GetBaseProperty(IEnumerable<(int index, int count, string[] identifier)> properties)
        //{
        //    var marked = properties.FirstOrDefault(property => property.identifier.First().StartsWith(basePropertyMark));
        //    if (marked.identifier != null)
        //    {
        //        return marked;
        //    }
        //    return properties.FirstOrDefault(property => String.Join(".", property.identifier) != idIdentifier);
        //}

        static void SetId(ref IEnumerable<List<dynamic>> values, IEnumerable<(int index, int count, string[] identifier)> properties, int baseIndex, int idIndex)
        {
            // 入力済みの $id に重複がないか確認
            var duplicates0 = values
                .GroupBy(x => x[idIndex])
                .Where(x => x.Key != null && x.Count() > 1)
                .Select(x => x.Key)
                .ToList();

            if (duplicates0.Count() > 0)
            {
                throw new Exception("$id が重複しています");
            }

            // 配列だとしてもそのまま利用

            List<(List<dynamic> row, string hash, int index)> idTarget = new List<(List<dynamic> row, string hash, int index)>();

            foreach (var row in values)
            {
                // 空欄の場合のみ付与
                if (row[idIndex] != null)
                {
                    continue;
                }

                var baseIdentifier = row[baseIndex];

                // null の行は id 付与しない
                if (baseIdentifier == null)
                {
                    continue;
                }

                var hash = GetHash(baseIdentifier.ToString());

                const int idLength = 4;

                row[idIndex] = hash.Substring(0, idLength);

                idTarget.Add((row: row, hash: hash, index: 0));
            }

            // TODO: 重複しなくなるまで hash.Substring の先頭をずらして取得し直す
            // TODO: 最後まで行っても重複したら例外（「重複を解決できません」）でOK
            // XXX: 一旦重複してたら例外投げとく
            var duplicates = values
                .GroupBy(x => x[idIndex])
                .Where(x => x.Key != null && x.Count() > 1)
                .Select(x => x.Key)
                .ToList();

            if (duplicates.Count() > 0)
            {
                throw new Exception($"生成した $id の重複を解決できません\n\n{duplicates[0]}");
            }
        }

        static string GetHash(string s)
        {
            byte[] data = Encoding.UTF8.GetBytes(s);
            var sha256 = new System.Security.Cryptography.SHA256CryptoServiceProvider();
            byte[] bs = sha256.ComputeHash(data);

            // リソースを解放する
            sha256.Clear();

            string base64 = Convert.ToBase64String(bs);

            // 用途は unique id なのでシンボルに使えない文字を削除
            s = "+/=".ToCharArray().Aggregate(base64, (_s, c) => _s.Replace(c.ToString(), ""));

            return s;
        }

        // $id 列がなければ最終列に追加
        static void AddIdColumn(ref IEnumerable<List<dynamic>> values)
        {
            if (values.First().Contains(idIdentifier))
            {
                return;
            }

            values.First().Add(idIdentifier);
            foreach (dynamic row in values.Skip(1))
            {
                row.Add(null);
            }
        }

        static IEnumerable<Dictionary<string, dynamic>> tableToKeyValuePairs(IEnumerable<List<dynamic>> values)
        {
            AddIdColumn(ref values);

            int idIndex;
            int baseIndex;
            var properties = getPropertiesFromHeader(values.First(), out idIndex, out baseIndex);

            // $id 以外で先頭のプロパティを基にhash値を求める
            //var baseProperty = properties.ElementAt(baseIndex);

            // $id しかない…？
            //if (baseProperty.identifier == null)
            //{
            //    string message = "Baseとなるプロパティがありません。";
            //    throw new Exception(message);
            //}

            values = new List<List<dynamic>>(values.Skip(1));

            DeleteEmptyRow(ref values, baseIndex);

            // base値に重複がないか確認
            var baseDuplicates = values
                .GroupBy(x => x[baseIndex])
                .Where(x => x.Count() > 1)
                .Select(x => x.Key)
                .ToList();

            if (baseDuplicates.Count() > 0)
            {
                throw new Exception($"Base値({String.Join(", ", baseDuplicates)})が重複しています");
            }

            SetId(ref values, properties, baseIndex, idIndex);

            //TrimValues(ref values, ref properties);

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
