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
using CsvHelper;
using System.Globalization;
using CsvHelper.Configuration;

namespace YAMLConvDNA
{
    public class MyAddin : IExcelAddIn
    {
        public static MyAddin Instance { get; private set; }

        Office.CommandBarButton exampleMenuItem;
        Application xlApp = (Application)ExcelDnaUtil.Application;

        // XXX: ここで持っておかないとガベコレされることがある

        Form1 form;

        private Office.CommandBar GetCellContextMenu()
        {
            return this.xlApp.CommandBars["Cell"];
        }

        static string JaggedArrayToTsv(IEnumerable<List<dynamic>> array2d)
        {
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);
            var config = new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = "\t" };

            using (var csv = new CsvWriter(sw, config))
            {
                foreach (var value in array2d)
                {
                    foreach (var item in value)
                    {
                        csv.WriteField(item);
                    }
                    csv.NextRecord();
                }
            }

            return sw.ToString();
        }
        static string tsvToYamlComment(string tsv)
        {
            // SkipLast が使えないので reverse, skip, reverse する
            return string.Join("\n", tsv.Split('\n').Reverse().Skip(1).Reverse().Select(x => "# " + x));
        }

        string Yaml { get; set; }
        string TsvComment { get; set; }

        private void TsvCommentCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            var checkBox = (System.Windows.Forms.CheckBox)sender;

            string s = "";

            if (checkBox.Checked)
            {
                s += TsvComment + "\n";
            }
            s += Yaml;

            Clipboard.SetText(s);
            form.SetText(s);
        }

        private void ConvertSelectionToYaml()
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

            string tsv = JaggedArrayToTsv(values);
            TsvComment = tsvToYamlComment(tsv);

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
            Yaml = serializer.Serialize(keyValuePairs);

            //MessageBox.Show(yaml);

            string s = TsvComment + "\n" + Yaml;
            Clipboard.SetText(s);
            form.SetText(s);
            form.ShowDialog();

            //selectedRange.Worksheet.Cells[firstRow, firstCol].Value = "foo";
            //MessageBox.Show($"{firstCol}, {firstRow}, {numCols}, {numRows}");
        }

        void exampleMenuItemClick(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ConvertSelectionToYaml();
        }

        public void RunToYamlFromRibbon()
        {
            ConvertSelectionToYaml();
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

        static IEnumerable<(int index, int count, string[] identifier)> getPropertiesFromHeader(
            IEnumerable<dynamic> values,
            out int idIndex,
            out List<int> baseIndices
        )
        {
            var _baseIndices = new List<int>();
            var properties = new List<(int index, int count, string[] identifier)>();

            idIndex = -1;
            baseIndices = _baseIndices;

            // $id 以外の最初の有効プロパティ
            int? firstPropertyIndex = null;

            for (int i = 0, n = values.Count(); i < n; i++)
            {
                var v = values.ElementAt(i);
                if (!(v != null && v is string)) continue;

                var s = (string)v;
                int count = 0;

                // $id は専用列
                if (s == idIdentifier)
                {
                    idIndex = i;
                    properties.Add((i, count, new string[] { s }));
                    continue;
                }

                // 配列列（[] 印）
                const string arrayMark = "[]";
                if (s.EndsWith(arrayMark))
                {
                    count = 1;
                    s = s.Substring(0, s.Length - arrayMark.Length);
                }

                // * は base 指定（複数可）
                bool marked = s.StartsWith(basePropertyMark);
                if (marked) s = s.Substring(1);

                // a.b.c 形式（ASCII 規則は既存どおり）
                string[] identifiers = s.Split('.');
                if (!identifiers.All(identifier => Regex.IsMatch(identifier, @"^[_a-zA-Z][_a-zA-Z0-9]*$")))
                {
                    continue;
                }

                if (marked)
                {
                    _baseIndices.Add(i); // 複数収集
                }
                else if (firstPropertyIndex == null)
                {
                    firstPropertyIndex = i;
                }

                properties.Add((i, count, identifiers));
            }

            //if (firstPropertyIndex == null)
            //{
            //    return properties;
            //}

            // * が一つも無ければ従来どおり先頭の有効列を base に採用
            if (_baseIndices.Count == 0 && firstPropertyIndex != null)
            {
                _baseIndices.Add(firstPropertyIndex.Value);
            }

            // base列をidentifier順に安定ソート
            var identifierMap = properties
                .Where(p => p.identifier != null)
                .ToDictionary(
                    p => p.index,
                    p => string.Join(".", p.identifier)
                );

            _baseIndices.Sort((a, b) =>
            {
                var nameA = identifierMap[a];
                var nameB = identifierMap[b];

                int cmp = string.CompareOrdinal(nameA, nameB);
                if (cmp != 0) return cmp;

                // 同名ヘッダーの保険（ほぼ起きないが美しい）
                return a.CompareTo(b);
            });

            // 配列幅の確定（番兵を使う既存ロジック）
            properties.Add((values.Count(), 0, null));
            for (int i = 0, n = properties.Count - 1; i < n; i++)
            {
                var property = properties[i];
                if (property.count == 0) continue;
                property.count = properties[i + 1].index - property.index;
                properties[i] = property;
            }
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
        static void DeleteEmptyRow(ref IEnumerable<List<dynamic>> values, List<int> baseIndices)
        {
            values = values.Where(x => BuildBaseKey(x, baseIndices) != null);
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

        static void SetId(
            ref IEnumerable<List<dynamic>> values,
            IEnumerable<(int index, int count, string[] identifier)> properties,
            List<int> baseIndices,
            int idIndex
        )
        {
            // 既存 $id の重複チェック（従来どおり）
            var duplicates0 = values
                .GroupBy(x => x[idIndex])
                .Where(x => x.Key != null && x.Count() > 1)
                .Select(x => x.Key)
                .ToList();

            if (duplicates0.Any())
            {
                var idList = string.Join("\n", duplicates0.Select(x => $"* {x}"));
                throw new Exception($"$id が重複しています\n{idList}");
            }

            // $id 未設定の行のみ、合成キーからハッシュ採番
            foreach (var row in values)
            {
                if (row[idIndex] != null) continue;

                var key = BuildBaseKey(row, baseIndices);
                if (key == null) continue; // 空行相当スキップ

                var hash = GetHash(key);
                const int idLength = 6;
                row[idIndex] = hash.Substring(0, idLength);
            }

            // 生成後 $id の重複チェック（可視キーでレポート）
            var duplicates = values
                .GroupBy(x => x[idIndex])
                .Where(x => x.Key != null && x.Count() > 1);

            if (duplicates.Any())
            {
                string s = "";
                const string indentString = " ";
                foreach (var dup in duplicates)
                {
                    s += $"{dup.Key.ToString()}:\n";
                    var keys = string.Join(
                        "\n",
                        dup.Select(x => $"{indentString}- {BuildBaseKeyDisplay(x, baseIndices)}")
                    );
                    s += $"{keys}\n";
                }
                throw new Exception($"生成した $id の重複を解決できません\n{s}");
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

        // 合成キー（長さプレフィックス方式）
        static string BuildBaseKey(List<dynamic> row, IReadOnlyList<int> baseIndices)
        {
            if (baseIndices == null || baseIndices.Count == 0) return null;

            // 値を順に取得（null は空文字に）
            var parts = baseIndices.Select(i => row[i]?.ToString() ?? string.Empty).ToList();

            // 全部が空ならキーなし
            if (parts.All(p => p.Length == 0)) return null;

            // "<len>:<value>" を連結
            var sb = new StringBuilder();
            foreach (var p in parts)
            {
                sb.Append(p.Length);
                sb.Append(':');
                sb.Append(p);
            }
            return sb.ToString();
        }

        // 表示用（ログ／エラーメッセージ用）の可視キー
        static string BuildBaseKeyDisplay(List<dynamic> row, IReadOnlyList<int> baseIndices)
        {
            if (baseIndices == null || baseIndices.Count == 0) return string.Empty;
            var parts = baseIndices.Select(i => row[i]?.ToString() ?? string.Empty);
            return string.Join(" | ", parts);
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
            List<int> baseIndices;

            var properties = getPropertiesFromHeader(values.First(), out idIndex, out baseIndices);

            if (baseIndices == null || baseIndices.Count == 0)
            {
                throw new Exception($"ヘッダー行に有効なプロパティがありません。");
            }

            // $id 以外で先頭のプロパティを基にhash値を求める
            //var baseProperty = properties.ElementAt(baseIndex);

            // $id しかない…？
            //if (baseProperty.identifier == null)
            //{
            //    string message = "Baseとなるプロパティがありません。";
            //    throw new Exception(message);
            //}

            values = new List<List<dynamic>>(values.Skip(1));

            DeleteEmptyRow(ref values, baseIndices);

            // base値に重複がないか確認
            var baseDuplicates = values
                .GroupBy(x => BuildBaseKey(x, baseIndices))
                .Where(x => x.Count() > 1)
                .Select(x => x.Key)
                .ToList();

            if (baseDuplicates.Count() > 0)
            {
                throw new Exception($"Base値({String.Join(", ", baseDuplicates)})が重複しています");
            }

            SetId(ref values, properties, baseIndices, idIndex);

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
            Instance = this;

            Office.MsoControlType menuItem = Office.MsoControlType.msoControlButton;
            exampleMenuItem = (Office.CommandBarButton)GetCellContextMenu().Controls.Add(menuItem, System.Reflection.Missing.Value, System.Reflection.Missing.Value, 1, true);

            exampleMenuItem.Style = Office.MsoButtonStyle.msoButtonCaption;
            exampleMenuItem.Caption = "to YAML";
            exampleMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(exampleMenuItemClick);

            form = new Form1();
            form.TsvCommentCheckBox_CheckedChanged += TsvCommentCheckBox_CheckedChanged;
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
