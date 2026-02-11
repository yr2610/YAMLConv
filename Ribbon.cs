using System;
using ExcelDna.Integration.CustomUI;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Reflection;

namespace YAMLConv
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonID)
        {
            string projectName = Assembly.GetExecutingAssembly().GetName().Name;
            return $@"
        <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
          <ribbon>
            <tabs>
              <tab id='tabYamlConv' label='{projectName}'>
                <group id='grpYamlConv' label='YAML'>
                  <button id='btnToYaml'
                          label='YAML出力'
                          size='large'
                          imageMso='ExportTextFile'
                          onAction='OnToYaml'/>
                  <checkBox id='tglGenerateId'
                            label='ID生成'
                            getPressed='GetGenerateId'
                            onAction='OnToggleGenerateId'/>
                  <checkBox id='tglIncludeTsv'
                            label='TSVコメント'
                            getPressed='GetIncludeTsv'
                            onAction='OnToggleIncludeTsv'/>
                  <dropDown id='idLength'
                            label='ID桁'
                            onAction='OnIdLengthChanged'
                            getItemCount='GetIdLengthCount'
                            getItemLabel='GetIdLengthLabel'
                            getSelectedItemIndex='GetIdLengthSelectedIndex'
                            sizeString='00' />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
        }

        public void OnToYaml(IRibbonControl control)
        {
            YamlExporterAddin.Instance.RunToYamlFromRibbon();
        }

        public bool GetGenerateId(IRibbonControl control)
        {
            return YamlExporterAddin.Instance != null && YamlExporterAddin.Instance.GenerateId;
        }

        public void OnToggleGenerateId(IRibbonControl control, bool pressed)
        {
            if (YamlExporterAddin.Instance == null) return;
            YamlExporterAddin.Instance.GenerateId = pressed;
        }

        public bool GetIncludeTsv(IRibbonControl control)
        {
            return YamlExporterAddin.Instance != null && YamlExporterAddin.Instance.IncludeTsvComment;
        }

        public void OnToggleIncludeTsv(IRibbonControl control, bool pressed)
        {
            if (YamlExporterAddin.Instance == null) return;
            YamlExporterAddin.Instance.IncludeTsvComment = pressed;
        }

        private static readonly int[] IdLengths = { 6, 16 };

        public int GetIdLengthCount(Office.IRibbonControl control)
        {
            return IdLengths.Length;
        }

        public string GetIdLengthLabel(Office.IRibbonControl control, int index)
        {
            return IdLengths[index].ToString();
        }

        public int GetIdLengthSelectedIndex(Office.IRibbonControl control)
        {
            if (YamlExporterAddin.Instance == null) return -1;
            return Array.IndexOf(IdLengths, YamlExporterAddin.Instance.IdLength);
        }

        public void OnIdLengthChanged(
            Office.IRibbonControl control,
            string selectedId,
            int selectedIndex)
        {
            if (YamlExporterAddin.Instance == null) return;
            YamlExporterAddin.Instance.IdLength = IdLengths[selectedIndex];
        }
    }
}
