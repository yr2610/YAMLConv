using System;
using ExcelDna.Integration.CustomUI;
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
                          label='to YAML'
                          size='large'
                          imageMso='ExportTextFile'
                          onAction='OnToYaml'/>
                  <checkBox id='tglGenerateId'
                            label='Generate $id'
                            getPressed='GetGenerateId'
                            onAction='OnToggleGenerateId'/>
                  <checkBox id='tglIncludeTsv'
                            label='Include TSV comment'
                            getPressed='GetIncludeTsv'
                            onAction='OnToggleIncludeTsv'/>
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
        }

        public void OnToYaml(IRibbonControl control)
        {
            MyAddin.Instance.RunToYamlFromRibbon();
        }

        public bool GetGenerateId(IRibbonControl control)
        {
            return MyAddin.Instance != null && MyAddin.Instance.GenerateId;
        }

        public void OnToggleGenerateId(IRibbonControl control, bool pressed)
        {
            if (MyAddin.Instance == null) return;
            MyAddin.Instance.GenerateId = pressed;
        }

        public bool GetIncludeTsv(IRibbonControl control)
        {
            return MyAddin.Instance != null && MyAddin.Instance.IncludeTsvComment;
        }

        public void OnToggleIncludeTsv(IRibbonControl control, bool pressed)
        {
            if (MyAddin.Instance == null) return;
            MyAddin.Instance.IncludeTsvComment = pressed;
        }

    }
}
