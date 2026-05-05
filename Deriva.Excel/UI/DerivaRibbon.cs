using ExcelDna.Integration.CustomUI;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Deriva.Excel.UI
{
    [ComVisible(true)]
    public sealed class DerivaRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonID)
        {
            return
@"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""DerivaTab"" label=""Deriva"" insertAfterMso=""TabDeveloper"">
        <group id=""DerivaToolsGroup"" label=""Deriva"">
          <button id=""DerivaDashboardButton""
                  label=""Dashboard""
                  size=""large""
                  imageMso=""RefreshAll""
                  onAction=""OnDashboard"" />
          <button id=""DerivaSettingsButton""
                  label=""Settings""
                  size=""large""
                  imageMso=""PropertySheet""
                  onAction=""OnSettings"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnDashboard(IRibbonControl control)
        {
            RunSafely("Unable to open Deriva Dashboard.", DerivaTaskPaneManager.ShowDashboard);
        }

        public void OnSettings(IRibbonControl control)
        {
            RunSafely("Unable to open Deriva Settings.", DerivaTaskPaneManager.ShowSettings);
        }

        private static void RunSafely(string message, Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    message + Environment.NewLine + ex.Message,
                    "Deriva",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
