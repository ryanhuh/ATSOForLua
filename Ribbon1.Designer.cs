
namespace ExcelAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.SkillSpecial = this.Factory.CreateRibbonGroup();
            this.bntImport = this.Factory.CreateRibbonButton();
            this.bntExport = this.Factory.CreateRibbonButton();
            this.bntVerify = this.Factory.CreateRibbonButton();
            this.grItem = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.SkillSpecial.SuspendLayout();
            this.grItem.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.SkillSpecial);
            this.tab1.Groups.Add(this.grItem);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // SkillSpecial
            // 
            this.SkillSpecial.Items.Add(this.bntImport);
            this.SkillSpecial.Items.Add(this.bntExport);
            this.SkillSpecial.Items.Add(this.bntVerify);
            this.SkillSpecial.Label = "SkillSpecial";
            this.SkillSpecial.Name = "SkillSpecial";
            // 
            // bntImport
            // 
            this.bntImport.Label = "Import";
            this.bntImport.Name = "bntImport";
            this.bntImport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bntImport_Click);
            // 
            // bntExport
            // 
            this.bntExport.Label = "Export";
            this.bntExport.Name = "bntExport";
            this.bntExport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bntExport_Click);
            // 
            // bntVerify
            // 
            this.bntVerify.Label = "검증";
            this.bntVerify.Name = "bntVerify";
            this.bntVerify.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bntVerify_Click);
            // 
            // grItem
            // 
            this.grItem.Items.Add(this.button1);
            this.grItem.Items.Add(this.button2);
            this.grItem.Label = "Item";
            this.grItem.Name = "grItem";
            // 
            // button1
            // 
            this.button1.Label = "Import";
            this.button1.Name = "button1";
            // 
            // button2
            // 
            this.button2.Label = "Export";
            this.button2.Name = "button2";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.SkillSpecial.ResumeLayout(false);
            this.SkillSpecial.PerformLayout();
            this.grItem.ResumeLayout(false);
            this.grItem.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SkillSpecial;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bntImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bntExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grItem;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bntVerify;   
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
