namespace ComplementoDesigner
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Grupo1 = this.Factory.CreateRibbonGroup();
            this.dropDownProyecto = this.Factory.CreateRibbonDropDown();
            this.txtTimeStamp = this.Factory.CreateRibbonEditBox();
            this.txtRemitente = this.Factory.CreateRibbonEditBox();
            this.Grupo2 = this.Factory.CreateRibbonGroup();
            this.btnGenerarTemplate = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnObtenerDatosMail = this.Factory.CreateRibbonButton();
            this.btnRegistrarMailRecibido = this.Factory.CreateRibbonButton();
            this.btnGuardarMailRecibido = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Grupo1.SuspendLayout();
            this.Grupo2.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Grupo1);
            this.tab1.Groups.Add(this.Grupo2);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Complemento";
            this.tab1.Name = "tab1";
            // 
            // Grupo1
            // 
            this.Grupo1.Items.Add(this.dropDownProyecto);
            this.Grupo1.Items.Add(this.txtTimeStamp);
            this.Grupo1.Items.Add(this.txtRemitente);
            this.Grupo1.Name = "Grupo1";
            // 
            // dropDownProyecto
            // 
            this.dropDownProyecto.Label = "PROYECTO";
            this.dropDownProyecto.Name = "dropDownProyecto";
            // 
            // txtTimeStamp
            // 
            this.txtTimeStamp.Label = "TIMESTAMP";
            this.txtTimeStamp.Name = "txtTimeStamp";
            this.txtTimeStamp.Text = null;
            // 
            // txtRemitente
            // 
            this.txtRemitente.Label = "REMITENTE";
            this.txtRemitente.Name = "txtRemitente";
            this.txtRemitente.Text = null;
            // 
            // Grupo2
            // 
            this.Grupo2.Items.Add(this.btnGenerarTemplate);
            this.Grupo2.Name = "Grupo2";
            // 
            // btnGenerarTemplate
            // 
            this.btnGenerarTemplate.Label = "GENERAR TEMPLATE";
            this.btnGenerarTemplate.Name = "btnGenerarTemplate";
            this.btnGenerarTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGenerarTemplate_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnObtenerDatosMail);
            this.group1.Items.Add(this.btnRegistrarMailRecibido);
            this.group1.Items.Add(this.btnGuardarMailRecibido);
            this.group1.Name = "group1";
            // 
            // btnObtenerDatosMail
            // 
            this.btnObtenerDatosMail.Label = "Obtener datos de email";
            this.btnObtenerDatosMail.Name = "btnObtenerDatosMail";
            this.btnObtenerDatosMail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnObtenerDatosMail_Click);
            // 
            // btnRegistrarMailRecibido
            // 
            this.btnRegistrarMailRecibido.Label = "Registrar mail recibido";
            this.btnRegistrarMailRecibido.Name = "btnRegistrarMailRecibido";
            this.btnRegistrarMailRecibido.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRegistrarMailRecibido_Click);
            // 
            // btnGuardarMailRecibido
            // 
            this.btnGuardarMailRecibido.Label = "Guardar email recibido";
            this.btnGuardarMailRecibido.Name = "btnGuardarMailRecibido";
            this.btnGuardarMailRecibido.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGuardarMailRecibido_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Grupo1.ResumeLayout(false);
            this.Grupo1.PerformLayout();
            this.Grupo2.ResumeLayout(false);
            this.Grupo2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Grupo1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownProyecto;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtTimeStamp;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtRemitente;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Grupo2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGenerarTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnObtenerDatosMail;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRegistrarMailRecibido;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGuardarMailRecibido;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
