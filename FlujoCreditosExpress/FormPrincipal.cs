using System;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace FlujoCreditosExpress
{
    public partial class FormPrincipal : Form
    {
        public FormPrincipal()
        {
            InitializeComponent();
        }
        #region Variables Globales

        bool isConfigSaved;
        bool isNewSesion;
        bool isCurrentSession;
        bool isIniciadorasOn;
        bool isDistribuidorasOn;
        bool isFirstSim;
        int sesionId;
        int configId;
        string periodo;
        int periodoActual;
        DateTime fecha;
        ArrayList flujoTitulos;
        ArrayList flujoData;
        ArrayList flujoTotales;
        string fontFamily;
        //Variables para el cálculo de flujos totales
        double saldoInicialTotal;
        double apCapitalITotal;
        double ivaInteresITotal;
        double seguroITotal;
        double comisionXAperturaITotal;
        double ingresoPROSAITotal;
        double cobroXplasticoITotal;
        double capRecuperadoITotal;
        double intRecuperadoITotal;
        double totalEntradasTotal;
        double colocacionETotal;
        double comisionesDistETotal;
        double ivaInteresETotal;
        double seguroETotal;
        double gastosFijosPROSAETotal;
        double gastosVarPROSAETotal;
        double gastosFijosZafyETotal;
        double gastosVarZafyETotal;
        double gastosXPublicidadETotal;
        double gastosXOutSourcingETotal;
        double bonosPremiosETotal;
        double retirosETotal;
        double totalSalidasTotal;
        double saldoFinalTotal;
        int clienteId;
        #endregion

        #region Métodos Privados

        /// <summary>
        /// Carga todos la información inicial del programa.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void FormPrincipal_Load(object sender, EventArgs e)
        {
            try
            {
                // TODO: This line of code loads data into the 'flujoDBDataSet.T_Clientes' table. You can move, or remove it, as needed.
                this.t_ClientesTableAdapter.Fill(this.flujoDBDataSet.T_Clientes);

                this.configId = 1;
                this.isConfigSaved = true;                                                          //Bandera para indicar si la configuración esta guardada.
                this.isNewSesion = true;                                                            //Bandera para indicar si la sesión es nueva.
                this.isFirstSim = true;                                                             //Bandera para indicar si es la primera simulación.
                this.isDistribuidorasOn = true;                                                     //Bandera para indicar que se inscribiran distribuidoras.
                this.sesionId = 1;                                                                  //Asigna la primer sesión en caso de no exisitir ninguna anterior.
                this.clienteId = 1;                                                                 //Asigna el primer cliente en caso de no existir ningun cliente anterior.
                this.periodo = "Q";                                                                 //Asigna el tipo de período.
                this.fecha = DateTime.Now;                                                          //Fecha actual.
                this.fontFamily = "Times New Roman";                                                //Fuente general de para la tabla de flujo.
                this.t_SesionesTableAdapter.Fill(this.flujoDBDataSet.T_Sesiones);                   //Carga los datos de la tabla de sesiones.
                this.t_AmortizacionesTableAdapter.Fill(this.flujoDBDataSet.T_Amortizaciones);       //Carga los datos de la tabla de amortizaciones.
                this.t_ConfiguracionesTableAdapter.Fill(this.flujoDBDataSet.T_Configuraciones);     //Carga los datos de la tabla de configuración.
                this.t_ClientesTableAdapter.Fill(this.flujoDBDataSet.T_Clientes);                   //Carga los datos de la tabla de clientes.
                this.LoadSession(sender, e);                                                        //Carga los datos de la sesión.
                this.SaveSession(sender, e);                                                        //Guarda los datos de la sesión.
                this.LoadConfig(sender, e);                                                         //Obtiene los valores por default para la configuración.
                this.LoadToolTips(sender, e);                                                       //Carga los tooltips para los controles.
                this.btnSaveConfig.Enabled = false;                                                 //Deshabilita el botón para guardar la configuración.
                this.btnSaveConfigR.Enabled = false;                                                 //Deshabilita el botón para guardar la configuración.
                this.StartPosition = FormStartPosition.CenterScreen;                                //Posiciona la pantalla en el centro de la pantalla.
                this.saldoInicialTotal = 0;                                                         //Inicializa la variable para el saldo inicial total.
                this.apCapitalITotal = 0;                                                           //Inicializa la variable para el aporte a capital total.
                this.ivaInteresITotal = 0;                                                          //Inicializa la variable para el IVA del interés total.
                this.seguroITotal = 0;                                                              //Inicializa la variable para el seguro total.
                this.comisionXAperturaITotal = 0;                                                   //Inicializa la variable para la comisión por apertura total.
                this.ingresoPROSAITotal = 0;                                                        //Inicializa la variable para el ingreso de PROSA total.
                this.cobroXplasticoITotal = 0;                                                      //Inicializa la variable para el costo por el plastico para créditos nuevos.
                this.capRecuperadoITotal = 0;                                                       //Inicializa la variable para el capital recuperado total.
                this.intRecuperadoITotal = 0;                                                       //Inicializa la variable para el interés recuperado total.
                this.totalEntradasTotal = 0;                                                        //Inicializa la variable para el total de entradas total.
                this.colocacionETotal = 0;                                                          //Inicializa la variable para colocación total.
                this.comisionesDistETotal = 0;                                                      //Inicializa la variable para comisión total.
                this.ivaInteresETotal = 0;                                                          //Inicializa la variable para el IVA del interés total.
                this.seguroETotal = 0;                                                              //Inicializa la variable para el seguro total.
                this.gastosFijosPROSAETotal = 0;                                                    //Inicializa la variable para los gastos fijos de PROSA total.
                this.gastosVarPROSAETotal = 0;                                                      //Inicializa la variable para los gastos variables de PROSA total.
                this.gastosFijosZafyETotal = 0;                                                     //Inicializa la variable para los gastos fijos de Zafy total.
                this.gastosVarZafyETotal = 0;                                                       //Inicializa la variable para los gastos variables de Zafy total.
                this.gastosXPublicidadETotal = 0;                                                   //Inicializa la variable para los gastos por publicidad total.
                this.gastosXOutSourcingETotal = 0;                                                  //Inicializa la variable para los gastos por outsourcing total.
                this.bonosPremiosETotal = 0;                                                        //Inicializa la variable para los bonos y premios total.
                this.retirosETotal = 0;                                                             //Inicializa la variable para retiros total.
                this.totalSalidasTotal = 0;                                                         //Inicializa la variable para total de salidas total.
                this.saldoFinalTotal = 0;                                                           //Inicializa la variable para el saldo final total.

                this.btnIniciarFlujo.Select();                                                      //Selecciona el botón de "Iniciar Flujo".
                Properties.Settings.Default.PeriodoActual = 1;                                      //Asigna 1 al período actual.
                Properties.Settings.Default.CantPeriodos = 1;                                       //Asigna 1 a la cantidad de períodos a procesar.
                Properties.Settings.Default.Capital = 0;                                            //Asigna 0 al capital.
                Properties.Settings.Default.ApCapital = 0;                                          //Asigna 0 a la aportación de capital.
                Properties.Settings.Default.AportacionAcumulada = 0;                                //Asigna 0 a la aportación acumulada.
                Properties.Settings.Default.IsAutomatic = false;                                    //Asigna falso a la bandera para indicar si se corren periodos automáticos.
                Properties.Settings.Default.isColocacionConfigured = false;                         //Asigna falso a la bandera para indicar si la colocación ha sido configurada.
                Properties.Settings.Default.Clientes2Credito = 0;                                   //Asigna 0 al número de clientes de segundo crédito.
                Properties.Settings.Default.ClientesNuevos = 0;                                     //Asigna 0 al número de clientes nuevos.
                Properties.Settings.Default.CarteraTotal = 0;                                       //Asigna 0 a la cartera total.
                Properties.Settings.Default.ColocacionE = 0;                                        //Asigna 0 a la colocación.
                Properties.Settings.Default.CantIniciadoras = 0;                                    //Asigna 0 a la cantidad de iniciadoras.
                Properties.Settings.Default.CantLideresH = 0;                                       //Asigna 0 a la cantidad de líderes Hijos.
                Properties.Settings.Default.CantLideresN = 0;                                       //Asigna 0 a la cantidad de líderes Nietos.
                Properties.Settings.Default.CantMiembros = 0;                                       //Asigna 0 a la cantidad de miembros de célula.
                Properties.Settings.Default.CantMiembrosIni = 0;                                    //Asigna 0 a la cantidad de miembros de célula invitados.
                Properties.Settings.Default.Distribuidoras = 0;                                     //Asigna 0 a la cantidad de distribuidoras.
                Properties.Settings.Default.DistribuidorasAnt = 0;                                  //Asigna 0 a la cantidad de distribuidoras anteriores.
                Properties.Settings.Default.ProsLIAnt = 0;                                          //Asigna 0 a la cantidad de prospectos a líderes iniciadoras.
                Properties.Settings.Default.ProsLHAnt = 0;                                          //Asigna 0 a la cantidad de prospectos a líderes hijos.
                Properties.Settings.Default.ProsLNAnt = 0;                                          //Asigna 0 a la cantidad de prospectos a líderes nietos.
                Properties.Settings.Default.CantMiembrosH = 0;                                      //Asigna 0 a la cantidad de miembros hijos.
                Properties.Settings.Default.CantMiembrosN = 0;                                      //Asigna 0 a la cantidad de miembros nietos.
                Properties.Settings.Default.CantMiembrosB = 0;                                      //Asigna 0 a la cantidad de miembros bisnieto.
                Properties.Settings.Default.HijasP = 0;                                             //Asigna 0 al porcentaje de hijas.
                Properties.Settings.Default.NietasP = 0;                                            //Asigna 0 al porcentaje de nietas.
                Properties.Settings.Default.BisnietasP = 0;                                         //Asigna 0 al porcentaje de bisnietas.
                Properties.Settings.Default.HijasProd = 0;                                          //Asigna 0 al porcentaje de hijas producidas.
                Properties.Settings.Default.NietasProd = 0;                                         //Asigna 0 al porcentaje de nietas producidas.
                Properties.Settings.Default.BisnietasProd = 0;                                      //Asigna 0 al porcentaje de bisnietas producidas.
                Properties.Settings.Default.HijasPerm = 0;                                          //Asigna 0 al porcentaje de permanencia de hijas.
                Properties.Settings.Default.NietasProd = 0;                                         //Asigna 0 al porcentaje de permanencia de nietas.
                Properties.Settings.Default.BisnietasPerm = 0;                                      //Asigna 0 al porcentaje de permanencia de bisnietas.
                Properties.Settings.Default.ClientesMCPerm = 0;                                     //Asigna 0 a la permanencia de miembros de célula.
                Properties.Settings.Default.ClientesMCHPerm = 0;                                    //Asigna 0 a la permanencia de miembros de célula hijas.
                Properties.Settings.Default.ClientesMCNPerm = 0;                                    //Asigna 0 a la permanencia de miembros de célula nietas.
                Properties.Settings.Default.ClientesMCBPerm = 0;                                    //Asigna 0 a la permanencia de miembros de célula bisnietas.
                Properties.Settings.Default.LideresIBonoB = 0;
                Properties.Settings.Default.LideresIBonoC = 0;
                Properties.Settings.Default.LideresIBonoD = 0;
                Properties.Settings.Default.LideresHBonoB = 0;
                Properties.Settings.Default.LideresHBonoC = 0;
                Properties.Settings.Default.LideresHBonoD = 0;
                Properties.Settings.Default.LideresNBonoB = 0;
                Properties.Settings.Default.LideresNBonoC = 0;
                Properties.Settings.Default.LideresNBonoD = 0;
                Properties.Settings.Default.ColocacionAcumuladaMC = 0;
                Properties.Settings.Default.VolumenAcumuladoMC = 0;
                Properties.Settings.Default.Save();                                                 //Guarda la configuracion de la configuración.

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        
        }

        /// <summary>
        /// Carga los datos de la sesión.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void LoadSession(object sender, EventArgs e)
        {
            try
            {
                if (!(flujoDBDataSet.T_Sesiones.Rows.Count == 0))
                {                    
                    FlujoDBDataSet.T_SesionesRow dr = 
                        (FlujoDBDataSet.T_SesionesRow)flujoDBDataSet.T_Sesiones.Rows[flujoDBDataSet.T_Sesiones.Rows.Count - 1];

                    sesionId += int.Parse(dr["Id"].ToString());
                }
                
                lblSesionId.Text = "Sesión Id: " + sesionId.ToString().PadLeft(6, '0');
                lblFecha.Text = "A " + GetDayName(fecha.DayOfWeek) + " " + fecha.Day + " de " + GetMonthName(fecha.Month) + " de " + fecha.Year ;

                if(!(flujoDBDataSet.T_Clientes.Rows.Count == 0))
                {
                    FlujoDBDataSet.T_ClientesRow dr =
                        (FlujoDBDataSet.T_ClientesRow)flujoDBDataSet.T_Clientes.Rows[flujoDBDataSet.T_Clientes.Rows.Count - 1];
                    clienteId += int.Parse(dr["IdCliente"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        /// <summary>
        /// Guarda la configuración de la sesión.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void SaveSession(object sender, EventArgs e)
        {
            try
            {
                if (isNewSesion)
                {
                    FlujoDBDataSet.T_SesionesRow tSesionesRow = flujoDBDataSet.T_Sesiones.NewT_SesionesRow();
                    int fechaS = int.Parse(fecha.Year +
                                         fecha.Month.ToString().PadLeft(2, '0') +
                                         fecha.Day.ToString().PadLeft(2, '0'));

                    tSesionesRow["Id"] = sesionId;
                    tSesionesRow["Fecha"] = fechaS;
                    tSesionesRow["Estatus"] = 1;
                    flujoDBDataSet.T_Sesiones.AddT_SesionesRow(tSesionesRow);

                    int result = t_SesionesTableAdapter.Update(flujoDBDataSet.T_Sesiones);

                    flujoDBDataSet.AcceptChanges(); 

                    isNewSesion = false;

                    Properties.Settings.Default.SessionId = sesionId;

                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        /// <summary>
        /// Carga la configuración del usuario.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void LoadConfig(object sender, EventArgs e)
        {
            try
            {
                this.GetProperties();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
           
        }

        /// <summary>
        /// Activa el botón para guardar la configuración.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void saveConfig(object sender, EventArgs e)
        {
            if(!btnSaveConfig.Enabled)
            {
                btnSaveConfig.Enabled = true;

                this.isConfigSaved = false;
            }
            if (!btnSaveConfigR.Enabled)
            {
                btnSaveConfigR.Enabled = true;

                this.isConfigSaved = false;
            }
        }

        /// <summary>
        /// Carga los tooltips.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void LoadToolTips(object sender, EventArgs e)
        {
            
            ToolTip toolTip1 = new ToolTip();

            toolTip1.AutoPopDelay = 2000;
            toolTip1.InitialDelay = 100;
            toolTip1.ReshowDelay = 1000;
            toolTip1.ShowAlways = true;
            
            toolTip1.SetToolTip(this.btnIniciarFlujo, "Inicia la simulación del flujo de micro créditos");
            toolTip1.SetToolTip(this.btnSiguienteFlujo, "Procesa el siguiente período");
            toolTip1.SetToolTip(this.btnTerminarFlujo, "Termina el flujo");
            toolTip1.SetToolTip(this.btnSaveConfig, "Guarda la configuración");
            toolTip1.SetToolTip(this.btnSaveConfigR, "Guarda la configuración");
            toolTip1.SetToolTip(this.btnLoadDefaults, "Carga los valores por defecto para la configuración");
            toolTip1.SetToolTip(this.btnLoadDefaultsR, "Carga los valores por defecto para la configuración");
            toolTip1.SetToolTip(this.btnIniciadoras, "Activa la opción de crear líderes de célula iniciadoras");
            toolTip1.SetToolTip(this.btnDetalle, "Muestra las estadisticas de la quincena procesada");
        }

        /// <summary>
        /// Guarda los datos de configuración.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            try
            {
                string[] cat = new string[] { "MCP", "MCC", "MCO" };
                TextBox ctrl;
                int porcentajeT;
                int porcentajeDist = 0;
                int porcentajeN1 = 0;
                int porcentajeN2 = 0;
                int porcentajeN3 = 0;
                bool IsPorcentajeOK = true;
                string msjErr = "";
                string control = string.Empty;
                //Verifica el porcentaje de la generación de miembros de célula
                foreach (string item in cat)
                {
                    porcentajeT = 0;
                    for (int i = 0; i < 8; i++)
                    {
                        control = "txtPProbTamCel" + (i + 2) + item;

                        ctrl = (TextBox)this.Controls.Find(control, true)[0];

                        porcentajeT += int.Parse(ctrl.Text);

                        if (porcentajeT > 100)
                        {
                            ctrl.Focus();
                            IsPorcentajeOK = false;
                            msjErr = "Error en el porcentaje de generación de miembros de célula";
                            break;
                        }
                    }
                    if (!IsPorcentajeOK)
                    {
                        break;
                    }
                }

                //Verifica el porcentaje de clientes para el pago de comisiones
                foreach (Control itemM in this.Controls)
                {
                    if(itemM.Name == "tabCContenedor")
                    {
                        foreach(Control itemS1 in itemM.Controls)
                        {
                            if(itemS1.Name == "tabConfigRed")
                            {
                                foreach (Control itemS2 in itemS1.Controls)
                                {
                                    if(itemS2.Name == "gpbComisiones")
                                    {
                                        foreach (Control itemS3 in itemS2.Controls)
                                        {
                                            if (itemS3 is TextBox)
                                            {
                                                if (itemS3.Name.Substring(0, 7) == "txtDist")
                                                {
                                                    porcentajeDist += int.Parse(itemS3.Text);
                                                }
                                                else if (itemS3.Name.Substring(0, 5) == "txtN1")
                                                {
                                                    porcentajeN1 += int.Parse(itemS3.Text);
                                                }
                                                else if (itemS3.Name.Substring(0, 5) == "txtN2")
                                                {
                                                    porcentajeN2 += int.Parse(itemS3.Text);
                                                }
                                                else if (itemS3.Name.Substring(0, 5) == "txtN3")
                                                {
                                                    porcentajeN3 += int.Parse(itemS3.Text);
                                                }
                                            }

                                            if (porcentajeDist > 100 ||
                                                porcentajeN1 > 100 ||
                                                porcentajeN2 > 100 ||
                                                porcentajeN3 > 100)
                                            {
                                                itemS3.Focus();
                                                IsPorcentajeOK = false;
                                                msjErr = "Error en el porcentaje de comisiones";
                                                break;
                                            }
                                        }
                                    }
                                    if (!IsPorcentajeOK)
                                    {
                                        break;
                                    }
                                }
                            }
                            if (!IsPorcentajeOK)
                            {
                                break;
                            }
                        }
                    }
                    if (!IsPorcentajeOK)
                    {
                        break;
                    }
                }

                if (IsPorcentajeOK)
                {
                    this.SetProperties();
                    Properties.Settings.Default.Save();
                    btnSaveConfig.Enabled = false;
                    btnSaveConfigR.Enabled = false;
                    this.isConfigSaved = true;
                }
                else
                {

                    MessageBox.Show("¡No puedes exceder el 100% \n" + msjErr, "Configuración no guardada");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Configuración no guardada");
            }
        }
        
        /// <summary>
        /// Genera el simulador de flujo de efectivo.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void btnGenerarFlujo_Click(object sender, EventArgs e)
        {
            try
            {
                if (isConfigSaved)
                {
                    Button btn = (Button)sender;
                    periodoActual = Properties.Settings.Default.PeriodoActual;

                    if (btn.Name == "btnIniciarFlujo")
                    {
                        if (isFirstSim)
                        {
                            dgvFlujo.Rows.Clear();
                            dgvFlujo.Columns.Clear();
                            dgvFlujo.Visible = false;
                            dgvFlujoT.Visible = false;
                            dgvFlujoP.Visible = false;
                            lblProcesados.Visible = false;
                            btnSiguienteFlujo.Visible = true;
                            btnTerminarFlujo.Visible = true;
                            Properties.Settings.Default.PeriodoActual = 1;
                            Properties.Settings.Default.CantPeriodos = 1;
                            Properties.Settings.Default.IsAutomatic = false;
                            Properties.Settings.Default.IsToFinish = false;
                            Properties.Settings.Default.ClientesNuevos = 0;
                            isNewSesion = false;
                            isCurrentSession = true;
                            if (isNewSesion)
                            {
                                this.FormPrincipal_Load(sender, e);
                            }
                            isFirstSim = false;
                            periodoActual = Properties.Settings.Default.PeriodoActual;
                        }
                        else
                        {
                            if (DialogResult.Yes == MessageBox.Show("¿Deseas iniciar un flujo?",
                            "Nuevo Flujo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                            {
                                t_SesionesTableAdapter.UpdateStatus(sesionId);
                                flujoDBDataSet.AcceptChanges();
                                dgvFlujo.Rows.Clear();
                                dgvFlujo.Columns.Clear();
                                dgvFlujo.Visible = false;
                                dgvFlujoT.Visible = false;
                                dgvFlujoP.Visible = false;
                                lblProcesados.Visible = false;
                                btnSiguienteFlujo.Visible = true;
                                btnTerminarFlujo.Visible = true;
                                Properties.Settings.Default.PeriodoActual = 1;
                                Properties.Settings.Default.CantPeriodos = 1;
                                Properties.Settings.Default.IsAutomatic = false;
                                Properties.Settings.Default.IsToFinish = false;
                                Properties.Settings.Default.ClientesNuevos = 0;
                                isNewSesion = true;
                                isCurrentSession = true;
                                if (isNewSesion)
                                {
                                    this.FormPrincipal_Load(sender, e);
                                }
                                isFirstSim = false;
                                periodoActual = Properties.Settings.Default.PeriodoActual;
                            }
                            else
                            {
                                isCurrentSession = false;
                            }
                        }
                    }
                    if (btn.Name == "btnTerminarFlujo")
                    {
                        Properties.Settings.Default.IsToFinish = true;
                        Properties.Settings.Default.IsAutomatic = true;
                        isCurrentSession = true;
                    }

                    if (btn.Name == "btnSiguienteFlujo")
                    {
                        isCurrentSession = true;
                    }
                    double cp = Properties.Settings.Default.CantPeriodos;
                    
                    Properties.Settings.Default.CantPP = 0;

                    pbProcesando.Visible = true;

                    while (cp > 0)
                    {
                        //Inicio
                        if (Properties.Settings.Default.IsAutomatic)
                        {
                            periodoActual = Properties.Settings.Default.PeriodoActual;
                            Cursor.Current = Cursors.WaitCursor;
                        }
                        
                        if (isCurrentSession)
                        {
                            this.Enabled = false;
                            t_ConfiguracionesTableAdapter.Fill(flujoDBDataSet.T_Configuraciones);
                            //Se produce la cantidad de clientes nuevos
                            this.CargaProduccionCtesConfig(sender, e);
                            //Abre la ventana para la colocación de créditos
                            Form frmColocacion = new FormColocacion();
                            frmColocacion.StartPosition = FormStartPosition.CenterParent;
                            frmColocacion.Text = "Colocación:  " + periodoActual;
                            frmColocacion.ShowDialog();
                            if(Properties.Settings.Default.IsConfigMod)
                            {
                                this.GetProperties();
                                btnSaveConfig_Click(sender, e);
                            }
                            this.Enabled = true;
                            btnSiguienteFlujo.Select();

                            //Se hace el cálculo del capital

                            t_AmortizacionesTableAdapter.Fill(flujoDBDataSet.T_Amortizaciones);

                            DataRow[] dRowList = flujoDBDataSet.T_Amortizaciones.Select("SesionId = " + sesionId
                                + " AND PeriodoInicial <= " + periodoActual
                                + " AND PeriodoFinal >= " + periodoActual);

                            if (dRowList.Length == 0)
                            {
                                isFirstSim = true;
                                Properties.Settings.Default.IsToFinish = false;
                                Properties.Settings.Default.IsAutomatic = false;
                                btnSiguienteFlujo.Visible = false;
                                btnTerminarFlujo.Visible = false;
                                Cursor.Current = Cursors.Arrow;
                                lblProcesados.Text = "| Quincena Final: " + (periodoActual - 1);
                                break;
                            }

                            //Se declaran variables temporales para los datos
                            double clientesNuevos = Properties.Settings.Default.ClientesNuevos;
                            double clientes2Credito = Properties.Settings.Default.Clientes2Credito;
                            double dist2Credito = Properties.Settings.Default.Dist2Credito;
                            double carteraVigente = 0;
                            double carteraVigenteN = 0;
                            double carteraVigenteP = 0;
                            double saldoInicial = Properties.Settings.Default.Capital;
                            double apCapitalI = Properties.Settings.Default.ApCapital;
                            double capRecuperadoI = 0;
                            double capRecuperadoN = 0;
                            double capRecuperadoP = 0;
                            double intRecuperadoI = 0;
                            double ivaInteresI = Properties.Settings.Default.IVAInteresI / 100;
                            double seguroI = Properties.Settings.Default.SeguroI;
                            double comisionXAperturaI = Properties.Settings.Default.ComAperturaI / 100;
                            double ingresoPROSAI = Properties.Settings.Default.IngresoProsaI;
                            double cobroXPlasticoI = Properties.Settings.Default.CobroXPlasticoI;
                            double totalEntradas = 0;
                            double perdidaE = Properties.Settings.Default.PerdidaE / 100;
                            double colocacionE = Properties.Settings.Default.ColocacionE;
                            double colocacionN = Properties.Settings.Default.ColocacionN;
                            double colocacionP = Properties.Settings.Default.ColocacionP;
                            double comisionesE = Properties.Settings.Default.ComisionDistE / 100;
                            double gastosFijosPROSAE = Properties.Settings.Default.GastosFijosPROSAE;
                            double gastosVarPROSAE = Properties.Settings.Default.GastosVarPROSAE;
                            double gastosFijosZafyE = Properties.Settings.Default.GastosFijosZafyE;
                            double gastosVarZafyE = Properties.Settings.Default.GastosVarZafyE / 100;
                            double gastosXPublicidadE = Properties.Settings.Default.GastosXPublicidadE;
                            double gastosXISRE = 0;
                            double bonosPremiosE = Properties.Settings.Default.BonosPremiosE / 100;
                            double retirosE = Properties.Settings.Default.RetirosE / 100;
                            double totalSalidas = 0;
                            double saldoFinal = 0;
                            int numeroPago = 0;
                            int periodoFinal = 0;
                            double clientesDistP = Properties.Settings.Default.ClientesDistP / 100;
                            double capitalProporcionalDist = 0;
                            //Variables para el cálculo de comisiones

                            double pagoPronto = Properties.Settings.Default.PagoPronto / 100;
                            double pagoTiempo = Properties.Settings.Default.PagoTiempo / 100;
                            double pagoTardio1 = Properties.Settings.Default.PagoTardio1 / 100;
                            double pagoTardio2 = Properties.Settings.Default.PagoTardio2 / 100;
                            double pagoTardio3 = Properties.Settings.Default.PagoTardio3 / 100;
                            double pagoTiempoCN1 = Properties.Settings.Default.PagoTiempoCN1 / 100;
                            double pagoTiempoCN2 = Properties.Settings.Default.PagoTiempoCN2 / 100;
                            double pagoTiempoCN3 = Properties.Settings.Default.PagoTiempoCN3 / 100;
                            double pagoTardio1CN1 = Properties.Settings.Default.PagoTardio1CN1 / 100;
                            double pagoTardio1CN2 = Properties.Settings.Default.PagoTardio1CN2 / 100;
                            double pagoTardio1CN3 = Properties.Settings.Default.PagoTardio1CN3 / 100;
                            double pagoTardio2CN1 = Properties.Settings.Default.PagoTardio2CN1 / 100;
                            double pagoTardio2CN2 = Properties.Settings.Default.PagoTardio2CN2 / 100;
                            double pagoTardio2CN3 = Properties.Settings.Default.PagoTardio2CN3 / 100;
                            double pagoTardio3CN1 = Properties.Settings.Default.PagoTardio2CN1 / 100;
                            double pagoTardio3CN2 = Properties.Settings.Default.PagoTardio2CN2 / 100;
                            double pagoTardio3CN3 = Properties.Settings.Default.PagoTardio2CN3 / 100;
                            double porColocacion = Properties.Settings.Default.PorColocacion / 100;
                            double porCumplimiento = Properties.Settings.Default.PorCumplimiento / 100;
                            double porCumplimientoPerc = Properties.Settings.Default.PorCumplimientoPerc / 100;
                            double porCrecimiento = Properties.Settings.Default.PorCrecimiento / 100;
                            double porCompletarCelula = Properties.Settings.Default.PorCompletarCelula / 100;
                            double porColocacionValorMenor = Properties.Settings.Default.PorColocacionValorMenor / 100;
                            double porColocacionMontoMenor = Properties.Settings.Default.PorColocacionMontoMenor * 1000;
                            double porColocacionValorEntre = Properties.Settings.Default.PorColocacionValorEntre / 100;
                            double porColocacionMontoDe = Properties.Settings.Default.PorColocacionMontoDe * 1000;
                            double porColocacionMontoHasta = Properties.Settings.Default.PorColocacionMontoHasta * 1000;
                            double porColocacionValorMayor = Properties.Settings.Default.PorColocacionValorMayor / 100;
                            double porColocacionMontoMayor = Properties.Settings.Default.PorColocacionMontoMayor * 1000;
                            double porVolumenValorMenor = Properties.Settings.Default.PorVolumenValorMenor / 100;
                            double porVolumenVolumMenor = Properties.Settings.Default.PorVolumenVolumenMenor;
                            double porVolumenValorEntre = Properties.Settings.Default.PorVolumenValorEntre / 100;
                            double porVolumenVolumDe = Properties.Settings.Default.PorVolumenVolumenDe;
                            double porVolumenVolumHasta = Properties.Settings.Default.PorVolumenVolumenHasta;
                            double porVolumenValorMayor = Properties.Settings.Default.PorVolumenValorMayor / 100;
                            double porVolumenVolumMayor = Properties.Settings.Default.PorVolumenVolumenMayor;

                            //Porcentaje de miembros
                            double pagoProntoPerc = Properties.Settings.Default.PagoProntoPerc / 100;
                            double pagoTiempoPerc = Properties.Settings.Default.PagoTiempoPerc / 100;
                            double pagoTardio1Perc = Properties.Settings.Default.PagoTardio1Perc / 100;
                            double pagoTardio2Perc = Properties.Settings.Default.PagoTardio2Perc / 100;
                            double pagoTardio3Perc = Properties.Settings.Default.PagoTardio3Perc / 100;
                            double pagoTiempoCN1Perc = Properties.Settings.Default.PagoTiempoCN1Perc / 100;
                            double pagoTiempoCN2Perc = Properties.Settings.Default.PagoTiempoCN2Perc / 100;
                            double pagoTiempoCN3Perc = Properties.Settings.Default.PagoTiempoCN3Perc / 100;
                            double pagoTardio1CN1Perc = Properties.Settings.Default.PagoTardio1CN1Perc / 100;
                            double pagoTardio1CN2Perc = Properties.Settings.Default.PagoTardio1CN2Perc / 100;
                            double pagoTardio1CN3Perc = Properties.Settings.Default.PagoTardio1CN3Perc / 100;
                            double pagoTardio2CN1Perc = Properties.Settings.Default.PagoTardio2CN1Perc / 100;
                            double pagoTardio2CN2Perc = Properties.Settings.Default.PagoTardio2CN2Perc / 100;
                            double pagoTardio2CN3Perc = Properties.Settings.Default.PagoTardio2CN3Perc / 100;
                            double pagoTardio3CN1Perc = Properties.Settings.Default.PagoTardio2CN1Perc / 100;
                            double pagoTardio3CN2Perc = Properties.Settings.Default.PagoTardio2CN2Perc / 100;
                            double pagoTardio3CN3Perc = Properties.Settings.Default.PagoTardio2CN3Perc / 100;
                            //Cantidad de clientes
                            double distribuidoras = Properties.Settings.Default.Distribuidoras;
                            double lideresIniciadoras = Properties.Settings.Default.CantIniciadoras;
                            double lideresIProd = Properties.Settings.Default.LideresIProd;
                            double lideresHijas = Properties.Settings.Default.CantLideresH;
                            double lideresHProd = Properties.Settings.Default.LideresHProd;
                            double lideresNietas = Properties.Settings.Default.CantLideresN;
                            double lideresNProd = Properties.Settings.Default.LideresNProd;
                            double lideresIBonoB = Properties.Settings.Default.LideresIBonoB;
                            double lideresIBonoC = Properties.Settings.Default.LideresIBonoC;
                            double lideresIBonoD = Properties.Settings.Default.LideresIBonoD;
                            double lideresHBonoB = Properties.Settings.Default.LideresHBonoB;
                            double lideresHBonoC = Properties.Settings.Default.LideresHBonoC;
                            double lideresHBonoD = Properties.Settings.Default.LideresHBonoD;
                            double lideresNBonoB = Properties.Settings.Default.LideresNBonoB;
                            double lideresNBonoC = Properties.Settings.Default.LideresNBonoC;
                            double lideresNBonoD = Properties.Settings.Default.LideresNBonoD;
                            double clientesMCT = 0;
                            double mcHT = 0;
                            double mcNT = 0;
                            double mcBT = 0;
                            double capRecuperadoMCN = 0;
                            double capRecHMCN = 0;
                            double capRecNMCN = 0;
                            double capRecBMCN = 0;
                            double capRecuperadoMCP = 0;
                            double capRecHMCP = 0;
                            double capRecNMCP = 0;
                            double capRecBMCP = 0;
                            double colocacionMCN = 0;
                            double colocacionMCP = 0;
                            double mcIProd = Properties.Settings.Default.CantMiembrosIni;
                            double mcProd = Properties.Settings.Default.CantMiembros;
                            double mcHProd = Properties.Settings.Default.CantMiembrosH;
                            double mcNProd = Properties.Settings.Default.CantMiembrosN;
                            double mcBProd = Properties.Settings.Default.CantMiembrosB;
                            double clientesMCPerm = 0;
                            double mcHPerm = 0;
                            double mcNPerm = 0;
                            double mcBPerm = 0;
                            double clientesMCPermT = 0;
                            double mcHPermT = 0;
                            double mcNPermT = 0;
                            double mcBPermT = 0;
                            //Variables para los bonos y comiciones
                            double bonoA = 0;
                            double bonoB = 0;
                            double bonoC = 0;
                            double bonoD = 0;
                            double bonoE = 0;
                            double bonoF = 0;
                            double bonoG = 0;
                            double comisionA1 = 0;
                            double comisionA2 = 0;
                            double comisionA3 = 0;
                            double comisionB = 0;
                            double c3mli = 0;
                            double c6mli = 0;
                            double c9mli = 0;
                            double c3mlh = 0;
                            double c6mlh = 0;
                            double c9mlh = 0;
                            double c3mln = 0;
                            double c6mln = 0;
                            double c9mln = 0;

                            //Obtiene los ingresos
                            foreach (DataRow dRow in dRowList)
                            {
                                periodoFinal = int.Parse(dRow["PeriodoFinal"].ToString());
                                numeroPago = (periodoFinal - periodoActual) + 1;

                                if (int.Parse(dRow["NumeroPagos"].ToString()) == numeroPago)
                                {
                                    capRecuperadoI += double.Parse(dRow["Capital"].ToString());
                                    intRecuperadoI += double.Parse(dRow["Interes"].ToString());
                                    carteraVigente++;

                                    if (int.Parse(dRow["MontoCredito"].ToString()) <= Properties.Settings.Default.MontoCredito02)
                                    {
                                        capRecuperadoN += double.Parse(dRow["Capital"].ToString());
                                        carteraVigenteN++;
                                    }
                                    else
                                    {
                                        capRecuperadoP += double.Parse(dRow["Capital"].ToString());
                                        carteraVigenteP++;
                                    }
                                }
                            }

                            clientesMCT = Math.Round((carteraVigente - distribuidoras) * (Properties.Settings.Default.ClientesMCP / 100));
                            mcHT = Math.Round(clientesMCT * (Properties.Settings.Default.HijasP / 100));
                            mcNT = Math.Round(clientesMCT * (Properties.Settings.Default.NietasP / 100));
                            mcBT = Math.Round(clientesMCT * (Properties.Settings.Default.BisnietasP / 100));
                            clientesMCPerm = Math.Round((clientes2Credito - dist2Credito) * (Properties.Settings.Default.CtesMCPPerm / 100));
                            mcHPerm = Math.Round(clientesMCPerm * (Properties.Settings.Default.HijasPerm / 100));
                            mcNPerm = Math.Round(clientesMCPerm * (Properties.Settings.Default.NietasPerm / 100));
                            mcBPerm = Math.Round(clientesMCPerm * (Properties.Settings.Default.BisnietasPerm / 100));
                            clientesMCPermT = Properties.Settings.Default.ClientesMCPerm;
                            mcHPermT = Properties.Settings.Default.ClientesMCHPerm;
                            mcNPermT = Properties.Settings.Default.ClientesMCNPerm;
                            mcBPermT = Properties.Settings.Default.ClientesMCBPerm;
                            c3mli = (mcHT / 3) < lideresIniciadoras ? 0 :
                                lideresIniciadoras > 0 ? (lideresIniciadoras - (lideresIBonoB + lideresIProd)) : 0;
                            c6mli = (mcHT / 6) < lideresIniciadoras ? 0 :
                                lideresIniciadoras > 0 ? (lideresIniciadoras - (lideresIBonoC + lideresIProd)) : 0;
                            c9mli = (mcHT / 9) < lideresIniciadoras ? 0 :
                                lideresIniciadoras > 0 ? (lideresIniciadoras - (lideresIBonoC + lideresIProd)) : 0;
                            c3mlh = (mcNT / 3) < lideresHijas ? 0 :
                                lideresHijas > 0 ? (lideresHijas - (lideresHBonoB + lideresHProd)) : 0;
                            c6mlh = (mcNT / 6) < lideresHijas ? 0 :
                                lideresHijas > 0 ? (lideresHijas - (lideresHBonoC + lideresHProd)) : 0;
                            c9mlh = (mcNT / 9) < lideresHijas ? 0 :
                                lideresHijas > 0 ? (lideresHijas - (lideresHBonoC + lideresHProd)) : 0;
                            c3mln = (mcBT / 3) < lideresNietas ? 0 :
                                lideresNietas > 0 ? (lideresNietas - (lideresNBonoB + lideresNProd)) : 0;
                            c6mln = (mcBT / 6) < lideresNietas ? 0 :
                                lideresNietas > 0 ? (lideresNietas - (lideresNBonoC + lideresNProd)) : 0;
                            c9mln = (mcBT / 9) < lideresNietas ? 0 :
                                lideresNietas > 0 ? (lideresNietas - (lideresNBonoC + lideresNProd)) : 0;

                            Properties.Settings.Default.LideresIBonoB += c3mli;
                            Properties.Settings.Default.LideresIBonoC += c6mli;
                            Properties.Settings.Default.LideresIBonoD += c9mli;
                            Properties.Settings.Default.LideresHBonoB += c3mlh;
                            Properties.Settings.Default.LideresHBonoC += c6mlh;
                            Properties.Settings.Default.LideresHBonoD += c9mlh;
                            Properties.Settings.Default.LideresNBonoB += c3mln;
                            Properties.Settings.Default.LideresNBonoC += c6mln;
                            Properties.Settings.Default.LideresNBonoD += c9mln;

                            capRecuperadoMCN = carteraVigenteN == 0 ? 0 :
                                ((clientesMCT * capRecuperadoN) / carteraVigenteN) * porCumplimientoPerc;
                            capRecHMCN = clientesMCT == 0 ? 0 :
                                (mcHT * capRecuperadoMCN) / clientesMCT;
                            capRecNMCN = clientesMCT == 0 ? 0 :
                                (mcNT * capRecuperadoMCN) / clientesMCT;
                            capRecBMCN = clientesMCT == 0 ? 0 :
                                (mcBT * capRecuperadoMCN) / clientesMCT;
                            capRecuperadoMCP = carteraVigenteP == 0 ? 0 :
                                ((clientesMCPermT * capRecuperadoP) / carteraVigenteP) * porCumplimientoPerc;
                            capRecHMCP = clientesMCPermT == 0 ? 0 :
                                (mcHPermT * capRecuperadoMCP) / clientesMCPermT;
                            capRecNMCP = clientesMCPermT == 0 ? 0 :
                                (mcNPermT * capRecuperadoMCP) / clientesMCPermT;
                            capRecBMCP = clientesMCPermT == 0 ? 0 :
                                (mcBPermT * capRecuperadoMCP) / clientesMCPermT;
                            colocacionMCN = clientesNuevos == 0 ? 0 :
                                (colocacionN * mcProd) / clientesNuevos;
                            colocacionMCP = clientes2Credito == 0 ? 0 :
                                (colocacionP * clientesMCPerm) / clientes2Credito;

                            capRecuperadoI = capRecuperadoI - (capRecuperadoMCN + capRecuperadoMCP);
                            capRecuperadoI = (capRecuperadoI - (capRecuperadoI * perdidaE)) +
                               (capRecuperadoMCN + capRecuperadoMCP);
                            capRecuperadoN = (capRecuperadoN - (capRecuperadoN * perdidaE)) +
                               capRecuperadoMCN;
                            capRecuperadoP = (capRecuperadoP - (capRecuperadoP * perdidaE)) +
                               capRecuperadoMCP;
                            intRecuperadoI = intRecuperadoI - (intRecuperadoI * perdidaE);
                            ivaInteresI = ivaInteresI * intRecuperadoI;
                            seguroI = seguroI * carteraVigente;
                            comisionXAperturaI = comisionXAperturaI * colocacionE;
                            ingresoPROSAI = ingresoPROSAI * (clientesNuevos + clientes2Credito);
                            cobroXPlasticoI = cobroXPlasticoI * clientesNuevos;
                            
                            //Obtiene los egresos
                            if (rBtnCapitalInteres.Checked)
                            {
                                capitalProporcionalDist = 
                                    (capRecuperadoI + intRecuperadoI + ivaInteresI + seguroI + comisionXAperturaI) * clientesDistP;
                                comisionesE = capitalProporcionalDist * comisionesE;
                            }
                            else if (rBtnCapital.Checked)
                            {
                                capitalProporcionalDist = (capRecuperadoI + seguroI + comisionXAperturaI) * clientesDistP;
                                comisionesE = capitalProporcionalDist * comisionesE;
                            }
                            gastosVarPROSAE = gastosVarPROSAE * clientesNuevos;
                            gastosVarZafyE = gastosVarZafyE * (capRecuperadoI + intRecuperadoI);
                            gastosXISRE = ((gastosXISRE * comisionesE) * 1.16d) + (comisionesE * .16d);
                            bonosPremiosE = intRecuperadoI * bonosPremiosE;
                            retirosE = retirosE * (capRecuperadoI + intRecuperadoI);

                            totalEntradas = 
                                Math.Ceiling(saldoInicial + apCapitalI + ivaInteresI + seguroI + comisionXAperturaI + ingresoPROSAI + cobroXPlasticoI + capRecuperadoI + intRecuperadoI);
                            totalSalidas = 
                                Math.Ceiling(colocacionE + comisionesE + ivaInteresI + seguroI + gastosFijosPROSAE + gastosVarPROSAE + gastosFijosZafyE + gastosVarZafyE + gastosXPublicidadE + bonosPremiosE + retirosE);
                            saldoFinal = Math.Ceiling(totalEntradas - totalSalidas);

                            //Obtiene los bonos y las comisiones
                            
                            bonoA = colocacionMCN * porColocacion;
                            bonoB = mcProd == 0 ? 0 : 
                                ((((c3mli * 3) * colocacionMCN) / mcProd) * porCrecimiento) +
                                ((((c3mlh * 3) * colocacionMCN) / mcProd) * porCrecimiento) +
                                ((((c3mln * 3) * colocacionMCN) / mcProd) * porCrecimiento);
                            bonoC = mcProd == 0 ? 0 :
                                ((((c6mli * 6) * colocacionMCN) / mcProd) * porCrecimiento) +
                                ((((c6mlh * 6) * colocacionMCN) / mcProd) * porCrecimiento) +
                                ((((c6mln * 6) * colocacionMCN) / mcProd) * porCrecimiento);
                            bonoD = mcProd == 0 ? 0 :
                                ((((c9mli * 9) * colocacionMCN) / mcProd) * porCompletarCelula) +
                                ((((c9mlh * 9) * colocacionMCN) / mcProd) * porCompletarCelula) +
                                ((((c9mln * 9) * colocacionMCN) / mcProd) * porCompletarCelula);
                            bonoE = (capRecuperadoMCN + capRecuperadoMCP) * porCumplimiento;

                            if ((periodoActual % 6) == 0)
                            {
                                Properties.Settings.Default.ColocacionAcumuladaMC += (colocacionMCN + colocacionMCP);
                                Properties.Settings.Default.VolumenAcumuladoMC = clientesMCT;
                                double colocacionAcumMC = Properties.Settings.Default.ColocacionAcumuladaMC;
                                double volumenAcumMC = Properties.Settings.Default.VolumenAcumuladoMC;
                                double colocacionAVal = 0;
                                double volumenAVal = 0;
                                double cantLideresCol = 0;
                                double cantLideresVol = 0;

                                //Valida el porcentaje por el volumen de colocación.
                                if (colocacionAcumMC < porColocacionMontoMenor)
                                {
                                    colocacionAVal = porColocacionValorMenor;
                                    cantLideresCol = Math.Truncate(colocacionAcumMC / porColocacionMontoMenor);
                                }
                                if (colocacionAcumMC >= porColocacionMontoDe && colocacionAcumMC <= porColocacionMontoHasta)
                                {
                                    colocacionAVal = porColocacionValorEntre;
                                    cantLideresCol = Math.Truncate(colocacionAcumMC / porColocacionMontoHasta);
                                }
                                if (colocacionAcumMC > porColocacionMontoMayor)
                                {
                                    colocacionAVal = porColocacionValorMayor;
                                    cantLideresCol = Math.Truncate(colocacionAcumMC / porColocacionMontoMayor);
                                }
                                //Valida el porcentaje por el volumen de miembros
                                if (volumenAcumMC < porVolumenVolumMenor)
                                {
                                    volumenAVal = porVolumenValorMenor;
                                    cantLideresVol = Math.Truncate(volumenAcumMC / porVolumenVolumMenor);
                                }
                                if (volumenAcumMC >= porVolumenVolumDe && volumenAcumMC <= porVolumenVolumHasta)
                                {
                                    volumenAVal = porVolumenValorEntre;
                                    cantLideresVol = Math.Truncate(volumenAcumMC / porVolumenVolumHasta);
                                }
                                if (volumenAcumMC > porVolumenVolumMayor)
                                {
                                    volumenAVal = porVolumenValorMayor;
                                    cantLideresVol = Math.Truncate(volumenAcumMC / porVolumenVolumMayor);
                                }

                                bonoF = colocacionAcumMC * colocacionAVal;
                                bonoG = colocacionAcumMC * volumenAVal;
                                Properties.Settings.Default.ColocacionAcumuladaMC = 0;
                                Properties.Settings.Default.VolumenAcumuladoMC = 0;
                            }
                            else
                            {
                                Properties.Settings.Default.ColocacionAcumuladaMC += (colocacionMCN + colocacionMCP);
                                Properties.Settings.Default.VolumenAcumuladoMC = clientesMCT;
                            }

                            comisionA1 =
                                ((capRecHMCN + capRecHMCP) * pagoTiempoCN1Perc) * pagoTiempoCN1 +
                                ((capRecHMCN + capRecHMCP) * pagoTardio1CN1Perc) * pagoTardio1CN1 +
                                ((capRecHMCN + capRecHMCP) * pagoTardio2CN1Perc) * pagoTardio2CN1 +
                                ((capRecHMCN + capRecHMCP) * pagoTardio3CN1Perc) * pagoTardio3CN1;
                            comisionA2 =
                                ((capRecNMCN + capRecNMCP) * pagoTiempoCN2Perc) * pagoTiempoCN2 +
                                ((capRecNMCN + capRecNMCP) * pagoTardio1CN2Perc) * pagoTardio1CN2 +
                                ((capRecNMCN + capRecNMCP) * pagoTardio2CN2Perc) * pagoTardio2CN2 +
                                ((capRecNMCN + capRecNMCP) * pagoTardio3CN2Perc) * pagoTardio3CN2;
                            comisionA3 =
                                ((capRecBMCN + capRecBMCP) * pagoTiempoCN3Perc) * pagoTiempoCN3 +
                                ((capRecBMCN + capRecBMCP) * pagoTardio1CN3Perc) * pagoTardio1CN3 +
                                ((capRecBMCN + capRecBMCP) * pagoTardio2CN3Perc) * pagoTardio2CN3 +
                                ((capRecBMCN + capRecBMCP) * pagoTardio3CN3Perc) * pagoTardio3CN3;

                            comisionB =
                                ((capitalProporcionalDist) * pagoProntoPerc) * pagoPronto +
                                ((capitalProporcionalDist) * pagoTiempoPerc) * pagoTiempo +
                                ((capitalProporcionalDist) * pagoTardio1Perc) * pagoTardio1 +
                                ((capitalProporcionalDist) * pagoTardio2Perc) * pagoTardio2 +
                                ((capitalProporcionalDist) * pagoTardio3Perc) * pagoTardio3;

                            bonosPremiosE += (bonoA + bonoB + bonoC + bonoD + bonoE + bonoF + bonoG);
                            comisionesE += (comisionA1 + comisionA2 + comisionA3 + comisionB);

                            bonosPremiosE = bonosPremiosE > 0 ? bonosPremiosE : 0;
                            comisionesE = comisionesE > 0 ? comisionesE : 0;

                            flujoData = new ArrayList();
                            flujoData.Add(saldoInicial);        //00
                            flujoData.Add(apCapitalI);          //01
                            flujoData.Add(capRecuperadoI);      //02
                            flujoData.Add(intRecuperadoI);      //03
                            flujoData.Add(ivaInteresI);         //04
                            flujoData.Add(seguroI);             //05
                            flujoData.Add(comisionXAperturaI);  //06
                            flujoData.Add(ingresoPROSAI);       //07
                            flujoData.Add(cobroXPlasticoI);     //08
                            flujoData.Add(totalEntradas);       //09
                            flujoData.Add(string.Empty);        //10
                            flujoData.Add(colocacionE);         //11
                            flujoData.Add(comisionesE);         //12
                            flujoData.Add(ivaInteresI);         //13
                            flujoData.Add(seguroI);             //14
                            flujoData.Add(gastosFijosPROSAE);   //15
                            flujoData.Add(gastosVarPROSAE);     //16
                            flujoData.Add(gastosFijosZafyE);    //17
                            flujoData.Add(gastosVarZafyE);      //18
                            flujoData.Add(gastosXPublicidadE);  //19
                            flujoData.Add(gastosXISRE); //20
                            flujoData.Add(bonosPremiosE);       //21
                            flujoData.Add(retirosE);            //22
                            flujoData.Add(totalSalidas);        //23
                            flujoData.Add(string.Empty);        //24
                            flujoData.Add(saldoFinal);          //25

                            
                            //Llena la tabla con los titulos
                            if (periodoActual == 1 && dgvFlujoT.Rows.Count == 0)
                            {
                                dgvFlujoT.Columns.Add("Conceptos", "Conceptos");
                                dgvFlujoT.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                                dgvFlujoT.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                                dgvFlujoT.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                dgvFlujoT.Columns[0].HeaderCell.Style.Font = new Font(fontFamily, 10f, FontStyle.Bold);
                                dgvFlujoT.Columns[0].Width = 100;
                                this.GetFlujoDataTitles();
                                dgvFlujoT.Rows.Add(flujoTitulos.Count);

                                for (int ir = 0; ir < flujoTitulos.Count; ir++)
                                {
                                    dgvFlujoT.Rows[ir].Cells[0].Value =
                                        ir == 8 || ir == 22 ? flujoTitulos[ir].ToString().PadLeft(25, ' ') : flujoTitulos[ir];
                                    dgvFlujoT.Rows[ir].Cells[0].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                                    dgvFlujoT.Rows[ir].Cells[0].Style.Font = new Font(fontFamily, 10f,
                                        ir == 8 || ir == 22 ? FontStyle.Underline : ir == 9 || ir == 23 || ir == 25 ? FontStyle.Bold : FontStyle.Regular);
                                    dgvFlujoT.Rows[ir].Cells[0].ReadOnly = true;
                                    dgvFlujoT.Rows[ir].Height = ir == 10 || ir == 24 ? 5 : 20;
                                }
                            }

                            //Llena la tabla con los datos
                            dgvFlujo.Columns.Add(periodo + periodoActual, periodo + periodoActual);
                            dgvFlujo.Columns[periodo + periodoActual].SortMode = DataGridViewColumnSortMode.NotSortable;
                            dgvFlujo.Columns[periodo + periodoActual].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                            dgvFlujo.Columns[periodo + periodoActual].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dgvFlujo.Columns[periodo + periodoActual].HeaderCell.Style.Font = new Font(fontFamily, 10f, FontStyle.Bold);
                            dgvFlujo.Columns[periodo + periodoActual].Width = 100;

                            if (periodoActual == 1)
                            {
                                dgvFlujo.Rows.Add(flujoData.Count);
                            }

                            //Aquí se llenan los registros de las columnas para cada quincena
                            for (int ir = 0; ir < flujoData.Count; ir++)
                            {
                                dgvFlujo.Rows[ir].Cells[periodo + periodoActual].Value =
                                    ir == 8 || ir == 22 ? string.Format("{0:C0}", flujoData[ir]).PadLeft(15, ' ') : string.Format("{0:C0}", flujoData[ir]);
                                dgvFlujo.Rows[ir].Cells[periodo + periodoActual].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                                dgvFlujo.Rows[ir].Cells[periodo + periodoActual].Style.Font =
                                    new Font(fontFamily, 10f, ir == 8 || ir == 22 ? FontStyle.Underline : ir == 9 || ir == 23 || ir == 25 ? FontStyle.Bold : FontStyle.Regular);
                                dgvFlujo.Rows[ir].Cells[periodo + periodoActual].Style.ForeColor = ir == 25 ?
                                    (double.Parse(flujoData[ir].ToString()) < 0 ? Color.DarkRed : Color.DarkGreen) : Color.Black;
                                dgvFlujo.Rows[ir].Cells[periodo + periodoActual].ReadOnly = true;
                                dgvFlujo.Rows[ir].Height = ir == 10 || ir == 24 ? 5 : 20;
                            }

                            //Cálcula los totales
                            if (dRowList.Length != 0)
                            {
                                this.saldoInicialTotal += saldoInicial;
                                this.apCapitalITotal += apCapitalI;
                                this.capRecuperadoITotal += capRecuperadoI;
                                this.intRecuperadoITotal += intRecuperadoI;
                                this.ivaInteresITotal += ivaInteresI;
                                this.seguroITotal += seguroI;
                                this.comisionXAperturaITotal += comisionXAperturaI;
                                this.ingresoPROSAITotal += ingresoPROSAI;
                                this.cobroXplasticoITotal += cobroXPlasticoI;
                                this.totalEntradasTotal += totalEntradas;
                                this.colocacionETotal += colocacionE;
                                this.comisionesDistETotal += comisionesE;
                                this.ivaInteresETotal += ivaInteresI;
                                this.seguroETotal += seguroETotal;
                                this.gastosFijosPROSAETotal += gastosFijosPROSAE;
                                this.gastosVarPROSAETotal += gastosVarPROSAE;
                                this.gastosFijosZafyETotal += gastosFijosZafyE;
                                this.gastosVarZafyETotal += gastosVarZafyE;
                                this.gastosXPublicidadETotal += gastosXPublicidadE;
                                this.gastosXOutSourcingETotal += gastosXISRE;
                                this.bonosPremiosETotal += bonosPremiosE;
                                this.retirosETotal += retirosE;
                                this.totalSalidasTotal += totalSalidas;
                                this.saldoFinalTotal += saldoFinal;
                            }

                            flujoTotales = new ArrayList();
                            flujoTotales.Add(this.saldoInicialTotal);       //00
                            flujoTotales.Add(this.apCapitalITotal);         //01    
                            flujoTotales.Add(this.capRecuperadoITotal);     //02
                            flujoTotales.Add(this.intRecuperadoITotal);     //03
                            flujoTotales.Add(this.ivaInteresITotal);        //04
                            flujoTotales.Add(this.seguroITotal);            //05
                            flujoTotales.Add(this.comisionXAperturaITotal); //06
                            flujoTotales.Add(this.ingresoPROSAITotal);      //07
                            flujoTotales.Add(this.cobroXplasticoITotal);    //08
                            flujoTotales.Add(this.totalEntradasTotal);      //09
                            flujoTotales.Add(string.Empty);                 //10
                            flujoTotales.Add(this.colocacionETotal);        //11
                            flujoTotales.Add(this.comisionesDistETotal);    //12
                            flujoTotales.Add(this.ivaInteresETotal);        //13
                            flujoTotales.Add(this.seguroETotal);            //14
                            flujoTotales.Add(this.gastosFijosPROSAETotal);  //15
                            flujoTotales.Add(this.gastosVarPROSAETotal);    //16
                            flujoTotales.Add(this.gastosFijosZafyETotal);   //17
                            flujoTotales.Add(this.gastosVarZafyETotal);     //18
                            flujoTotales.Add(this.gastosXPublicidadETotal); //19
                            flujoTotales.Add(this.gastosXOutSourcingETotal);//20
                            flujoTotales.Add(this.bonosPremiosETotal);      //21
                            flujoTotales.Add(this.retirosETotal);           //22
                            flujoTotales.Add(this.totalSalidasTotal);       //23
                            flujoTotales.Add(string.Empty);                 //24
                            flujoTotales.Add(this.saldoFinalTotal);         //25

                            if (periodoActual == 1 && dgvFlujoP.Rows.Count == 0)
                            {
                                dgvFlujoP.Columns.Add("Totales", "Totales");
                                dgvFlujoP.Rows.Add(flujoTotales.Count);
                            }

                            dgvFlujoP.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                            dgvFlujoP.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                            dgvFlujoP.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dgvFlujoP.Columns[0].HeaderCell.Style.Font = new Font(fontFamily, 10f, FontStyle.Bold);


                            for (int ir = 0; ir < flujoTotales.Count; ir++)
                            {
                                dgvFlujoP.Rows[ir].Cells[0].Value =
                                     ir == 8 || ir == 22 ? string.Format("{0:C0}", flujoTotales[ir]).PadLeft(15, ' ') : string.Format("{0:C0}", flujoTotales[ir]);
                                dgvFlujoP.Rows[ir].Cells[0].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                                dgvFlujoP.Rows[ir].Cells[0].Style.Font =
                                    new Font(fontFamily, 10f, ir == 8 || ir == 22 ? FontStyle.Underline : ir == 9 || ir == 23 || ir == 25 ? FontStyle.Bold : FontStyle.Regular);
                                dgvFlujoP.Rows[ir].Cells[0].Style.ForeColor = ir == 25 ?
                                    (double.Parse(flujoTotales[ir].ToString()) < 0 ? Color.DarkRed : Color.DarkGreen) : Color.Black;
                                dgvFlujoP.Rows[ir].Cells[0].ReadOnly = true;
                                dgvFlujoP.Rows[ir].Height = ir == 10 || ir == 24 ? 5 : 20;
                            }

                            //Cálcula el tamaño del grid
                            int Width = 20;

                            for (int i = 0; i < dgvFlujo.Columns.Count; i++)
                            {
                                Width += dgvFlujo.Columns[i].Width;
                            }

                            dgvFlujoT.Height = 26;
                            dgvFlujo.Height = 26;
                            dgvFlujoP.Height = 26;

                            if (dgvFlujo.Width < Width)
                            {
                                dgvFlujo.Height = 44;
                                dgvFlujo.FirstDisplayedScrollingColumnIndex = periodoActual - 1;
                            }

                            for (int i = 0; i < dgvFlujo.Rows.Count; i++)
                            {
                                dgvFlujoT.Height += dgvFlujo.Rows[i].Height;
                                dgvFlujo.Height += dgvFlujo.Rows[i].Height;
                                dgvFlujoP.Height += dgvFlujo.Rows[i].Height;

                            }

                            //Guarda las propiedades
                            Properties.Settings.Default.PeriodoActual = periodoActual + 1;
                            Properties.Settings.Default.Capital = saldoFinal;
                            Properties.Settings.Default.AportacionAcumulada = apCapitalITotal;
                            Properties.Settings.Default.ApCapital = 0;
                            Properties.Settings.Default.CarteraTotal = carteraVigente;

                            //Muestra los controles ocultos
                            dgvFlujoT.Visible = true;
                            dgvFlujo.Visible = true;
                            dgvFlujoP.Visible = true;
                            lblProcesados.Visible = true;
                            lblProcesados.Text = "| Quincena procesada: " + periodoActual + "   |   Quincena siguiente: " + (periodoActual + 1);
                            
                        }
                        if (!Properties.Settings.Default.IsToFinish)
                        {
                            cp = Properties.Settings.Default.CantPeriodos;
                            Properties.Settings.Default.CantPP++;
                            cp = cp - Properties.Settings.Default.CantPP;
                        }


                        if (cp == 0)
                        {
                            Properties.Settings.Default.CantPeriodos = 1;
                            Properties.Settings.Default.IsAutomatic = false;
                            Cursor.Current = Cursors.Arrow;
                        }
                        //Fin
                    }
                    pbProcesando.Visible = false;
                }
                else
                {
                    if (DialogResult.Yes == MessageBox.Show("¡La configuración no ha sido guardada! \n\n \t ¿Desea guardarla ahora?",
                        "Error en configuración", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                    {
                        btnSaveConfig_Click(sender, e);
                        btnGenerarFlujo_Click(sender, e);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
                this.Enabled = true;
            }
            
        }

        /// <summary>
        /// Carga la configuración de la producción de clientes.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void CargaProduccionCtesConfig(object sender, EventArgs e)
        {
            try
            {
                double incrementoD = 0;         //Incremento Distribuidoras
                double incrementoMM = 0;        //Incremento Medios Masivos
                double incrementoCZ = 0;        //Incremento Clientes Zafy
                double permanenciaD = 0;        //Permanencia Distribuidoras
                double permanenciaMM = 0;       //Permanencia Medios Masivos
                double permanenciaCZ = 0;       //Permanencia Clientes Zafy
                double permanenciaMC = 0;       //Permanencia de miembros de célula
                double cantD = 0;               //Cantidad de distribuidoras
                double cantXD = 0;               //Cantidad de clientes por distribuidoras
                double credD = 0;               //Créditos otorgados a distribuidoras
                double cantidadD = 0;           //Cantidad total de créditos a distribuidoras
                double cantMM = 0;              //Cantidad de clientes por medios masivos
                double prosMM = 0;              //Prospectos de clientes de medios masivos
                double credMM = 0;              //Créditos otorgados a clientes de medios masivos
                double cantidadMM = 0;          //Cantidad total de créditos a medios masivos
                double cantCZ = 0;              //Cantidad de clientes Zafy
                double prosCZ = 0;              //Prospectos de clientes de Zafy
                double credCZ = 0;              //Créditos otorgados a clientes Zafy
                double cantidadCZ = 0;          //Cantidad total de créditos a clientes Zafy
                double cantCtes = 0;            //Cantidad total de los tres canales
                double prosLC = 0;              //Prospectos de líderes de célula
                double lideLC = 0;              //Líderes de célula
                double cantidadLC = 0;          //Cantidad total de líderes de célula
                double probTamCel2MC = 0;       //Probabilidad de tamaño de célula de 2 miembros
                double probTamCel3MC = 0;       //Probabilidad de tamaño de célula de 3 miembros
                double probTamCel4MC = 0;       //Probabilidad de tamaño de célula de 4 miembros
                double probTamCel5MC = 0;       //Probabilidad de tamaño de célula de 5 miembros
                double probTamCel6MC = 0;       //Probabilidad de tamaño de célula de 6 miembros
                double probTamCel7MC = 0;       //Probabilidad de tamaño de célula de 7 miembros
                double probTamCel8MC = 0;       //Probabilidad de tamaño de célula de 8 miembros
                double probTamCel9MC = 0;       //Probabilidad de tamaño de célula de 9 miembros
                double cantidadMCH = 0;         //Cantidad total de miembros de célula hijas
                double cantidadMCN = 0;         //Cantidad total de miembros de célula nietas
                double cantidadMCB = 0;         //Cantidad total de miembros de célula bisnietas    

                //Canal de Distribuidoras
                if (cmbEscD.Text == "Pesimista")
                {
                    incrementoD = Properties.Settings.Default.IncrementoDP / 100;
                    permanenciaD = Properties.Settings.Default.PermanenciaDP / 100;

                    cantD = Properties.Settings.Default.CantDP;
                    cantXD = Properties.Settings.Default.CantXDP;
                    credD = Properties.Settings.Default.CredDP / 100;
                }
                else if (cmbEscD.Text == "Conservador")
                {
                    incrementoD = Properties.Settings.Default.IncrementoDC / 100;
                    permanenciaD = Properties.Settings.Default.PermanenciaDC / 100;

                    cantD = Properties.Settings.Default.CantDC;
                    cantXD = Properties.Settings.Default.CantXDC;
                    credD = Properties.Settings.Default.CredDC / 100;
                }
                else if (cmbEscD.Text == "Oprimista")
                {
                    incrementoD = Properties.Settings.Default.IncrementoDO / 100;
                    permanenciaD = Properties.Settings.Default.PermanenciaDO / 100;

                    cantD = Properties.Settings.Default.CantDO;
                    cantXD = Properties.Settings.Default.CantXDO;
                    credD = Properties.Settings.Default.CredDO / 100;
                }

                //Canal de Medios Masivos
                if (cmbEscMM.Text == "Pesimista")
                {
                    incrementoMM = Properties.Settings.Default.IncrementoMMP / 100;
                    permanenciaMM = Properties.Settings.Default.PermanenciaMMP / 100;

                    cantMM = Properties.Settings.Default.CantMMP;
                    prosMM = Properties.Settings.Default.ProsMMP / 100;
                    credMM = Properties.Settings.Default.CredMMP / 100;
                }
                else if (cmbEscMM.Text == "Conservador")
                {
                    incrementoMM = Properties.Settings.Default.IncrementoMMC / 100;
                    permanenciaMM = Properties.Settings.Default.PermanenciaMMC / 100;

                    cantMM = Properties.Settings.Default.CantMMC;
                    prosMM = Properties.Settings.Default.ProsMMC / 100;
                    credMM = Properties.Settings.Default.CredMMC / 100;
                }
                else if (cmbEscMM.Text == "Oprimista")
                {
                    incrementoMM = Properties.Settings.Default.IncrementoMMO / 100;
                    permanenciaMM = Properties.Settings.Default.PermanenciaMMO / 100;

                    cantMM = Properties.Settings.Default.CantMMO;
                    prosMM = Properties.Settings.Default.ProsMMO / 100;
                    credMM = Properties.Settings.Default.CredMMO / 100;
                }

                //Canal de Clientes Zafy
                if (cmbEscCZ.Text == "Pesimista")
                {
                    incrementoCZ = Properties.Settings.Default.IncrementoCZP / 100;
                    permanenciaCZ = Properties.Settings.Default.PermanenciaCZP / 100;

                    cantCZ = Properties.Settings.Default.CantCZP;
                    prosCZ = Properties.Settings.Default.ProsCZP / 100;
                    credCZ = Properties.Settings.Default.CredCZP / 100;
                }
                else if (cmbEscCZ.Text == "Conservador")
                {
                    incrementoCZ = Properties.Settings.Default.IncrementoCZC / 100;
                    permanenciaCZ = Properties.Settings.Default.PermanenciaCZC / 100;

                    cantCZ = Properties.Settings.Default.CantCZC;
                    prosCZ = Properties.Settings.Default.ProsCZC / 100;
                    credCZ = Properties.Settings.Default.CredCZC / 100;
                }
                else if (cmbEscCZ.Text == "Optimista")
                {
                    incrementoCZ = Properties.Settings.Default.IncrementoCZO / 100;
                    permanenciaCZ = Properties.Settings.Default.PermanenciaCZO / 100;

                    cantCZ = Properties.Settings.Default.CantCZO;
                    prosCZ = Properties.Settings.Default.ProsCZO / 100;
                    credCZ = Properties.Settings.Default.CredCZO / 100;
                }

                //Líderes y miembros de célula

                if (cmbEscLC.Text == "Pesimista")
                {
                    permanenciaMC = Properties.Settings.Default.PermanenciaLCP / 100;

                    prosLC = Properties.Settings.Default.ProsLCP / 100;
                    lideLC = Properties.Settings.Default.LideLCP / 100;

                    probTamCel2MC = Properties.Settings.Default.ProbTamCel2MCP / 100;
                    probTamCel3MC = Properties.Settings.Default.ProbTamCel3MCP / 100;
                    probTamCel4MC = Properties.Settings.Default.ProbTamCel4MCP / 100;
                    probTamCel5MC = Properties.Settings.Default.ProbTamCel5MCP / 100;
                    probTamCel6MC = Properties.Settings.Default.ProbTamCel6MCP / 100;
                    probTamCel7MC = Properties.Settings.Default.ProbTamCel7MCP / 100;
                    probTamCel8MC = Properties.Settings.Default.ProbTamCel8MCP / 100;
                    probTamCel9MC = Properties.Settings.Default.ProbTamCel9MCP / 100;
                }
                else if (cmbEscLC.Text == "Conservador")
                {
                    permanenciaMC = Properties.Settings.Default.PermanenciaLCC / 100;

                    prosLC = Properties.Settings.Default.ProsLCC / 100;
                    lideLC = Properties.Settings.Default.LideLCC / 100;

                    probTamCel2MC = Properties.Settings.Default.ProbTamCel2MCC / 100;
                    probTamCel3MC = Properties.Settings.Default.ProbTamCel3MCC / 100;
                    probTamCel4MC = Properties.Settings.Default.ProbTamCel4MCC / 100;
                    probTamCel5MC = Properties.Settings.Default.ProbTamCel5MCC / 100;
                    probTamCel6MC = Properties.Settings.Default.ProbTamCel6MCC / 100;
                    probTamCel7MC = Properties.Settings.Default.ProbTamCel7MCC / 100;
                    probTamCel8MC = Properties.Settings.Default.ProbTamCel8MCC / 100;
                    probTamCel9MC = Properties.Settings.Default.ProbTamCel9MCC / 100;
                }
                else if (cmbEscLC.Text == "Optimista")
                {
                    permanenciaMC = Properties.Settings.Default.PermanenciaLCO / 100;

                    prosLC = Properties.Settings.Default.ProsLCO / 100;
                    lideLC = Properties.Settings.Default.LideLCO / 100;

                    probTamCel2MC = Properties.Settings.Default.ProbTamCel2MCO / 100;
                    probTamCel3MC = Properties.Settings.Default.ProbTamCel3MCO / 100;
                    probTamCel4MC = Properties.Settings.Default.ProbTamCel4MCO / 100;
                    probTamCel5MC = Properties.Settings.Default.ProbTamCel5MCO / 100;
                    probTamCel6MC = Properties.Settings.Default.ProbTamCel6MCO / 100;
                    probTamCel7MC = Properties.Settings.Default.ProbTamCel7MCO / 100;
                    probTamCel8MC = Properties.Settings.Default.ProbTamCel8MCO / 100;
                    probTamCel9MC = Properties.Settings.Default.ProbTamCel9MCO / 100;
                }

                if (!isDistribuidorasOn)
                {
                    cantD = Properties.Settings.Default.Distribuidoras;
                }

                cantidadD = cantD * cantXD * credD;
                cantidadMM = cantMM * prosMM * credMM;
                cantidadCZ = cantCZ * prosCZ * credCZ;
                cantCtes = cantidadD + cantidadMM + cantidadCZ;

                Properties.Settings.Default.IsDistribuidorasOn = isDistribuidorasOn;

                if (isDistribuidorasOn)
                {
                    cantCtes += cantD;
                }

                if (Properties.Settings.Default.ClientesNuevos > 0)
                {
                    cantCtes = Properties.Settings.Default.ClientesNuevos;
                }

                if (Properties.Settings.Default.DistribuidorasAnt == 0)
                {
                    Properties.Settings.Default.Distribuidoras = cantD;
                    Properties.Settings.Default.DistribuidorasAnt = cantD;
                }
                //Obtiene la cantidad de iniciadoras y lideres.

                double clientesMCP = Properties.Settings.Default.ClientesMCP / 100;
                double creditos2QMM = 0;
                double creditos2QCZ = 0;
                double creditos2QMC = 0;
                double creditos2QMCH = 0;
                double creditos2QMCN = 0;
                double creditos2QMCB = 0;
                double creditos2Q = 0;
                double creditos2QD = 0;
                double creditos4QMM = 0;
                double creditos4QCZ = 0;
                double creditos4QMC = 0;
                double creditos4QMCH = 0;
                double creditos4QMCN = 0;
                double creditos4QMCB = 0;
                double creditos4Q = 0;
                double creditos4QD = 0;
                double creditosPros = 0;
                double clientesMMP2CT = 0;
                double clientesZP2CT = 0;
                double clientesMCP2CT = 0;
                double clientesMCHP2C = 0;
                double clientesMCNP2C = 0;
                double clientesMCBP2C = 0;
                double creditosMCHPerdidos = 0;
                double creditosMCNPerdidos = 0;
                double creditosMCBPerdidos = 0;

                double cantidadMiembros = 0;
                double cantidadMiembrosH = 0;
                double cantidadMiembrosN = 0;
                double cantidadLideresMCH = 0;
                double cantidadLideresMCN = 0;
                double cantMiembrosProd = 0;
                //Comienza la generación de miembros y líderes de célula.

                //Obtiene la cantidad de prospectos a líderes o miembros de célula
                if (flujoDBDataSet.T_Configuraciones.Rows.Count > 0)
                {
                    DataRow[] drL;
                    if (periodoActual >= 3)
                    {
                        drL = flujoDBDataSet.T_Configuraciones.Select(
                        " SesionId = " + Properties.Settings.Default.SessionId +
                        " AND Campo = 'CD'" +
                        " AND TipoDato = 'N" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos2QD = double.Parse(dr["Valor"].ToString());
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                        " SesionId = " + Properties.Settings.Default.SessionId +
                        " AND Campo = 'CtesMMP'" +
                        " AND TipoDato = 'P" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMMP2CT = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesZP'" +
                            " AND TipoDato = 'P" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesZP2CT = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMC'" +
                            " AND TipoDato = 'P" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCP2CT = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                        " SesionId = " + Properties.Settings.Default.SessionId +
                        " AND Campo = 'CtesMCHP'" +
                        " AND TipoDato = 'P" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCHP2C = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMCNP'" +
                            " AND TipoDato = 'P" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCNP2C = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMCBP'" +
                            " AND TipoDato = 'P" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCBP2C = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo LIKE '%Q02'" +
                            " AND TipoDato = 'N" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos2Q += double.Parse(dr["Valor"].ToString());
                        }

                        if (creditos2Q > 0)
                        {
                            creditos2QMM = Math.Round(((creditos2Q - creditos2QD) * clientesMMP2CT) * permanenciaMM);
                            creditos2QCZ = Math.Round(((creditos2Q - creditos2QD) * clientesZP2CT) * permanenciaCZ);
                            creditos2QMC = Math.Round(((creditos2Q - creditos2QD) * clientesMCP2CT));
                            creditos2QMCH = Math.Round((creditos2QMC * clientesMCHP2C) * permanenciaMC);
                            creditos2QMCN = Math.Round((creditos2QMC * clientesMCNP2C) * permanenciaMC);
                            creditos2QMCB = Math.Round((creditos2QMC * clientesMCBP2C) * permanenciaMC);
                            creditosMCHPerdidos += Math.Round((creditos2QMC * clientesMCHP2C) - creditos2QMCH);
                            creditosMCNPerdidos += Math.Round((creditos2QMC * clientesMCNP2C) - creditos2QMCN);
                            creditosMCBPerdidos += Math.Round((creditos2QMC * clientesMCBP2C) - creditos2QMCB);
                            creditos2Q = Math.Round(creditos2QMM + creditos2QCZ);
                        }
                    }

                    if (periodoActual >= 5)
                    {
                        drL = flujoDBDataSet.T_Configuraciones.Select(
                        " SesionId = " + Properties.Settings.Default.SessionId +
                        " AND Campo = 'CD'" +
                        " AND TipoDato = 'N" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos4QD = double.Parse(dr["Valor"].ToString());
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                        " SesionId = " + Properties.Settings.Default.SessionId +
                        " AND Campo = 'CtesMMP'" +
                        " AND TipoDato = 'P" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMMP2CT = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesZP'" +
                            " AND TipoDato = 'P" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesZP2CT = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMC'" +
                            " AND TipoDato = 'P" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCP2CT = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMCHP'" +
                            " AND TipoDato = 'P" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCHP2C = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMCNP'" +
                            " AND TipoDato = 'P" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCNP2C = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMCBP'" +
                            " AND TipoDato = 'P" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCBP2C = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo LIKE '%Q04'" +
                            " AND TipoDato = 'N" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos4Q += double.Parse(dr["Valor"].ToString());
                        }

                        if (creditos4Q > 0)
                        {
                            creditos4QMM = Math.Round(((creditos4Q - creditos4QD) * clientesMMP2CT) * permanenciaMM);
                            creditos4QCZ = Math.Round(((creditos4Q - creditos4QD) * clientesZP2CT) * permanenciaCZ);
                            creditos4QMC = Math.Round(((creditos4Q - creditos4QD) * clientesMCP2CT));
                            creditos4QMCH = Math.Round((creditos4QMC * clientesMCHP2C) * permanenciaMC);
                            creditos4QMCN = Math.Round((creditos4QMC * clientesMCNP2C) * permanenciaMC);
                            creditos4QMCB = Math.Round((creditos4QMC * clientesMCBP2C) * permanenciaMC);
                            creditosMCHPerdidos += Math.Round((creditos4QMC * clientesMCHP2C) - creditos4QMCH);
                            creditosMCNPerdidos += Math.Round((creditos4QMC * clientesMCNP2C) - creditos4QMCN);
                            creditosMCBPerdidos += Math.Round((creditos4QMC * clientesMCBP2C) - creditos4QMCB);
                            creditos4Q = Math.Round(creditos4QMM + creditos4QCZ);
                        }
                    }
                }

                //Obtiene los prospectos totales
                creditosPros = Math.Round(creditos2Q + creditos4Q) + 
                    Properties.Settings.Default.ProsLIAnt;

                if (creditosPros > 0)
                {
                    //Crea líderes iniciadoras y/o miembros de célula
                    if (isIniciadorasOn)
                    {
                        //Cálcula la creación de líderes iniciadoras
                        cantidadLC = Math.Truncate(creditosPros * prosLC * lideLC);
                        Properties.Settings.Default.CantIniciadoras += cantidadLC;
                        Properties.Settings.Default.ProsLIAnt = creditosPros - cantidadLC;
                        Properties.Settings.Default.LideresIProd = cantidadLC;
                    }
                    else
                    {
                        //cálcula la creación de miembros de célula
                        cantidadMiembros = Math.Truncate(creditosPros * prosLC * lideLC);
                        Properties.Settings.Default.ProsLIAnt = creditosPros - cantidadMiembros;
                        Properties.Settings.Default.CantMiembrosIni = cantidadMiembros;
                    }

                    //Genera miembros de célula hijos
                    cantidadMCH += (Properties.Settings.Default.CantIniciadoras * probTamCel2MC) * 2;
                    cantidadMCH += (Properties.Settings.Default.CantIniciadoras * probTamCel3MC) * 3;
                    cantidadMCH += (Properties.Settings.Default.CantIniciadoras * probTamCel4MC) * 4;
                    cantidadMCH += (Properties.Settings.Default.CantIniciadoras * probTamCel5MC) * 5;
                    cantidadMCH += (Properties.Settings.Default.CantIniciadoras * probTamCel6MC) * 6;
                    cantidadMCH += (Properties.Settings.Default.CantIniciadoras * probTamCel7MC) * 7;
                    cantidadMCH += (Properties.Settings.Default.CantIniciadoras * probTamCel8MC) * 8;
                    cantidadMCH += (Properties.Settings.Default.CantIniciadoras * probTamCel9MC) * 9;
                    cantidadMCH = Math.Truncate(cantidadMCH);

                    //Cálcula la creación de líderes hijos
                    cantidadMiembrosH = Math.Round(creditos2QMCH + creditos4QMCH) +
                            Properties.Settings.Default.ProsLHAnt;
                    cantidadLideresMCH = Math.Truncate(cantidadMiembrosH * prosLC * lideLC);
                    Properties.Settings.Default.CantLideresH += cantidadLideresMCH;
                    Properties.Settings.Default.ProsLHAnt = cantidadMiembrosH - cantidadLideresMCH;
                    Properties.Settings.Default.LideresHProd = cantidadLideresMCH;
                    //Genera miembros de célula nietos
                    cantidadMCN += (Properties.Settings.Default.CantLideresH * probTamCel2MC) * 2;
                    cantidadMCN += (Properties.Settings.Default.CantLideresH * probTamCel3MC) * 3;
                    cantidadMCN += (Properties.Settings.Default.CantLideresH * probTamCel4MC) * 4;
                    cantidadMCN += (Properties.Settings.Default.CantLideresH * probTamCel5MC) * 5;
                    cantidadMCN += (Properties.Settings.Default.CantLideresH * probTamCel6MC) * 6;
                    cantidadMCN += (Properties.Settings.Default.CantLideresH * probTamCel7MC) * 7;
                    cantidadMCN += (Properties.Settings.Default.CantLideresH * probTamCel8MC) * 8;
                    cantidadMCN += (Properties.Settings.Default.CantLideresH * probTamCel9MC) * 9;
                    cantidadMCN = Math.Truncate(cantidadMCN);

                    //Cálcula la creación de líderes nietos
                    cantidadMiembrosN = Math.Round(creditos2QMCN + creditos4QMCN) +
                            Properties.Settings.Default.ProsLNAnt;
                    cantidadLideresMCN = Math.Truncate(cantidadMiembrosN * prosLC * lideLC);
                    Properties.Settings.Default.CantLideresN += cantidadLideresMCN;
                    Properties.Settings.Default.ProsLNAnt = cantidadMiembrosN - cantidadLideresMCN;
                    Properties.Settings.Default.LideresNProd = cantidadLideresMCN;
                    //Genera miembros de célula bisnietos
                    cantidadMCB += (Properties.Settings.Default.CantLideresN * probTamCel2MC) * 2;
                    cantidadMCB += (Properties.Settings.Default.CantLideresN * probTamCel3MC) * 3;
                    cantidadMCB += (Properties.Settings.Default.CantLideresN * probTamCel4MC) * 4;
                    cantidadMCB += (Properties.Settings.Default.CantLideresN * probTamCel5MC) * 5;
                    cantidadMCB += (Properties.Settings.Default.CantLideresN * probTamCel6MC) * 6;
                    cantidadMCB += (Properties.Settings.Default.CantLideresN * probTamCel7MC) * 7;
                    cantidadMCB += (Properties.Settings.Default.CantLideresN * probTamCel8MC) * 8;
                    cantidadMCB += (Properties.Settings.Default.CantLideresN * probTamCel9MC) * 9;
                    cantidadMCB = Math.Truncate(cantidadMCB);

                    /*************************************************************************************
                    *******Se reparten los miembros producidos cuando se apaga el botón de iniciadoras****
                    *************************************************************************************/

                    if(!isIniciadorasOn && cantidadMiembros > 0)
                    {
                        //Define los límites de Madres, hijas, nietas y bisnietas de la célula
                        double madresL = Properties.Settings.Default.CantIniciadoras;
                        double hijasL = madresL * 9;
                        double nietasL = hijasL * 9;
                        double bisnietasL = nietasL * 9;
                        double miembrosCL = hijasL + nietasL + bisnietasL;

                        //Obtiene la dimensión actual de la red
                        double carteraT = Properties.Settings.Default.CarteraTotal;
                        double distribuidoras = Properties.Settings.Default.Distribuidoras;
                        double clientesMCCT = Properties.Settings.Default.ClientesMCP / 100;
                        double ctesMC = Math.Round((carteraT - distribuidoras) * clientesMCCT);
                        double hijasC = Math.Round((ctesMC * (Properties.Settings.Default.HijasP / 100))) - creditosMCHPerdidos;
                        double nietasC = Math.Round((ctesMC * (Properties.Settings.Default.NietasP / 100))) - creditosMCNPerdidos;
                        double bisnietasC = Math.Round((ctesMC * (Properties.Settings.Default.BisnietasP / 100))) - creditosMCBPerdidos;

                        //Valida que no se sobrepasen el máximo de los miembros para cada nivel y reparte los miembros producidos
                        double hijasT = hijasC + cantidadMCH;
                        double nietasT = nietasC + cantidadMCN;
                        double bisnietasT = bisnietasC + cantidadMCB;

                        if(hijasT < hijasL)
                        {
                            if((hijasT + cantidadMiembros) <= hijasL)
                            {
                                cantidadMCH += cantidadMiembros;
                                cantidadMiembros = 0;
                            }
                             else
                            {
                                cantidadMCH += (hijasL - hijasT);
                                cantidadMiembros = (cantidadMiembros - (hijasL - hijasT));
                            }
                        }
                        else
                        {
                            cantidadMiembros += (cantidadMCH - (hijasL - hijasC));
                            cantidadMCH = (hijasL - hijasC);
                        }

                        if (nietasT < nietasL)
                        {
                            if ((nietasT + cantidadMiembros) <= nietasL)
                            {
                                cantidadMCN += cantidadMiembros;
                                cantidadMiembros = 0;
                            }
                            else
                            {
                                cantidadMCN += (nietasL - nietasT);
                                cantidadMiembros = (cantidadMiembros - (nietasL - nietasT));
                            }
                        }
                        else
                        {
                            cantidadMiembros += (cantidadMCN - (nietasL - nietasC));
                            cantidadMCN = (nietasL - nietasC);
                        }

                        if (bisnietasT < bisnietasL)
                        {
                            if ((bisnietasT + cantidadMiembros) <= bisnietasL)
                            {
                                cantidadMCB += cantidadMiembros;
                                cantidadMiembros = 0;
                            }
                            else
                            {
                                cantidadMCB += (bisnietasL - bisnietasT);
                                cantidadMiembros = (cantidadMiembros - (bisnietasL - bisnietasT));
                            }
                        }
                        else
                        {
                            cantidadMiembros += (cantidadMCB - (bisnietasL - bisnietasC));
                            cantidadMCB = (bisnietasL - bisnietasC);
                        }

                        if (cantidadMiembros > 0)
                        {
                            Properties.Settings.Default.ProsLIAnt += cantidadMiembros;
                        }

                        //Guarda los valores para el detalle
                        Properties.Settings.Default.MadresC = madresL;
                        Properties.Settings.Default.HijasC = hijasL;
                        Properties.Settings.Default.NietasC = nietasL;
                        Properties.Settings.Default.BisnietasC = bisnietasL;
                        Properties.Settings.Default.MiembrosC = miembrosCL;

                    }

                    //Obtiene el total de miembros producidos para este período
                    cantMiembrosProd = cantidadMCH + cantidadMCN + cantidadMCB;
                }

                //Guarda la configuración por canal de incremento, permanencia, clientes nuevos y miembros de célula;
                Properties.Settings.Default.IncrementoDVal = incrementoD;
                Properties.Settings.Default.IncrementoMMVal = incrementoMM;
                Properties.Settings.Default.IncrementoCZVal = incrementoCZ;
                Properties.Settings.Default.PermanenciaDVal = permanenciaD;
                Properties.Settings.Default.PermanenciaMMVal = permanenciaMM;
                Properties.Settings.Default.PermanenciaCZVal = permanenciaCZ;
                Properties.Settings.Default.PermanenciaMCVal = permanenciaMC;
                Properties.Settings.Default.ClientesNuevos = cantCtes;
                Properties.Settings.Default.CantMiembros = cantMiembrosProd;
                Properties.Settings.Default.CantMiembrosH = cantidadMCH;
                Properties.Settings.Default.CantMiembrosN = cantidadMCN;
                Properties.Settings.Default.CantMiembrosB = cantidadMCB;
                Properties.Settings.Default.ClientesXDist = cantXD;
                Properties.Settings.Default.CreditosXDistP = credD;

                //Guarda las proporciones de hijas, nietas y bisnietas que se producen.
                if (cantMiembrosProd > 0)
                {
                    Properties.Settings.Default.HijasProd = (cantidadMCH / cantMiembrosProd) * 100;
                    Properties.Settings.Default.NietasProd = (cantidadMCN / cantMiembrosProd) * 100;
                    Properties.Settings.Default.BisnietasProd = (cantidadMCB / cantMiembrosProd) * 100;

                    //Obtiene el último Id de configuración
                    FlujoDBDataSet.T_ConfiguracionesRow tcrId;
                    FlujoDBDataSet.T_ConfiguracionesRow tConfiguracionesRow;
                    if (flujoDBDataSet.T_Configuraciones.Rows.Count > 0)
                    {
                        tcrId =
                        (FlujoDBDataSet.T_ConfiguracionesRow)flujoDBDataSet.T_Configuraciones.Rows[
                            flujoDBDataSet.T_Configuraciones.Rows.Count - 1];

                        configId = int.Parse(tcrId["Id"].ToString()) + 1;

                    }
                    /*Guarda la proporciones de miembros de célula producidos*/
                    //Clientes miembros de célula Hijos
                    tConfiguracionesRow = flujoDBDataSet.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesMCHP";
                    tConfiguracionesRow["Valor"] = 
                        Properties.Settings.Default.HijasP > 0 ? Properties.Settings.Default.HijasP : 0;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    //Clientes miembros de célula Nietos
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesMCNP";
                    tConfiguracionesRow["Valor"] = 
                        Properties.Settings.Default.NietasP > 0 ? Properties.Settings.Default.NietasP : 0;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    //Clientes miembros de célula Bisnietos
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesMCBP";
                    tConfiguracionesRow["Valor"] = 
                        Properties.Settings.Default.BisnietasP > 0 ? Properties.Settings.Default.BisnietasP : 0;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    //Guarda todos los registros en la base de datos
                    int result = 0;

                    foreach (FlujoDBDataSet.T_ConfiguracionesRow dr in flujoDBDataSet.T_Configuraciones.Rows)
                    {
                        result = t_ConfiguracionesTableAdapter.UpdateTConfiguracion(
                            int.Parse(dr["Id"].ToString()), int.Parse(dr["SesionId"].ToString()),
                            dr["Campo"].ToString(), dr["Valor"].ToString(), dr["TipoDato"].ToString(), 
                            short.Parse(dr["Estatus"].ToString()));

                        if (result == 0)
                        {
                            result = t_ConfiguracionesTableAdapter.InsertTConfiguracion(
                            int.Parse(dr["Id"].ToString()), int.Parse(dr["SesionId"].ToString()),
                            dr["Campo"].ToString(), dr["Valor"].ToString(), dr["TipoDato"].ToString(), 
                            short.Parse(dr["Estatus"].ToString()));
                        }
                        result = 0;
                    }
                    flujoDBDataSet.AcceptChanges();
                }
                //Guarda las proporciones de clientes producidos para el primero período.
                if (periodoActual == 1)
                {
                    Properties.Settings.Default.CtesDistPProd = (cantidadD / (cantCtes - cantD)) * 100;
                    Properties.Settings.Default.CtesMMPProd = (cantidadMM / (cantCtes - cantD)) * 100;
                    Properties.Settings.Default.CtesCZPProd = (cantidadCZ / (cantCtes - cantD)) * 100;
                    Properties.Settings.Default.CtesMCPProd = (cantidadMCH / (cantCtes - cantD)) * 100;
                    Properties.Settings.Default.ClientesDistP = (cantidadD / (cantCtes - cantD)) * 100;
                    Properties.Settings.Default.ClientesMMP = (cantidadMM / (cantCtes - cantD)) * 100;
                    Properties.Settings.Default.ClientesZafyP = (cantidadCZ / (cantCtes - cantD)) * 100;
                    Properties.Settings.Default.ClientesMCP = (cantidadMCH / (cantCtes - cantD)) * 100;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), 
                    "Error en la producción de clientes");
            }
        }

        /// <summary>
        /// Obtiene los titulos de los renglones
        /// </summary>
        private void GetFlujoDataTitles()
        {
            flujoTitulos = new ArrayList();
            flujoTitulos.Add("+ Saldo Inicial");            //00
            flujoTitulos.Add("+ Aportación");               //01
            flujoTitulos.Add("+ Capital Recuperado");       //02
            flujoTitulos.Add("+ Interés Recuperado");       //03
            flujoTitulos.Add("+ IVA del Interés");          //04
            flujoTitulos.Add("+ Seguro");                   //05
            flujoTitulos.Add("+ Comisión por Apertura");    //06
            flujoTitulos.Add("+ Ingreso PROSA");            //07
            flujoTitulos.Add("+ Cobro por Plastico");       //08
            flujoTitulos.Add("Total Entradas =");           //09
            flujoTitulos.Add(string.Empty);                 //10
            flujoTitulos.Add("- Colocación");               //11
            flujoTitulos.Add("- Comisiones");               //12
            flujoTitulos.Add("- IVA del Interés");          //13
            flujoTitulos.Add("- Seguro");                   //14
            flujoTitulos.Add("- Gastos Fijos PROSA");       //15
            flujoTitulos.Add("- Gastos Var PROSA");         //16
            flujoTitulos.Add("- Gastos Fijos Zafy");        //17
            flujoTitulos.Add("- Gastos Var Zafy");          //18
            flujoTitulos.Add("- Gastos por Publicidad");    //19
            flujoTitulos.Add("- Gastos por ISR");           //20
            flujoTitulos.Add("- Bonos y Premios");          //21           
            flujoTitulos.Add("- Retiros");                  //22
            flujoTitulos.Add("Total Salidas =");            //23
            flujoTitulos.Add(string.Empty);                 //24
            flujoTitulos.Add("Saldo Final");                //25
        }
        #endregion

        #region Métodos de Validación

        /// <summary>
        /// Valida que sean números
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        public void checkNumbers(object sender, KeyPressEventArgs e)
        {
            TextBox txtB = (TextBox)sender;
            if (!char.IsControl(e.KeyChar) && 
                !char.IsDigit(e.KeyChar) && 
                e.KeyChar != 46)
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// Valida que no este vacio el campo y que no exceda el 100%.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        public void checkEmpty(object sender, KeyEventArgs e)
        {
            if(sender is TextBox)
            {
                TextBox txtB = (TextBox)sender;
                if (txtB.Text == "")
                {
                    ValidaCampo(this, txtB);
                }
            }
            else if (sender is NumericUpDown)
            {
                NumericUpDown txtB = (NumericUpDown)sender;
                if (txtB.Text == "")
                {
                    txtB.Text = "4";
                }
            }
        }

        /// <summary>
        /// Valida que el campo no esté vacío.
        /// </summary>
        /// <param name="ctrl">El control a validar</param>
        /// <param name="txtB">El campo de texto a validar</param>
        public static void ValidaCampo(Control ctrl, TextBox txtB)
        {
            foreach (Control control in ctrl.Controls)
            {
                if (control is TextBox)
                {
                    if (control.Name == txtB.Name)
                    {
                        control.Text = "0";
                        break;
                    }
                }
                else if (control is TabControl || 
                    control is TabPage || 
                    control is GroupBox)
                {
                    ValidaCampo(control, txtB);
                }
            }
        }

        #endregion

        #region Metodos de Configuración

        /// <summary>
        /// Elimina los ceros a la izquierda.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void clearZeros(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;

            txt.Text = double.Parse(txt.Text).ToString();
        }

        /// <summary>
        /// Obtiene el nombre del mes.
        /// </summary>
        /// <param name="month">El número de mes</param>
        /// <returns>El nombre de mes</returns>
        private string GetMonthName(int month)
        {
            string monthName = string.Empty;
            
            switch (month)
            {
                case 1:
                    monthName = "Enero";
                    break;
                case 2:
                    monthName = "Febrero";
                    break;
                case 3:
                    monthName = "Marzo";
                    break;
                case 4:
                    monthName = "Abril";
                    break;
                case 5:
                    monthName = "Mayo";
                    break;
                case 6:
                    monthName = "Junio";
                    break;
                case 7:
                    monthName = "Julio";
                    break;
                case 8:
                    monthName = "Agosto";
                    break;
                case 9:
                    monthName = "Septiembre";
                    break;
                case 10:
                    monthName = "Octubre";
                    break;
                case 11:
                    monthName = "Noviembre";
                    break;
                case 12:
                    monthName = "Diciembre";
                    break;
            }
            return monthName;
        }

        /// <summary>
        /// Obtiene el nombre del día
        /// </summary>
        /// <param name="dayOfWeek">El día de la semana</param>
        /// <returns>El nombre del día</returns>
        private string GetDayName(DayOfWeek dayOfWeek)
        {
            string dayName = string.Empty;

            switch (dayOfWeek.ToString())
            {
                case "Sunday":
                    dayName = "Domingo";
                    break;
                case "Monday":
                    dayName = "Lunes";
                    break;
                case "Tuesday":
                    dayName = "Martes";
                    break;
                case "Wednesday":
                    dayName = "Miércoles";
                    break;
                case "Thrusday":
                    dayName = "Jueves";
                    break;
                case "Friday":
                    dayName = "Viernes";
                    break;
                case "Saturday":
                    dayName = "Sábado";
                    break;
            }
            return dayName;
        }

        /// <summary>
        /// Carga los valores a las propiedades.
        /// </summary>
        private void SetProperties()
        {
            Properties.Settings.Default.Tasa02Q = double.Parse(txtTasa02Q.Text);
            Properties.Settings.Default.Tasa04Q = double.Parse(txtTasa04Q.Text);
            Properties.Settings.Default.Tasa06Q = double.Parse(txtTasa06Q.Text);
            Properties.Settings.Default.Tasa08Q = double.Parse(txtTasa08Q.Text);
            Properties.Settings.Default.Tasa10Q = double.Parse(txtTasa10Q.Text);
            Properties.Settings.Default.Tasa12Q = double.Parse(txtTasa12Q.Text);
            Properties.Settings.Default.MontoCredito01 = double.Parse(txtMontoCredito01.Text);
            Properties.Settings.Default.MontoCredito02 = double.Parse(txtMontoCredito02.Text);
            Properties.Settings.Default.MontoCredito03 = double.Parse(txtMontoCredito03.Text);
            Properties.Settings.Default.MontoCredito04 = double.Parse(txtMontoCredito04.Text);
            Properties.Settings.Default.MontoCredito05 = double.Parse(txtMontoCredito05.Text);
            Properties.Settings.Default.MontoCredito06 = double.Parse(txtMontoCredito06.Text);
            Properties.Settings.Default.IncrementoDP = double.Parse(txtPIncrementoDP.Text);
            Properties.Settings.Default.PermanenciaDP = double.Parse(txtPPermanenciaDP.Text);
            Properties.Settings.Default.IncrementoDC = double.Parse(txtPIncrementoDC.Text);
            Properties.Settings.Default.PermanenciaDC = double.Parse(txtPPermanenciaDC.Text);
            Properties.Settings.Default.IncrementoDO = double.Parse(txtPIncrementoDO.Text);
            Properties.Settings.Default.PermanenciaDO = double.Parse(txtPPermanenciaDO.Text);
            Properties.Settings.Default.IncrementoMMP = double.Parse(txtPIncrementoMMP.Text);
            Properties.Settings.Default.PermanenciaMMP = double.Parse(txtPPermanenciaMMP.Text);
            Properties.Settings.Default.IncrementoMMC = double.Parse(txtPIncrementoMMC.Text);
            Properties.Settings.Default.PermanenciaMMC = double.Parse(txtPPermanenciaMMC.Text);
            Properties.Settings.Default.IncrementoMMO = double.Parse(txtPIncrementoMMO.Text);
            Properties.Settings.Default.PermanenciaMMO = double.Parse(txtPPermanenciaMMO.Text);
            Properties.Settings.Default.IncrementoCZP = double.Parse(txtPIncrementoCZP.Text);
            Properties.Settings.Default.PermanenciaCZP = double.Parse(txtPPermanenciaCZP.Text);
            Properties.Settings.Default.IncrementoCZC = double.Parse(txtPIncrementoCZC.Text);
            Properties.Settings.Default.PermanenciaCZC = double.Parse(txtPPermanenciaCZC.Text);
            Properties.Settings.Default.IncrementoCZO = double.Parse(txtPIncrementoCZO.Text);
            Properties.Settings.Default.PermanenciaCZO = double.Parse(txtPPermanenciaCZO.Text);
            Properties.Settings.Default.CantDP = double.Parse(txtNCantDP.Text);
            Properties.Settings.Default.CantDC = double.Parse(txtNCantDC.Text);
            Properties.Settings.Default.CantDO = double.Parse(txtNCantDO.Text);
            Properties.Settings.Default.CantXDP = double.Parse(txtNCantXDP.Text);
            Properties.Settings.Default.CantXDC = double.Parse(txtNCantXDC.Text);
            Properties.Settings.Default.CantXDO = double.Parse(txtNCantXDO.Text);
            Properties.Settings.Default.CredDP = double.Parse(txtPCredDP.Text);
            Properties.Settings.Default.CredDC = double.Parse(txtPCredDC.Text);
            Properties.Settings.Default.CredDO = double.Parse(txtPCredDO.Text);
            Properties.Settings.Default.CantMMP = double.Parse(txtNCantMMP.Text);
            Properties.Settings.Default.CantMMC = double.Parse(txtNCantMMC.Text);
            Properties.Settings.Default.CantMMO = double.Parse(txtNCantMMO.Text);
            Properties.Settings.Default.ProsMMP = double.Parse(txtPProsMMP.Text);
            Properties.Settings.Default.ProsMMC = double.Parse(txtPProsMMC.Text);
            Properties.Settings.Default.ProsMMO = double.Parse(txtPProsMMO.Text);
            Properties.Settings.Default.CredMMP = double.Parse(txtPCredMMP.Text);
            Properties.Settings.Default.CredMMC = double.Parse(txtPCredMMC.Text);
            Properties.Settings.Default.CredMMO = double.Parse(txtPCredMMO.Text);
            Properties.Settings.Default.CantCZP = double.Parse(txtNCantCZP.Text);
            Properties.Settings.Default.CantCZC = double.Parse(txtNCantCZC.Text);
            Properties.Settings.Default.CantCZO = double.Parse(txtNCantCZO.Text);
            Properties.Settings.Default.ProsCZP = double.Parse(txtPProsCZP.Text);
            Properties.Settings.Default.ProsCZC = double.Parse(txtPProsCZC.Text);
            Properties.Settings.Default.ProsCZO = double.Parse(txtPProsCZO.Text);
            Properties.Settings.Default.CredCZP = double.Parse(txtPCredCZP.Text);
            Properties.Settings.Default.CredCZC = double.Parse(txtPCredCZC.Text);
            Properties.Settings.Default.CredCZO = double.Parse(txtPCredCZO.Text);
            Properties.Settings.Default.EscD = cmbEscD.SelectedIndex;
            Properties.Settings.Default.EscMM = cmbEscMM.SelectedIndex;
            Properties.Settings.Default.EscCZ = cmbEscCZ.SelectedIndex;
            Properties.Settings.Default.EscLC = cmbEscLC.SelectedIndex;
            Properties.Settings.Default.ProsLCP = double.Parse(txtPProsLCP.Text);
            Properties.Settings.Default.ProsLCC = double.Parse(txtPProsLCC.Text);
            Properties.Settings.Default.ProsLCO = double.Parse(txtPProsLCO.Text);
            Properties.Settings.Default.LideLCP = double.Parse(txtPInicLCP.Text);
            Properties.Settings.Default.LideLCC = double.Parse(txtPInicLCC.Text);
            Properties.Settings.Default.LideLCO = double.Parse(txtPInicLCO.Text);
            Properties.Settings.Default.PermanenciaLCP = double.Parse(txtPPermanenciaLCP.Text);
            Properties.Settings.Default.PermanenciaLCC = double.Parse(txtPPermanenciaLCC.Text);
            Properties.Settings.Default.PermanenciaLCO = double.Parse(txtPPermanenciaLCO.Text);
            Properties.Settings.Default.ProbTamCel2MCP = double.Parse(txtPProbTamCel2MCP.Text);
            Properties.Settings.Default.ProbTamCel3MCP = double.Parse(txtPProbTamCel3MCP.Text);
            Properties.Settings.Default.ProbTamCel4MCP = double.Parse(txtPProbTamCel4MCP.Text);
            Properties.Settings.Default.ProbTamCel5MCP = double.Parse(txtPProbTamCel5MCP.Text);
            Properties.Settings.Default.ProbTamCel6MCP = double.Parse(txtPProbTamCel6MCP.Text);
            Properties.Settings.Default.ProbTamCel7MCP = double.Parse(txtPProbTamCel7MCP.Text);
            Properties.Settings.Default.ProbTamCel8MCP = double.Parse(txtPProbTamCel8MCP.Text);
            Properties.Settings.Default.ProbTamCel9MCP = double.Parse(txtPProbTamCel9MCP.Text);
            Properties.Settings.Default.ProbTamCel2MCC = double.Parse(txtPProbTamCel2MCC.Text);
            Properties.Settings.Default.ProbTamCel3MCC = double.Parse(txtPProbTamCel3MCC.Text);
            Properties.Settings.Default.ProbTamCel4MCC = double.Parse(txtPProbTamCel4MCC.Text);
            Properties.Settings.Default.ProbTamCel5MCC = double.Parse(txtPProbTamCel5MCC.Text);
            Properties.Settings.Default.ProbTamCel6MCC = double.Parse(txtPProbTamCel6MCC.Text);
            Properties.Settings.Default.ProbTamCel7MCC = double.Parse(txtPProbTamCel7MCC.Text);
            Properties.Settings.Default.ProbTamCel8MCC = double.Parse(txtPProbTamCel8MCC.Text);
            Properties.Settings.Default.ProbTamCel9MCC = double.Parse(txtPProbTamCel9MCC.Text);
            Properties.Settings.Default.ProbTamCel2MCO = double.Parse(txtPProbTamCel2MCO.Text);
            Properties.Settings.Default.ProbTamCel3MCO = double.Parse(txtPProbTamCel3MCO.Text);
            Properties.Settings.Default.ProbTamCel4MCO = double.Parse(txtPProbTamCel4MCO.Text);
            Properties.Settings.Default.ProbTamCel5MCO = double.Parse(txtPProbTamCel5MCO.Text);
            Properties.Settings.Default.ProbTamCel6MCO = double.Parse(txtPProbTamCel6MCO.Text);
            Properties.Settings.Default.ProbTamCel7MCO = double.Parse(txtPProbTamCel7MCO.Text);
            Properties.Settings.Default.ProbTamCel8MCO = double.Parse(txtPProbTamCel8MCO.Text);
            Properties.Settings.Default.ProbTamCel9MCO = double.Parse(txtPProbTamCel9MCO.Text);
            Properties.Settings.Default.SeguroI = double.Parse(txtSeguroI.Text);
            Properties.Settings.Default.ComAperturaI = double.Parse(txtComAperturaI.Text);
            Properties.Settings.Default.CobroXPlasticoI = double.Parse(txtCobroXPlasticoI.Text);
            Properties.Settings.Default.IngresoProsaI = double.Parse(txtIngresoProsaI.Text);
            Properties.Settings.Default.PerdidaE = double.Parse(txtPerdidaE.Text);
            Properties.Settings.Default.ComisionDistE = double.Parse(txtComisionDistE.Text);
            Properties.Settings.Default.GastosFijosPROSAE = double.Parse(txtGastosFijosPROSAE.Text);
            Properties.Settings.Default.GastosVarPROSAE = double.Parse(txtGastosVarPROSAE.Text);
            Properties.Settings.Default.GastosFijosZafyE = double.Parse(txtGastosFijosZafyE.Text);
            Properties.Settings.Default.GastosVarZafyE = double.Parse(txtGastosVarZafyE.Text);
            Properties.Settings.Default.GastosXPublicidadE = double.Parse(txtGastosXPublicidadE.Text);
            Properties.Settings.Default.LimiteXISRE = double.Parse(txtLimiteXISRE.Text);
            Properties.Settings.Default.BonosPremiosE = double.Parse(txtBonosPremiosE.Text);
            Properties.Settings.Default.RetirosE = double.Parse(txtRetirosE.Text);

            Properties.Settings.Default.PagoPronto = double.Parse(txtPagoProntoVal.Text);
            Properties.Settings.Default.PagoTiempo = double.Parse(txtPagoTiempoVal.Text);
            Properties.Settings.Default.PagoTardio1 = double.Parse(txtPagoTardio1Val.Text);
            Properties.Settings.Default.PagoTardio2 = double.Parse(txtPagoTardio2Val.Text);
            Properties.Settings.Default.PagoTardio3 = double.Parse(txtPagoTardio3Val.Text);
            Properties.Settings.Default.PagoProntoPerc = double.Parse(txtDistPagoProntoPerc.Text);
            Properties.Settings.Default.PagoTiempoPerc = double.Parse(txtDistPagoTiempoPerc.Text);
            Properties.Settings.Default.PagoTardio1Perc = double.Parse(txtDistPagoTardio1Perc.Text);
            Properties.Settings.Default.PagoTardio2Perc = double.Parse(txtDistPagoTardio2Perc.Text);
            Properties.Settings.Default.PagoTardio3Perc = double.Parse(txtDistPagoTardio3Perc.Text);
            Properties.Settings.Default.PagoTiempoCN1 = double.Parse(txtPagoTiempoCN1Val.Text);
            Properties.Settings.Default.PagoTiempoCN2 = double.Parse(txtPagoTiempoCN2Val.Text);
            Properties.Settings.Default.PagoTiempoCN3 = double.Parse(txtPagoTiempoCN3Val.Text);
            Properties.Settings.Default.PagoTardio1CN1 = double.Parse(txtPagoTardio1CN1Val.Text);
            Properties.Settings.Default.PagoTardio1CN2 = double.Parse(txtPagoTardio1CN2Val.Text);
            Properties.Settings.Default.PagoTardio1CN3 = double.Parse(txtPagoTardio1CN3Val.Text);
            Properties.Settings.Default.PagoTardio2CN1 = double.Parse(txtPagoTardio2CN1Val.Text);
            Properties.Settings.Default.PagoTardio2CN2 = double.Parse(txtPagoTardio2CN2Val.Text);
            Properties.Settings.Default.PagoTardio2CN3 = double.Parse(txtPagoTardio2CN3Val.Text);
            Properties.Settings.Default.PagoTardio3CN1 = double.Parse(txtPagoTardio3CN1Val.Text);
            Properties.Settings.Default.PagoTardio3CN2 = double.Parse(txtPagoTardio3CN2Val.Text);
            Properties.Settings.Default.PagoTardio3CN3 = double.Parse(txtPagoTardio3CN3Val.Text);
            Properties.Settings.Default.PorColocacion = double.Parse(txtPorColocacionVal.Text);
            Properties.Settings.Default.PagoTiempoCN1Perc = double.Parse(txtN1PagoTiempoCN1Perc.Text);
            Properties.Settings.Default.PagoTiempoCN2Perc = double.Parse(txtN2PagoTiempoCN2Perc.Text);
            Properties.Settings.Default.PagoTiempoCN3Perc = double.Parse(txtN3PagoTiempoCN3Perc.Text);
            Properties.Settings.Default.PagoTardio1CN1Perc = double.Parse(txtN1PagoTardio1CN1Perc.Text);
            Properties.Settings.Default.PagoTardio1CN2Perc = double.Parse(txtN2PagoTardio1CN2Perc.Text);
            Properties.Settings.Default.PagoTardio1CN3Perc = double.Parse(txtN3PagoTardio1CN3Perc.Text);
            Properties.Settings.Default.PagoTardio2CN1Perc = double.Parse(txtN1PagoTardio2CN1Perc.Text);
            Properties.Settings.Default.PagoTardio2CN2Perc = double.Parse(txtN2PagoTardio2CN2Perc.Text);
            Properties.Settings.Default.PagoTardio2CN3Perc = double.Parse(txtN3PagoTardio2CN3Perc.Text);
            Properties.Settings.Default.PagoTardio3CN1Perc = double.Parse(txtN1PagoTardio3CN1Perc.Text);
            Properties.Settings.Default.PagoTardio3CN2Perc = double.Parse(txtN2PagoTardio3CN2Perc.Text);
            Properties.Settings.Default.PagoTardio3CN3Perc = double.Parse(txtN3PagoTardio3CN3Perc.Text);
            Properties.Settings.Default.PorCumplimiento = double.Parse(txtPorCumplimientoVal.Text);
            Properties.Settings.Default.PorCumplimientoPerc = double.Parse(txtPorCumplimientoPerc.Text);
            Properties.Settings.Default.PorCrecimiento = double.Parse(txtPorCrecimientoVal.Text);
            Properties.Settings.Default.PorCompletarCelula = double.Parse(txtPorCompletarCelulaVal.Text);
            Properties.Settings.Default.PorColocacionValorMenor = double.Parse(txtPorVolumenColocadoValorMenor.Text);
            Properties.Settings.Default.PorColocacionMontoMenor = double.Parse(txtPorVolumenColocadoMontoMenor.Text);
            Properties.Settings.Default.PorColocacionValorEntre = double.Parse(txtPorVolumenMiembrosValorEntre.Text);
            Properties.Settings.Default.PorColocacionMontoDe = double.Parse(txtPorVolumenColocadoMontoDe.Text);
            Properties.Settings.Default.PorColocacionMontoHasta = double.Parse(txtPorVolumenColocadoMontoHasta.Text);
            Properties.Settings.Default.PorColocacionValorMayor = double.Parse(txtPorVolumenMiembrosValorMayor.Text);
            Properties.Settings.Default.PorColocacionMontoMayor = double.Parse(txtPorVolumenColocadoMontoMayor.Text);
            Properties.Settings.Default.PorVolumenValorMenor = double.Parse(txtPorVolumenMiembrosValorMenor.Text);
            Properties.Settings.Default.PorVolumenVolumenMenor = double.Parse(txtPorVolumenMiembrosVolumenMenor.Text);
            Properties.Settings.Default.PorVolumenValorEntre = double.Parse(txtPorVolumenMiembrosValorEntre.Text);
            Properties.Settings.Default.PorVolumenVolumenDe = double.Parse(txtPorVolumenMiembrosVolumenDe.Text);
            Properties.Settings.Default.PorVolumenVolumenHasta = double.Parse(txtPorVolumenMiembrosVolumenHasta.Text);
            Properties.Settings.Default.PorVolumenValorMayor = double.Parse(txtPorVolumenMiembrosValorMayor.Text);
            Properties.Settings.Default.PorVolumenVolumenMayor = double.Parse(txtPorVolumenMiembrosVolumenMayor.Text);
        }

        /// <summary>
        /// Obtiene el valor de las propiedades.
        /// </summary>
        private void GetProperties()
        {
            //Asigna los valores a las campos de texto en pantalla
            txtTasa02Q.Text = Properties.Settings.Default.Tasa02Q.ToString();
            txtTasa04Q.Text = Properties.Settings.Default.Tasa04Q.ToString();
            txtTasa06Q.Text = Properties.Settings.Default.Tasa06Q.ToString();
            txtTasa08Q.Text = Properties.Settings.Default.Tasa08Q.ToString();
            txtTasa10Q.Text = Properties.Settings.Default.Tasa10Q.ToString();
            txtTasa12Q.Text = Properties.Settings.Default.Tasa12Q.ToString();
            txtMontoCredito01.Text = Properties.Settings.Default.MontoCredito01.ToString();
            txtMontoCredito02.Text = Properties.Settings.Default.MontoCredito02.ToString();
            txtMontoCredito03.Text = Properties.Settings.Default.MontoCredito03.ToString();
            txtMontoCredito04.Text = Properties.Settings.Default.MontoCredito04.ToString();
            txtMontoCredito05.Text = Properties.Settings.Default.MontoCredito05.ToString();
            txtMontoCredito06.Text = Properties.Settings.Default.MontoCredito06.ToString();
            txtPIncrementoDP.Text = Properties.Settings.Default.IncrementoDP.ToString();
            txtPPermanenciaDP.Text = Properties.Settings.Default.PermanenciaDP.ToString();
            txtPIncrementoDC.Text = Properties.Settings.Default.IncrementoDC.ToString();
            txtPPermanenciaDC.Text = Properties.Settings.Default.PermanenciaDC.ToString();
            txtPIncrementoDO.Text = Properties.Settings.Default.IncrementoDO.ToString();
            txtPPermanenciaDO.Text = Properties.Settings.Default.PermanenciaDO.ToString();
            txtPIncrementoMMP.Text = Properties.Settings.Default.IncrementoMMP.ToString();
            txtPPermanenciaMMP.Text = Properties.Settings.Default.PermanenciaMMP.ToString();
            txtPIncrementoMMC.Text = Properties.Settings.Default.IncrementoMMC.ToString();
            txtPPermanenciaMMC.Text = Properties.Settings.Default.PermanenciaMMC.ToString();
            txtPIncrementoMMO.Text = Properties.Settings.Default.IncrementoMMO.ToString();
            txtPPermanenciaMMO.Text = Properties.Settings.Default.PermanenciaMMO.ToString();
            txtPIncrementoCZP.Text = Properties.Settings.Default.IncrementoCZP.ToString();
            txtPPermanenciaCZP.Text = Properties.Settings.Default.PermanenciaCZP.ToString();
            txtPIncrementoCZC.Text = Properties.Settings.Default.IncrementoCZC.ToString();
            txtPPermanenciaCZC.Text = Properties.Settings.Default.PermanenciaCZC.ToString();
            txtPIncrementoCZO.Text = Properties.Settings.Default.IncrementoCZO.ToString();
            txtPPermanenciaCZO.Text = Properties.Settings.Default.PermanenciaCZO.ToString();
            txtNCantDP.Text = Properties.Settings.Default.CantDP.ToString();
            txtNCantDC.Text = Properties.Settings.Default.CantDC.ToString();
            txtNCantDO.Text = Properties.Settings.Default.CantDO.ToString();
            txtNCantXDP.Text = Properties.Settings.Default.CantXDP.ToString();
            txtNCantXDC.Text = Properties.Settings.Default.CantXDC.ToString();
            txtNCantXDO.Text = Properties.Settings.Default.CantXDO.ToString();
            txtPCredDP.Text = Properties.Settings.Default.CredDP.ToString();
            txtPCredDC.Text = Properties.Settings.Default.CredDC.ToString();
            txtPCredDO.Text = Properties.Settings.Default.CredDO.ToString();
            txtNCantMMP.Text = Properties.Settings.Default.CantMMP.ToString();
            txtNCantMMC.Text = Properties.Settings.Default.CantMMC.ToString();
            txtNCantMMO.Text = Properties.Settings.Default.CantMMO.ToString();
            txtPProsMMP.Text = Properties.Settings.Default.ProsMMP.ToString();
            txtPProsMMC.Text = Properties.Settings.Default.ProsMMC.ToString();
            txtPProsMMO.Text = Properties.Settings.Default.ProsMMO.ToString();
            txtPCredMMP.Text = Properties.Settings.Default.CredMMP.ToString();
            txtPCredMMC.Text = Properties.Settings.Default.CredMMC.ToString();
            txtPCredMMO.Text = Properties.Settings.Default.CredMMO.ToString();
            txtNCantCZP.Text = Properties.Settings.Default.CantCZP.ToString();
            txtNCantCZC.Text = Properties.Settings.Default.CantCZC.ToString();
            txtNCantCZO.Text = Properties.Settings.Default.CantCZO.ToString();
            txtPProsCZP.Text = Properties.Settings.Default.ProsCZP.ToString();
            txtPProsCZC.Text = Properties.Settings.Default.ProsCZC.ToString();
            txtPProsCZO.Text = Properties.Settings.Default.ProsCZO.ToString();
            txtPCredCZP.Text = Properties.Settings.Default.CredCZP.ToString();
            txtPCredCZC.Text = Properties.Settings.Default.CredCZC.ToString();
            txtPCredCZO.Text = Properties.Settings.Default.CredCZO.ToString();
            cmbEscD.SelectedIndex = Properties.Settings.Default.EscD;
            cmbEscMM.SelectedIndex = Properties.Settings.Default.EscMM;
            cmbEscCZ.SelectedIndex = Properties.Settings.Default.EscCZ;
            cmbEscLC.SelectedIndex = Properties.Settings.Default.EscLC;
            txtPProsLCP.Text = Properties.Settings.Default.ProsLCP.ToString();
            txtPProsLCC.Text = Properties.Settings.Default.ProsLCC.ToString();
            txtPProsLCO.Text = Properties.Settings.Default.ProsLCO.ToString();
            txtPInicLCP.Text = Properties.Settings.Default.LideLCP.ToString();
            txtPInicLCC.Text = Properties.Settings.Default.LideLCC.ToString();
            txtPInicLCO.Text = Properties.Settings.Default.LideLCO.ToString();
            txtPPermanenciaLCP.Text = Properties.Settings.Default.PermanenciaLCP.ToString();
            txtPPermanenciaLCC.Text = Properties.Settings.Default.PermanenciaLCC.ToString();
            txtPPermanenciaLCO.Text = Properties.Settings.Default.PermanenciaLCO.ToString();
            txtPProbTamCel2MCP.Text = Properties.Settings.Default.ProbTamCel2MCP.ToString();
            txtPProbTamCel3MCP.Text = Properties.Settings.Default.ProbTamCel3MCP.ToString();
            txtPProbTamCel4MCP.Text = Properties.Settings.Default.ProbTamCel4MCP.ToString();
            txtPProbTamCel5MCP.Text = Properties.Settings.Default.ProbTamCel5MCP.ToString();
            txtPProbTamCel6MCP.Text = Properties.Settings.Default.ProbTamCel6MCP.ToString();
            txtPProbTamCel7MCP.Text = Properties.Settings.Default.ProbTamCel7MCP.ToString();
            txtPProbTamCel8MCP.Text = Properties.Settings.Default.ProbTamCel8MCP.ToString();
            txtPProbTamCel9MCP.Text = Properties.Settings.Default.ProbTamCel9MCP.ToString();
            txtPProbTamCel2MCC.Text = Properties.Settings.Default.ProbTamCel2MCC.ToString();
            txtPProbTamCel3MCC.Text = Properties.Settings.Default.ProbTamCel3MCC.ToString();
            txtPProbTamCel4MCC.Text = Properties.Settings.Default.ProbTamCel4MCC.ToString();
            txtPProbTamCel5MCC.Text = Properties.Settings.Default.ProbTamCel5MCC.ToString();
            txtPProbTamCel6MCC.Text = Properties.Settings.Default.ProbTamCel6MCC.ToString();
            txtPProbTamCel7MCC.Text = Properties.Settings.Default.ProbTamCel7MCC.ToString();
            txtPProbTamCel8MCC.Text = Properties.Settings.Default.ProbTamCel8MCC.ToString();
            txtPProbTamCel9MCC.Text = Properties.Settings.Default.ProbTamCel9MCC.ToString();
            txtPProbTamCel2MCO.Text = Properties.Settings.Default.ProbTamCel2MCO.ToString();
            txtPProbTamCel3MCO.Text = Properties.Settings.Default.ProbTamCel3MCO.ToString();
            txtPProbTamCel4MCO.Text = Properties.Settings.Default.ProbTamCel4MCO.ToString();
            txtPProbTamCel5MCO.Text = Properties.Settings.Default.ProbTamCel5MCO.ToString();
            txtPProbTamCel6MCO.Text = Properties.Settings.Default.ProbTamCel6MCO.ToString();
            txtPProbTamCel7MCO.Text = Properties.Settings.Default.ProbTamCel7MCO.ToString();
            txtPProbTamCel8MCO.Text = Properties.Settings.Default.ProbTamCel8MCO.ToString();
            txtPProbTamCel9MCO.Text = Properties.Settings.Default.ProbTamCel9MCO.ToString();
            txtSeguroI.Text = Properties.Settings.Default.SeguroI.ToString();
            txtComAperturaI.Text = Properties.Settings.Default.ComAperturaI.ToString();
            txtCobroXPlasticoI.Text = Properties.Settings.Default.CobroXPlasticoI.ToString();
            txtIngresoProsaI.Text = Properties.Settings.Default.IngresoProsaI.ToString();
            txtPerdidaE.Text = Properties.Settings.Default.PerdidaE.ToString();
            txtComisionDistE.Text = Properties.Settings.Default.ComisionDistE.ToString();
            txtGastosFijosPROSAE.Text = Properties.Settings.Default.GastosFijosPROSAE.ToString();
            txtGastosVarPROSAE.Text = Properties.Settings.Default.GastosVarPROSAE.ToString();
            txtGastosFijosZafyE.Text = Properties.Settings.Default.GastosFijosZafyE.ToString();
            txtGastosVarZafyE.Text = Properties.Settings.Default.GastosVarZafyE.ToString();
            txtGastosXPublicidadE.Text = Properties.Settings.Default.GastosXPublicidadE.ToString();
            txtLimiteXISRE.Text = Properties.Settings.Default.LimiteXISRE.ToString();
            txtBonosPremiosE.Text = Properties.Settings.Default.BonosPremiosE.ToString();
            txtRetirosE.Text = Properties.Settings.Default.RetirosE.ToString();

            txtPagoProntoVal.Text = Properties.Settings.Default.PagoPronto.ToString();
            txtPagoTiempoVal.Text = Properties.Settings.Default.PagoTiempo.ToString();
            txtPagoTardio1Val.Text = Properties.Settings.Default.PagoTardio1.ToString();
            txtPagoTardio2Val.Text = Properties.Settings.Default.PagoTardio2.ToString();
            txtPagoTardio3Val.Text = Properties.Settings.Default.PagoTardio3.ToString();
            txtDistPagoProntoPerc.Text = Properties.Settings.Default.PagoProntoPerc.ToString();
            txtDistPagoTiempoPerc.Text = Properties.Settings.Default.PagoTiempoPerc.ToString();
            txtDistPagoTardio1Perc.Text = Properties.Settings.Default.PagoTardio1Perc.ToString();
            txtDistPagoTardio2Perc.Text = Properties.Settings.Default.PagoTardio2Perc.ToString();
            txtDistPagoTardio3Perc.Text = Properties.Settings.Default.PagoTardio3Perc.ToString();
            txtPagoTiempoCN1Val.Text = Properties.Settings.Default.PagoTiempoCN1.ToString();
            txtPagoTiempoCN2Val.Text = Properties.Settings.Default.PagoTiempoCN2.ToString();
            txtPagoTiempoCN3Val.Text = Properties.Settings.Default.PagoTiempoCN3.ToString();
            txtPagoTardio1CN1Val.Text = Properties.Settings.Default.PagoTardio1CN1.ToString();
            txtPagoTardio1CN2Val.Text = Properties.Settings.Default.PagoTardio1CN2.ToString();
            txtPagoTardio1CN3Val.Text = Properties.Settings.Default.PagoTardio1CN3.ToString();
            txtPagoTardio2CN1Val.Text = Properties.Settings.Default.PagoTardio2CN1.ToString();
            txtPagoTardio2CN2Val.Text = Properties.Settings.Default.PagoTardio2CN2.ToString();
            txtPagoTardio2CN3Val.Text = Properties.Settings.Default.PagoTardio2CN3.ToString();
            txtPagoTardio3CN1Val.Text = Properties.Settings.Default.PagoTardio3CN1.ToString();
            txtPagoTardio3CN2Val.Text = Properties.Settings.Default.PagoTardio3CN2.ToString();
            txtPagoTardio3CN3Val.Text = Properties.Settings.Default.PagoTardio3CN3.ToString();
            txtPorColocacionVal.Text = Properties.Settings.Default.PorColocacion.ToString();
            txtN1PagoTiempoCN1Perc.Text = Properties.Settings.Default.PagoTiempoCN1Perc.ToString();
            txtN2PagoTiempoCN2Perc.Text = Properties.Settings.Default.PagoTiempoCN2Perc.ToString();
            txtN3PagoTiempoCN3Perc.Text = Properties.Settings.Default.PagoTiempoCN3Perc.ToString();
            txtN1PagoTardio1CN1Perc.Text = Properties.Settings.Default.PagoTardio1CN1Perc.ToString();
            txtN2PagoTardio1CN2Perc.Text = Properties.Settings.Default.PagoTardio1CN2Perc.ToString();
            txtN3PagoTardio1CN3Perc.Text = Properties.Settings.Default.PagoTardio1CN3Perc.ToString();
            txtN1PagoTardio2CN1Perc.Text = Properties.Settings.Default.PagoTardio2CN1Perc.ToString();
            txtN2PagoTardio2CN2Perc.Text = Properties.Settings.Default.PagoTardio2CN2Perc.ToString();
            txtN3PagoTardio2CN3Perc.Text = Properties.Settings.Default.PagoTardio2CN3Perc.ToString();
            txtN1PagoTardio3CN1Perc.Text = Properties.Settings.Default.PagoTardio3CN1Perc.ToString();
            txtN2PagoTardio3CN2Perc.Text = Properties.Settings.Default.PagoTardio3CN2Perc.ToString();
            txtN3PagoTardio3CN3Perc.Text = Properties.Settings.Default.PagoTardio3CN3Perc.ToString();
            txtPorCumplimientoVal.Text = Properties.Settings.Default.PorCumplimiento.ToString();
            txtPorCumplimientoPerc.Text = Properties.Settings.Default.PorCumplimientoPerc.ToString();
            txtPorCrecimientoVal.Text = Properties.Settings.Default.PorCrecimiento.ToString();
            txtPorCompletarCelulaVal.Text = Properties.Settings.Default.PorCompletarCelula.ToString();
            txtPorVolumenColocadoValorMenor.Text = Properties.Settings.Default.PorColocacionValorMenor.ToString();
            txtPorVolumenColocadoMontoMenor.Text = Properties.Settings.Default.PorColocacionMontoMenor.ToString();
            txtPorVolumenMiembrosValorEntre.Text = Properties.Settings.Default.PorColocacionValorEntre.ToString();
            txtPorVolumenColocadoMontoDe.Text = Properties.Settings.Default.PorColocacionMontoDe.ToString();
            txtPorVolumenColocadoMontoHasta.Text = Properties.Settings.Default.PorColocacionMontoHasta.ToString();
            txtPorVolumenMiembrosValorMayor.Text = Properties.Settings.Default.PorColocacionValorMayor.ToString();
            txtPorVolumenColocadoMontoMayor.Text = Properties.Settings.Default.PorColocacionMontoMayor.ToString();
            txtPorVolumenMiembrosValorMenor.Text = Properties.Settings.Default.PorVolumenValorMenor.ToString();
            txtPorVolumenMiembrosVolumenMenor.Text = Properties.Settings.Default.PorVolumenVolumenMenor.ToString();
            txtPorVolumenMiembrosValorEntre.Text = Properties.Settings.Default.PorVolumenValorEntre.ToString();
            txtPorVolumenMiembrosVolumenDe.Text = Properties.Settings.Default.PorVolumenVolumenDe.ToString();
            txtPorVolumenMiembrosVolumenHasta.Text = Properties.Settings.Default.PorVolumenVolumenHasta.ToString();
            txtPorVolumenMiembrosValorMayor.Text = Properties.Settings.Default.PorVolumenValorMayor.ToString();
            txtPorVolumenMiembrosVolumenMayor.Text = Properties.Settings.Default.PorVolumenVolumenMayor.ToString();
        }
        
        /// <summary>
        /// Activa o desactiva la función para generar a las líderes iniciadoras.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void btnIniciadoras_Click(object sender, EventArgs e)
        {
            if (!isIniciadorasOn)
            {
                isIniciadorasOn = true;
                btnIniciadoras.BackColor = Color.Green;
                btnIniciadoras.Text = "I";
            }
            else
            {
                isIniciadorasOn = false;
                btnIniciadoras.BackColor = Color.Red;
                btnIniciadoras.Text = "O";
            }
        }

        /// <summary>
        /// Carga los valores por defecto de las propiedades de configuración.
        /// </summary>
        /// <param name="sender">Los objetos que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void btnLoadDefaults_Click(object sender, EventArgs e)
        {
            if (periodoActual > 1)
            {
                MessageBox.Show(
                    "No se puede restaurar la configuración hasta terminar el "
                    + "flujo actual", "Configuración no restaurada");
            }
            else
            {
                if (DialogResult.Yes == MessageBox.Show(
                    "¿Desea cargar los valores iniciales?", 
                    "Restaurar configuración", MessageBoxButtons.YesNo))
                {
                    Properties.Settings.Default.Reset();
                    Properties.Settings.Default.SessionId = sesionId;
                    Properties.Settings.Default.Save();
                    this.GetProperties();
                }
            }
        }

        /// <summary>
        /// Selecciona las celdas.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void dgvFlujoT_SelectionChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dgvFlujoT.Rows.Count; i++)
            {

                dgvFlujo.Rows[i].Selected = dgvFlujoT.Rows[i].Selected;
                dgvFlujoP.Rows[i].Selected = dgvFlujoT.Rows[i].Selected;
            }
        }

        /// <summary>
        /// Selecciona las celdas.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void dgvFlujoT_CellClick(object sender, 
            DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < dgvFlujoT.Rows.Count; i++)
            {
                if (e.ColumnIndex > 0)
                {
                    dgvFlujo.Rows[i].Selected = 
                        dgvFlujoT.Rows[i].Cells[e.ColumnIndex].Selected;
                    dgvFlujoP.Rows[i].Selected = 
                        dgvFlujoT.Rows[i].Cells[e.ColumnIndex].Selected;
                }
            }
        }

        /// <summary>
        /// Selecciona el primer dato
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void tabConfiguracion_Click(object sender, EventArgs e)
        {
            txtSeguroI.Select();
        }

        /// <summary>
        /// Activa o desactiva la función para agregar a las distribuidoras
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void btnInscribirDistribuidoras_Click(object sender, EventArgs e)
        {
            if (!isDistribuidorasOn)
            {
                isDistribuidorasOn = true;
                btnInscribirDistribuidoras.BackColor = Color.Green;
                btnInscribirDistribuidoras.Text = "I";
            }
            else
            {
                isDistribuidorasOn = false;
                btnInscribirDistribuidoras.BackColor = Color.Red;
                btnInscribirDistribuidoras.Text = "O";
            }
        }
        
        /// <summary>
        /// Muestra los detalles de la cartera actual
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void btnDetalle_Click(object sender, EventArgs e)
        {
            FormDetalle frmDetalle = new FormDetalle();
            frmDetalle.StartPosition = FormStartPosition.CenterScreen;
            frmDetalle.ShowDialog();
        }
        #endregion

        private void btnReporte_Click(object sender, EventArgs e)
        {
            Impresor impresor = new Impresor();

            impresor.CreatePDFNoTemplate();
        }
    }
}
