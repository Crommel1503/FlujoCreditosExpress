using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace FlujoCreditosExpress
{
    public partial class FormColocacion : Form
    {
        public FormColocacion()
        {
            InitializeComponent();
        }

        #region Global Variables

        int configId = 1;
        int clienteID = Properties.Settings.Default.LastIdCliente;
        ArrayList clienteIDList = new ArrayList();
        double apCapital = 0;
        double apCapitalAcum = 0;
        double aportacionAcum = Properties.Settings.Default.AportacionAcumulada;
        double cobroXPlastico = Properties.Settings.Default.CobroXPlasticoI;
        double creditosCN = 0;
        double creditosCA = 0;
        double creditos2Q = 0;
        double creditos2QD = 0;
        double creditos2QCD = 0;
        double creditos2QMM = 0;
        double creditos2QCZ = 0;
        double creditos2QMC = 0;
        double creditos2QCDT = 0;
        double creditos2QMMT = 0;
        double creditos2QCZT = 0;
        double creditos2QMCT = 0;
        double creditos4Q = 0;
        double creditos4QD = 0;
        double creditos4QCD = 0;
        double creditos4QMM = 0;
        double creditos4QCZ = 0;
        double creditos4QMC = 0;
        double creditos4QCDT = 0;
        double creditos4QMMT = 0;
        double creditos4QCZT = 0;
        double creditos4QMCT = 0;
        double creditos6Q = 0;
        double creditos6QD = 0;
        double creditos6QCDT = 0;
        double creditos6QMMT = 0;
        double creditos6QCZT = 0;
        double creditos6QMCT = 0;
        double creditos8Q = 0;
        double creditos8QD = 0;
        double creditos8QCDT = 0;
        double creditos8QMMT = 0;
        double creditos8QCZT = 0;
        double creditos8QMCT = 0;
        double creditos10Q = 0;
        double creditos10QD = 0;
        double creditos10QCDT = 0;
        double creditos10QMMT = 0;
        double creditos10QCZT = 0;
        double creditos10QMCT = 0;
        double creditos12Q = 0;
        double creditos12QD = 0;
        double creditos12QCDT = 0;
        double creditos12QMMT = 0;
        double creditos12QCZT = 0;
        double creditos12QMCT = 0;
        double creditosI = 0;
        double creditosRestantesN = 0;
        double creditosRestantesA = 0;
        double capitalRestante = 0;
        double apCapitalTemp = 0;
        double colocacion = 0;
        double periodoActual = Properties.Settings.Default.PeriodoActual;
        double capital = Properties.Settings.Default.Capital;
        double clientesPeriodoAnt = Properties.Settings.Default.ClientesNuevos;
        double incremento = 0;
        double permanencia = 0;
        double permanenciaD = Properties.Settings.Default.PermanenciaDVal;
        double permanenciaMM = Properties.Settings.Default.PermanenciaMMVal;
        double permanenciaCZ = Properties.Settings.Default.PermanenciaCZVal;
        double permanenciaMC = Properties.Settings.Default.PermanenciaMCVal;
        double clientesNuevos = 0;
        double clientes2Credito = 0;
        double dist2Credito = 0;
        double clientesD2Credito = 0;
        double clientesMM2Credito = 0;
        double clientesZ2Credito = 0;
        double clientesMC2Credito = 0;
        double clientesT2Credito = 0;
        double clientesDP2C = 0;
        double clientesMMP2C = 0;
        double clientesZafyP2C = 0;
        double clientesMCP2C = 0;
        double clientesDP2CPerm = 0;
        double clientesMMP2CPerm = 0;
        double clientesZP2CPerm = 0;
        double clientesMCP2CPerm = 0;
        double clientesPermanencia = 0;
        double c1 = Properties.Settings.Default.MontoCredito01;
        double c2 = Properties.Settings.Default.MontoCredito02;
        double c3 = Properties.Settings.Default.MontoCredito03;
        double c4 = Properties.Settings.Default.MontoCredito04;
        double c5 = Properties.Settings.Default.MontoCredito05;
        double c6 = Properties.Settings.Default.MontoCredito06;
        double i02 = Properties.Settings.Default.Tasa02Q / 100;
        double i04 = Properties.Settings.Default.Tasa04Q / 100;
        double i06 = Properties.Settings.Default.Tasa06Q / 100;
        double i08 = Properties.Settings.Default.Tasa08Q / 100;
        double i10 = Properties.Settings.Default.Tasa10Q / 100;
        double i12 = Properties.Settings.Default.Tasa12Q / 100;
        double perdidaE = Properties.Settings.Default.PerdidaE;
        double comisionDistE = Properties.Settings.Default.ComisionDistE;
        double gastosFijosPROSAE = Properties.Settings.Default.GastosFijosPROSAE;
        double gastosVarPROSAE = Properties.Settings.Default.GastosVarPROSAE;
        double gastosFijosZafyE = Properties.Settings.Default.GastosFijosZafyE;
        double gastosVarZafyE = Properties.Settings.Default.GastosVarZafyE;
        double gastosXPublicidad = Properties.Settings.Default.GastosXPublicidadE;
        double bonosPremiosE = Properties.Settings.Default.BonosPremiosE;
        double retiroE = Properties.Settings.Default.RetirosE;
        double carteraTotal = Properties.Settings.Default.CarteraTotal;
        bool isApCapCte = false;
        FlujoDBDataSet.T_ConfiguracionesRow tConfiguracionesRow;
        double distribuidoras;
        double ctesDistAnt;
        double ctesDist;
        double ctesMMAnt;
        double ctesMM;
        double ctesCZAnt;
        double ctesCZ;
        double ctesMCAnt;
        double ctesMC;

        #endregion

        #region Métodos privados
        /// <summary>
        /// Carga el formulario.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void FormColocacion_Load(object sender, EventArgs e)
        {
            t_ConfiguracionesTableAdapter1.Fill(flujoDBDataSet1.T_Configuraciones);
            t_AmortizacionesTableAdapter1.Fill(flujoDBDataSet1.T_Amortizaciones);
            this.LoadData(sender, e);
            this.SaveInitialConfig(sender, e);

            if (Properties.Settings.Default.IsAutomatic && periodoActual != 5)
            {
                btnSaveColocacion_Click(sender, e);
            }
            else if (Properties.Settings.Default.IsAutomatic)
            {
                MessageBox.Show("Se requiere más información para continuar. " +
                    "\nEs necesario completar la colocación de segundos créditos", "Alerta");
                txtNCantPeriodos.Value = 
                    decimal.Parse((Properties.Settings.Default.CantPeriodos - Properties.Settings.Default.CantPP).ToString());
                Properties.Settings.Default.CantPP = 0;
                txtNC03Q06.Select();
            }
        }

        /// <summary>
        /// Carga la información inicial de la configuración.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void LoadData(object sender, EventArgs e)
        {
            try
            {
                double distribuidorasAnt = Properties.Settings.Default.DistribuidorasAnt;
                double clientesXDist = Properties.Settings.Default.ClientesXDist;
                double creditosXDistP = Properties.Settings.Default.CreditosXDistP;
                double clientesDP = Properties.Settings.Default.ClientesDistP / 100;
                double clientesMMP = Properties.Settings.Default.ClientesMMP / 100;
                double clientesZafyP = Properties.Settings.Default.ClientesZafyP / 100;
                double clientesMCP = Properties.Settings.Default.ClientesMCP / 100;
                double carteraTotal = Properties.Settings.Default.CarteraTotal;
                double ctesDistPProd = Properties.Settings.Default.CtesDistPProd / 100;
                double ctesMMPProd = Properties.Settings.Default.CtesMMPProd / 100;
                double ctesCZPProd = Properties.Settings.Default.CtesCZPProd / 100;
                double ctesMCPProd = Properties.Settings.Default.CtesMCPProd / 100;

                if (flujoDBDataSet1.T_Configuraciones.Rows.Count > 0)
                {
                    //Carga los clientes que terminaron credito de 2 y 4 quincenas
                    DataRow[] drL;
                    if (periodoActual >= 3)
                    {
                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CD'" +
                            " AND TipoDato = 'N" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos2QD = double.Parse(dr["Valor"].ToString());
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesDP'" +
                            " AND TipoDato = 'P" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesDP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMMP'" +
                            " AND TipoDato = 'P" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMMP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesZP'" +
                            " AND TipoDato = 'P" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesZP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMC'" +
                            " AND TipoDato = 'P" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }
                        
                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo LIKE '%Q02'" +
                            " AND TipoDato = 'N" + (periodoActual - 2).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos2Q += double.Parse(dr["Valor"].ToString());
                        }

                        if (creditos2Q > 0)
                        {
                            clientes2Credito += creditos2Q;

                            creditos2QCDT = Math.Round(((creditos2Q - creditos2QD) * clientesDP2CPerm));
                            creditos2QMMT = Math.Round(((creditos2Q - creditos2QD) * clientesMMP2CPerm));
                            creditos2QCZT = Math.Round(((creditos2Q - creditos2QD) * clientesZP2CPerm));
                            creditos2QMCT = Math.Round(((creditos2Q - creditos2QD) * clientesMCP2CPerm));
                            creditos2QCD = Math.Round(creditos2QCDT * permanenciaD);
                            creditos2QMM = Math.Round(creditos2QMMT * permanenciaMM);
                            creditos2QCZ = Math.Round(creditos2QCZT * permanenciaCZ);
                            creditos2QMC = Math.Round(creditos2QMCT * permanenciaMC);
                            //___________________________________________________________________________________________________
                            creditos2Q = Math.Round(creditos2QD + creditos2QCD + creditos2QMM + creditos2QCZ + creditos2QMC);

                            dist2Credito += creditos2QD;
                            clientesD2Credito += creditos2QCD;
                            clientesMM2Credito += creditos2QMM;
                            clientesZ2Credito += creditos2QCZ;
                            clientesMC2Credito += creditos2QMC;
                            clientesT2Credito += creditos2Q;

                            clientesDP2C = clientesD2Credito > 0 ?
                                (clientesD2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMMP2C = clientesMM2Credito > 0 ?
                                (clientesMM2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesZafyP2C = clientesZ2Credito > 0 ?
                                (clientesZ2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMCP2C = clientesMC2Credito > 0 ?
                                (clientesMC2Credito / (clientesT2Credito - dist2Credito)) : 0;
                        }
                    }

                    if (periodoActual >= 5)
                    {
                        //Obtiene créditos de 4 quincenas
                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CD'" +
                            " AND TipoDato = 'N" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos4QD = double.Parse(dr["Valor"].ToString());
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesDP'" +
                            " AND TipoDato = 'P" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesDP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMMP'" +
                            " AND TipoDato = 'P" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMMP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesZP'" +
                            " AND TipoDato = 'P" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesZP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMC'" +
                            " AND TipoDato = 'P" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo LIKE '%Q04'" +
                            " AND TipoDato = 'N" + (periodoActual - 4).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos4Q += double.Parse(dr["Valor"].ToString());
                        }

                        if (creditos4Q > 0)
                        {
                            clientes2Credito += creditos4Q;

                            creditos4QCDT = Math.Round(((creditos4Q - creditos4QD) * clientesDP2CPerm));
                            creditos4QMMT = Math.Round(((creditos4Q - creditos4QD) * clientesMMP2CPerm));
                            creditos4QCZT = Math.Round(((creditos4Q - creditos4QD) * clientesZP2CPerm));
                            creditos4QMCT = Math.Round(((creditos4Q - creditos4QD) * clientesMCP2CPerm));
                            creditos4QCD = Math.Round(creditos4QCDT * permanenciaD);
                            creditos4QMM = Math.Round(creditos4QMMT * permanenciaMM);
                            creditos4QCZ = Math.Round(creditos4QCZT * permanenciaCZ);
                            creditos4QMC = Math.Round(creditos4QMCT * permanenciaMC);
                            //___________________________________________________________________________________________________
                            creditos4Q = Math.Round(creditos4QD + creditos4QCD + creditos4QMM + creditos4QCZ + creditos4QMC);

                            dist2Credito += creditos4QD;
                            clientesD2Credito += creditos4QCD;
                            clientesMM2Credito += creditos4QMM;
                            clientesZ2Credito += creditos4QCZ;
                            clientesMC2Credito += creditos4QMC;
                            clientesT2Credito += creditos4Q;
                            
                            clientesDP2C = clientesD2Credito > 0 ? 
                                (clientesD2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMMP2C = clientesMM2Credito > 0 ? 
                                (clientesMM2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesZafyP2C = clientesZ2Credito > 0 ? 
                                (clientesZ2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMCP2C = clientesMC2Credito > 0 ? 
                                (clientesMC2Credito / (clientesT2Credito - dist2Credito)) : 0;
                        }
                        //Fin de créditos de 4Q

                        //Obtiene créditos de 6 quincenas
                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CD'" +
                            " AND TipoDato = 'N" + (periodoActual - 6).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos6QD = double.Parse(dr["Valor"].ToString());
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesDPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 6).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesDP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMMPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 6).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMMP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesZPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 6).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesZP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMCPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 6).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo LIKE '%Q06'" +
                            " AND TipoDato = 'N" + (periodoActual - 6).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos6Q += double.Parse(dr["Valor"].ToString());
                        }

                        if (creditos6Q > 0)
                        {
                            clientes2Credito += creditos6Q;

                            creditos6QCDT = Math.Round(((creditos6Q - creditos6QD) * clientesDP2CPerm));
                            creditos6QMMT = Math.Round(((creditos6Q - creditos6QD) * clientesMMP2CPerm));
                            creditos6QCZT = Math.Round(((creditos6Q - creditos6QD) * clientesZP2CPerm));
                            creditos6QMCT = Math.Round(((creditos6Q - creditos6QD) * clientesMCP2CPerm));
                            //___________________________________________________________________________________________________
                            creditos6Q = Math.Round(creditos6QD + creditos6QCDT + creditos6QMMT + creditos6QCZT + creditos6QMCT);

                            dist2Credito += creditos6QD;
                            clientesD2Credito += creditos6QCDT;
                            clientesMM2Credito += creditos6QMMT;
                            clientesZ2Credito += creditos6QCZT;
                            clientesMC2Credito += creditos6QMCT;
                            clientesT2Credito += creditos6Q;

                            clientesDP2C = clientesD2Credito > 0 ?
                                (clientesD2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMMP2C = clientesMM2Credito > 0 ?
                                (clientesMM2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesZafyP2C = clientesZ2Credito > 0 ?
                                (clientesZ2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMCP2C = clientesMC2Credito > 0 ?
                                (clientesMC2Credito / (clientesT2Credito - dist2Credito)) : 0;
                        }
                        //Fin de créditos de 6Q

                        //Obtiene créditos de 8 quincenas
                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CD'" +
                            " AND TipoDato = 'N" + (periodoActual - 8).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos8QD = double.Parse(dr["Valor"].ToString());
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesDPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 8).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesDP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMMPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 8).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMMP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesZPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 8).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesZP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMCPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 8).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo LIKE '%Q08'" +
                            " AND TipoDato = 'N" + (periodoActual - 8).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos8Q += double.Parse(dr["Valor"].ToString());
                        }

                        if (creditos8Q > 0)
                        {
                            clientes2Credito += creditos8Q;

                            creditos8QCDT = Math.Round(((creditos8Q - creditos8QD) * clientesDP2CPerm));
                            creditos8QMMT = Math.Round(((creditos8Q - creditos8QD) * clientesMMP2CPerm));
                            creditos8QCZT = Math.Round(((creditos8Q - creditos8QD) * clientesZP2CPerm));
                            creditos8QMCT = Math.Round(((creditos8Q - creditos8QD) * clientesMCP2CPerm));
                            //___________________________________________________________________________________________________
                            creditos8Q = Math.Round(creditos8QD + creditos8QCDT + creditos8QMMT + creditos8QCZT + creditos8QMCT);

                            dist2Credito += creditos8QD;
                            clientesD2Credito += creditos8QCDT;
                            clientesMM2Credito += creditos8QMMT;
                            clientesZ2Credito += creditos8QCZT;
                            clientesMC2Credito += creditos8QMCT;
                            clientesT2Credito += creditos8Q;

                            clientesDP2C = clientesD2Credito > 0 ?
                                (clientesD2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMMP2C = clientesMM2Credito > 0 ?
                                (clientesMM2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesZafyP2C = clientesZ2Credito > 0 ?
                                (clientesZ2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMCP2C = clientesMC2Credito > 0 ?
                                (clientesMC2Credito / (clientesT2Credito - dist2Credito)) : 0;
                        }
                        //Fin de créditos de 8Q

                        //Obtiene créditos de 10 quincenas
                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CD'" +
                            " AND TipoDato = 'N" + (periodoActual - 10).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos10QD = double.Parse(dr["Valor"].ToString());
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesDPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 10).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesDP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMMPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 10).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMMP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesZPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 10).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesZP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMCPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 10).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo LIKE '%Q10'" +
                            " AND TipoDato = 'N" + (periodoActual - 10).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos10Q += double.Parse(dr["Valor"].ToString());
                        }

                        if (creditos10Q > 0)
                        {
                            clientes2Credito += creditos10Q;

                            creditos10QCDT = Math.Round(((creditos10Q - creditos10QD) * clientesDP2CPerm));
                            creditos10QMMT = Math.Round(((creditos10Q - creditos10QD) * clientesMMP2CPerm));
                            creditos10QCZT = Math.Round(((creditos10Q - creditos10QD) * clientesZP2CPerm));
                            creditos10QMCT = Math.Round(((creditos10Q - creditos10QD) * clientesMCP2CPerm));
                            //___________________________________________________________________________________________________
                            creditos10Q = Math.Round(creditos10QD + creditos10QCDT + creditos10QMMT + creditos10QCZT + creditos10QMCT);

                            dist2Credito += creditos10QD;
                            clientesD2Credito += creditos10QCDT;
                            clientesMM2Credito += creditos10QMMT;
                            clientesZ2Credito += creditos10QCZT;
                            clientesMC2Credito += creditos10QMCT;
                            clientesT2Credito += creditos10Q;

                            clientesDP2C = clientesD2Credito > 0 ?
                                (clientesD2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMMP2C = clientesMM2Credito > 0 ?
                                (clientesMM2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesZafyP2C = clientesZ2Credito > 0 ?
                                (clientesZ2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMCP2C = clientesMC2Credito > 0 ?
                                (clientesMC2Credito / (clientesT2Credito - dist2Credito)) : 0;
                        }
                        //Fin de créditos de 10Q

                        //Obtiene créditos de 12 quincenas
                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CD'" +
                            " AND TipoDato = 'N" + (periodoActual - 12).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos12QD = double.Parse(dr["Valor"].ToString());
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesDPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 12).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesDP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMMPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 12).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMMP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesZPPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 12).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesZP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo = 'CtesMCPerm'" +
                            " AND TipoDato = 'P" + (periodoActual - 12).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            clientesMCP2CPerm = double.Parse(dr["Valor"].ToString()) / 100;
                        }

                        drL = flujoDBDataSet1.T_Configuraciones.Select(
                            " SesionId = " + Properties.Settings.Default.SessionId +
                            " AND Campo LIKE '%Q12'" +
                            " AND TipoDato = 'N" + (periodoActual - 12).ToString().PadLeft(3, '0') + "C'");

                        foreach (DataRow dr in drL)
                        {
                            creditos12Q += double.Parse(dr["Valor"].ToString());
                        }

                        if (creditos12Q > 0)
                        {
                            clientes2Credito += creditos12Q;

                            creditos12QCDT = Math.Round(((creditos12Q - creditos12QD) * clientesDP2CPerm));
                            creditos12QMMT = Math.Round(((creditos12Q - creditos12QD) * clientesMMP2CPerm));
                            creditos12QCZT = Math.Round(((creditos12Q - creditos12QD) * clientesZP2CPerm));
                            creditos12QMCT = Math.Round(((creditos12Q - creditos12QD) * clientesMCP2CPerm));
                            //___________________________________________________________________________________________________
                            creditos12Q = Math.Round(creditos12QD + creditos12QCDT + creditos12QMMT + creditos12QCZT + creditos12QMCT);

                            dist2Credito += creditos12QD;
                            clientesD2Credito += creditos12QCDT;
                            clientesMM2Credito += creditos12QMMT;
                            clientesZ2Credito += creditos12QCZT;
                            clientesMC2Credito += creditos12QMCT;
                            clientesT2Credito += creditos12Q;

                            clientesDP2C = clientesD2Credito > 0 ?
                                (clientesD2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMMP2C = clientesMM2Credito > 0 ?
                                (clientesMM2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesZafyP2C = clientesZ2Credito > 0 ?
                                (clientesZ2Credito / (clientesT2Credito - dist2Credito)) : 0;
                            clientesMCP2C = clientesMC2Credito > 0 ?
                                (clientesMC2Credito / (clientesT2Credito - dist2Credito)) : 0;
                        }
                        //Fin de créditos de 12Q

                    }

                    Properties.Settings.Default.Clientes2Credito = clientes2Credito;
                    permanencia = ((clientesT2Credito) / clientes2Credito);
                }

                if (periodoActual == 1)
                {
                    lblAportacionAnt.Visible = false;
                    txtAportacionAcum.Visible = false;
                    lblApCapital.Visible = false;
                    btnAgregarCapital.Visible = false;
                    txtMApCapital.Visible = false;
                    lblCtesPeriodoAnt.Visible = false;
                    txtNCtesPeriodoAnt.Visible = false;
                    lblIncremento.Visible = false;
                    txtPIncremento.Visible = false;
                    lblCtes2Credito.Visible = false;
                    txtNCtes2Credito.Visible = false;
                    lblPermanencia.Visible = false;
                    txtPPermanencia.Visible = false;
                    lblCtesPermanencia.Visible = false;
                    txtNCtesPermanencia.Visible = false;

                    clientesNuevos = clientesPeriodoAnt;
                }
                else if (periodoActual > 1)
                {
                    lblAportacionAnt.Visible = true;
                    txtAportacionAcum.Visible = true;
                    lblApCapital.Visible = true;
                    btnAgregarCapital.Visible = true;
                    txtMApCapital.Visible = true;
                    lblCtesPeriodoAnt.Visible = true;
                    txtNCtesPeriodoAnt.Visible = true;
                    lblIncremento.Visible = true;
                    txtPIncremento.Visible = true;
                    //Producción de clientes
                    double dist = Properties.Settings.Default.DistribuidorasAnt * (Properties.Settings.Default.IncrementoDVal);

                    if (Properties.Settings.Default.IsDistribuidorasOn)
                    {
                        Properties.Settings.Default.Distribuidoras = Properties.Settings.Default.Distribuidoras + Math.Truncate(dist);
                    }

                    this.distribuidoras = Properties.Settings.Default.Distribuidoras;

                    ctesDistAnt = Math.Round(((carteraTotal - (distribuidoras - dist)) * clientesDP) - 
                        (creditos2QCDT + creditos4QCDT + creditos6QCDT));
                    ctesDist = Math.Round(distribuidoras * clientesXDist * creditosXDistP);

                    ctesMMAnt = Math.Round(((carteraTotal - (distribuidoras - dist)) * clientesMMP) - 
                        (creditos2QMMT + creditos4QMMT + creditos6QMMT));
                    ctesMM = Math.Round(((clientesPeriodoAnt - distribuidorasAnt) * ctesMMPProd)
                        * Properties.Settings.Default.IncrementoMMVal);

                    ctesCZAnt = Math.Round(((carteraTotal - (distribuidoras - dist)) * clientesZafyP) - 
                        (creditos2QCZT + creditos4QCZT + creditos6QCZT));
                    ctesCZ = Math.Round(((clientesPeriodoAnt - distribuidorasAnt) * ctesCZPProd)
                        * Properties.Settings.Default.IncrementoCZVal);

                    ctesMCAnt = Math.Round(((carteraTotal - (distribuidoras - dist)) * clientesMCP) - 
                        (creditos2QMCT + creditos4QMCT + creditos6QMCT));
                    ctesMC = Properties.Settings.Default.CantMiembros;

                    clientesNuevos = Math.Truncate(
                        dist + ctesDist + ctesMM + ctesCZ + ctesMC);

                    Properties.Settings.Default.CtesDistPProd = (ctesDist / (clientesNuevos - dist)) * 100;
                    Properties.Settings.Default.CtesMMPProd = (ctesMM / (clientesNuevos - dist)) * 100;
                    Properties.Settings.Default.CtesCZPProd = (ctesCZ / (clientesNuevos - dist)) * 100;
                    Properties.Settings.Default.CtesMCPProd = (ctesMC / (clientesNuevos - dist)) * 100;
                    
                    Properties.Settings.Default.DistribuidorasAnt = dist;
                    Properties.Settings.Default.ClientesDistP = 
                        ((ctesDist + clientesD2Credito + ctesDistAnt) / 
                        ((clientesNuevos + creditos2Q + creditos4Q + (carteraTotal - clientes2Credito)) - distribuidoras)) * 100;
                    Properties.Settings.Default.ClientesMMP = 
                        ((ctesMM + clientesMM2Credito + ctesMMAnt) / 
                        ((clientesNuevos + creditos2Q + creditos4Q + (carteraTotal - clientes2Credito)) - distribuidoras)) * 100;
                    Properties.Settings.Default.ClientesZafyP = 
                        ((ctesCZ + clientesZ2Credito + ctesCZAnt) / 
                        ((clientesNuevos + creditos2Q + creditos4Q + (carteraTotal - clientes2Credito)) - distribuidoras)) * 100;
                    Properties.Settings.Default.ClientesMCP = 
                        ((ctesMC + clientesMC2Credito + ctesMCAnt) / 
                        ((clientesNuevos + creditos2Q + creditos4Q + (carteraTotal - clientes2Credito)) - distribuidoras)) * 100;

                    //Identifica a las hijas, madres y nietas.
                                        
                    incremento = (clientesNuevos / clientesPeriodoAnt);

                    if (capital < 0)
                    {
                        apCapital = Math.Ceiling(capital * -1);
                        txtMApCapital.Text = apCapital.ToString();
                        Properties.Settings.Default.Capital = 0;
                        capital = 0;
                    }
                }

                if (periodoActual <= 4)
                {
                    lblCredito03.Visible = false;
                    lblCredito04.Visible = false;
                    lblCredito05.Visible = false;
                    lblCredito06.Visible = false;

                    txtNC03Q06.Visible = false;
                    txtNC03Q08.Visible = false;
                    txtNC04Q06.Visible = false;
                    txtNC04Q08.Visible = false;
                    txtNC04Q10.Visible = false;
                    txtNC05Q08.Visible = false;
                    txtNC05Q10.Visible = false;
                    txtNC05Q12.Visible = false;
                    txtNC06Q10.Visible = false;
                    txtNC06Q12.Visible = false;

                    lblQ06.Visible = false;
                    lblQ08.Visible = false;
                    lblQ10.Visible = false;
                    lblQ12.Visible = false;
                }
                else
                {
                    lblCredito03.Visible = true;
                    lblCredito04.Visible = true;
                    lblCredito05.Visible = true;
                    lblCredito06.Visible = true;

                    txtNC03Q06.Visible = true;
                    txtNC03Q08.Visible = true;
                    txtNC04Q06.Visible = true;
                    txtNC04Q08.Visible = true;
                    txtNC04Q10.Visible = true;
                    txtNC05Q08.Visible = true;
                    txtNC05Q10.Visible = true;
                    txtNC05Q12.Visible = true;
                    txtNC06Q10.Visible = true;
                    txtNC06Q12.Visible = true;

                    lblQ06.Visible = true;
                    lblQ08.Visible = true;
                    lblQ10.Visible = true;
                    lblQ12.Visible = true;
                }

                if (periodoActual <= 2)
                {
                    lblCtes2Credito.Visible = false;
                    txtNCtes2Credito.Visible = false;
                    lblPermanencia.Visible = false;
                    txtPPermanencia.Visible = false;
                    lblCtesPermanencia.Visible = false;
                    txtNCtesPermanencia.Visible = false;
                }
                else
                {
                    lblCtes2Credito.Visible = true;
                    txtNCtes2Credito.Visible = true;
                    lblPermanencia.Visible = true;
                    txtPPermanencia.Visible = true;
                    lblCtesPermanencia.Visible = true;
                    txtNCtesPermanencia.Visible = true;

                    clientesPermanencia = Math.Truncate(clientesT2Credito);
                }

                double clientesPermanenciaAnt = Properties.Settings.Default.ClientesPermanenciaAnt;

                if (Properties.Settings.Default.IsAutomatic && Properties.Settings.Default.CantPeriodos > 1)
                {
                    //Carga la configuración anterior
                    DataRow[] colDt = flujoDBDataSet1.T_Configuraciones.Select(
                        " SesionId = " + Properties.Settings.Default.SessionId +
                        " AND Campo LIKE 'C%'" +
                        " AND TipoDato = 'N" + (periodoActual - 1).ToString().PadLeft(3, '0') + "C'");
                    double valorAnt = 0;
                    double valorNvo = 0;

                    foreach (FlujoDBDataSet.T_ConfiguracionesRow item in colDt)
                    {
                        if (item.Campo.Trim().Length == 6)
                        {
                            Control[] ctrl = Controls.Find("txtN" + item.Campo.Trim(), true);

                            valorAnt = double.Parse(item.Valor);

                            if (int.Parse(item.Campo.Substring(1, 2)) > 2)
                            {
                                if (clientesPermanenciaAnt > 0)
                                {
                                    valorNvo = Math.Truncate(clientesPermanencia * (valorAnt / clientesPermanenciaAnt));
                                }
                            }
                            else
                            {
                                valorNvo = Math.Truncate((clientesNuevos + clientesPermanencia) * (valorAnt / clientesPeriodoAnt));
                            }

                            ctrl[0].Text = valorNvo.ToString();
                        }
                    }
                }

                Properties.Settings.Default.ClientesPermanenciaAnt = clientesPermanencia;

                lblCredito01.Text = string.Format("{0:C2}", c1);
                lblCredito02.Text = string.Format("{0:C2}", c2);
                lblCredito03.Text = string.Format("{0:C2}", c3);
                lblCredito04.Text = string.Format("{0:C2}", c4);
                lblCredito05.Text = string.Format("{0:C2}", c5);
                lblCredito06.Text = string.Format("{0:C2}", c6);

                txtMCapital.Text = string.Format("{0:C0}", capital);
                txtNCtesPeriodoAnt.Text = string.Format("{0:N0}", clientesPeriodoAnt);
                txtPIncremento.Text = string.Format("{0:P2}", incremento);
                txtNCtesNuevos.Text = string.Format("{0:N0}", clientesNuevos);
                txtNCtes2Credito.Text = string.Format("{0:N0}", clientes2Credito);
                txtPPermanencia.Text = string.Format("{0:P0}", permanencia.ToString() == "NaN" ? 0 : permanencia);
                txtNCtesPermanencia.Text = string.Format("{0:N0}", clientesPermanencia);

                txtAportacionAcum.Text = string.Format("{0:C0}", aportacionAcum);
                txtMApCapital.Text = apCapitalTemp > 0 ? apCapitalTemp.ToString() : apCapital.ToString();
                txtPPerdida.Text = perdidaE.ToString();
                txtPComisionDist.Text = comisionDistE.ToString();
                txtMGastosFijosPROSA.Text = gastosFijosPROSAE.ToString();
                txtMGastosVarPROSA.Text = gastosVarPROSAE.ToString();
                txtMGastosFijosZafy.Text = gastosFijosZafyE.ToString();
                txtMGastosVarZafy.Text = gastosVarZafyE.ToString();
                txtMGastosXPublicidad.Text = gastosXPublicidad.ToString();
                txtPBonosPremios.Text = bonosPremiosE.ToString();
                txtPRetiros.Text = retiroE.ToString();

                txtCarteraTotal.Text = string.Format("{0:N0}", carteraTotal);

                txtNC01Q02.Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }

            
        }

        /// <summary>
        /// Guarda la configuración inicial.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void SaveInitialConfig(object sender, EventArgs e)
        {
            try
            {
                foreach (Control control in this.Controls)
                {
                    if (control.Name == "gpbConfiguracion")
                    {
                       foreach(Control config in control.Controls)
                       {
                            if (config.Name.Substring(0, 3) == "txt")
                            {
                                string tipoDato = config.Name.Substring(3, 1);
                                tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();

                                FlujoDBDataSet.T_ConfiguracionesRow tcrId;
                                if (flujoDBDataSet1.T_Configuraciones.Rows.Count > 0)
                                {
                                    tcrId =
                                    (FlujoDBDataSet.T_ConfiguracionesRow)flujoDBDataSet1.T_Configuraciones.Rows[
                                        flujoDBDataSet1.T_Configuraciones.Rows.Count - 1];

                                    configId = int.Parse(tcrId["Id"].ToString()) + 1;

                                }

                                //Asigna los valores del nuevo campo
                                tConfiguracionesRow["Id"] = configId;
                                tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                                tConfiguracionesRow["Campo"] = config.Name.Substring(4).Trim();
                                tConfiguracionesRow["Valor"] = config.Text.Trim().Replace("$", "").Replace(",", "").Replace(".", "").Replace("%", "");
                                tConfiguracionesRow["TipoDato"] =
                                    tipoDato +
                                    (Properties.Settings.Default.PeriodoActual > 0 ?
                                    Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                                    "I";
                                tConfiguracionesRow["Estatus"] = "1";

                                flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);
                            }
                        } 
                    }
                }

                int result = 0;

                foreach (FlujoDBDataSet.T_ConfiguracionesRow dr in flujoDBDataSet1.T_Configuraciones.Rows)
                {
                    result = t_ConfiguracionesTableAdapter1.UpdateTConfiguracion(
                        int.Parse(dr["Id"].ToString()), int.Parse(dr["SesionId"].ToString()),
                        dr["Campo"].ToString(), dr["Valor"].ToString(), dr["TipoDato"].ToString(), short.Parse(dr["Estatus"].ToString()));

                    if (result == 0)
                    {
                        result = t_ConfiguracionesTableAdapter1.InsertTConfiguracion(
                        int.Parse(dr["Id"].ToString()), int.Parse(dr["SesionId"].ToString()),
                        dr["Campo"].ToString(), dr["Valor"].ToString(), dr["TipoDato"].ToString(), short.Parse(dr["Estatus"].ToString()));
                    }

                    result = 0;
                }

                flujoDBDataSet1.AcceptChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        /// <summary>
        /// Guarda los datos de la configuración final.
        /// </summary>
        private void SaveFinalConfig()
        {
            try
            {
                foreach (Control control in this.Controls)
                {
                    if (control.Name == "gpbConfiguracion")
                    {
                       foreach(Control config in control.Controls)
                        {
                            if (config.Name.Substring(0, 3) == "txt")
                            {
                                string tipoDato = config.Name.Substring(3, 1);
                                tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();

                                FlujoDBDataSet.T_ConfiguracionesRow tcrId;
                                if (flujoDBDataSet1.T_Configuraciones.Rows.Count > 0)
                                {
                                    tcrId =
                                    (FlujoDBDataSet.T_ConfiguracionesRow)flujoDBDataSet1.T_Configuraciones.Rows[
                                        flujoDBDataSet1.T_Configuraciones.Rows.Count - 1];

                                    configId = int.Parse(tcrId["Id"].ToString()) + 1;

                                }

                                //Asigna los valores del nuevo campo
                                tConfiguracionesRow["Id"] = configId;
                                tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                                tConfiguracionesRow["Campo"] = config.Name.Substring(4).Trim();
                                tConfiguracionesRow["Valor"] = config.Text.Trim().Replace("$", "").Replace(",", "").Replace(".", "").Replace("%", "");
                                tConfiguracionesRow["TipoDato"] = 
                                    tipoDato +
                                    (Properties.Settings.Default.PeriodoActual > 0 ?
                                    Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                                    "F";
                                tConfiguracionesRow["Estatus"] = "1";

                                flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);
                            }
                        } 
                    }
                }

                int result = 0;

                foreach (FlujoDBDataSet.T_ConfiguracionesRow dr in flujoDBDataSet1.T_Configuraciones.Rows)
                {
                    result = t_ConfiguracionesTableAdapter1.UpdateTConfiguracion(
                        int.Parse(dr["Id"].ToString()), int.Parse(dr["SesionId"].ToString()),
                        dr["Campo"].ToString(), dr["Valor"].ToString(), dr["TipoDato"].ToString(), short.Parse(dr["Estatus"].ToString()));

                    if (result == 0)
                    {
                        result = t_ConfiguracionesTableAdapter1.InsertTConfiguracion(
                        int.Parse(dr["Id"].ToString()), int.Parse(dr["SesionId"].ToString()),
                        dr["Campo"].ToString(), dr["Valor"].ToString(), dr["TipoDato"].ToString(), short.Parse(dr["Estatus"].ToString()));
                    }

                    result = 0;
                }

                flujoDBDataSet1.AcceptChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        /// <summary>
        /// Guarda la colocación en la base de datos.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void btnSaveColocacion_Click(object sender, EventArgs e)
        {
            try
            {
                bool isToApCap = true;

                if (double.Parse(txtMApCapital.Text) > 0)
                {
                    if (Properties.Settings.Default.IsAutomatic)
                    {
                        this.btnAgregarCapital_Click(sender, e);
                    }
                    else
                    {
                        if (DialogResult.Yes == MessageBox.Show("¿Desea agregar " +
                        string.Format("{0:C0}", double.Parse(txtMApCapital.Text)) +
                        " al capital?", "Aportación", MessageBoxButtons.YesNo))
                        {
                            this.btnAgregarCapital_Click(sender, e);
                        }
                        else
                        {
                            isToApCap = false;
                        }
                    }
                    
                }

                if (colocacion == capital)
                {
                    txtMCapital.Text = "0";
                }
                else if (capital > 0 && capital < colocacion)
                {
                    if (isApCapCte)
                    {
                        MessageBox.Show(
                            "Coloque los créditos necesarios para no exceder el capital o haga una aportación", 
                            "Capital Insuficiente");
                        isToApCap = false;
                    }
                }

                if (isToApCap)
                {
                    foreach (Control config in this.Controls)
                    {
                        if (config.Name == "gpbColocacion")
                        {
                            foreach (Control dataG in config.Controls)
                            {
                                if (dataG is GroupBox)
                                {
                                    foreach (Control data in dataG.Controls)
                                    {
                                        if (data.Name.Substring(0, 3) == "txt")
                                        {
                                            string tipoDato = data.Name.Substring(3, 1);
                                            tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();

                                            FlujoDBDataSet.T_ConfiguracionesRow tcrId;
                                            if (flujoDBDataSet1.T_Configuraciones.Rows.Count > 0)
                                            {
                                                tcrId =
                                                (FlujoDBDataSet.T_ConfiguracionesRow)flujoDBDataSet1.T_Configuraciones.Rows[
                                                    flujoDBDataSet1.T_Configuraciones.Rows.Count - 1];

                                                configId = int.Parse(tcrId["Id"].ToString()) + 1;
                                            }

                                            if (data.Name.Substring(0, 5) == "txtNC" && data.Text != "0")
                                            {
                                                this.generaTablaDeAmortizacion(data);
                                            }

                                            //Asigna los valores del nuevo campo
                                            tConfiguracionesRow["Id"] = configId;
                                            tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                                            tConfiguracionesRow["Campo"] = data.Name.Substring(4).Trim();
                                            tConfiguracionesRow["Valor"] = data.Text.Trim().Replace("$", "").Replace(",", "").Replace(".", "");
                                            tConfiguracionesRow["TipoDato"] =
                                                tipoDato +
                                                (Properties.Settings.Default.PeriodoActual > 0 ?
                                                Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                                                "C";
                                            tConfiguracionesRow["Estatus"] = "1";

                                            flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //Guarda la cantidad de distribuidoras para el período actual
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CD";
                    tConfiguracionesRow["Valor"] = Properties.Settings.Default.DistribuidorasAnt;
                    tConfiguracionesRow["TipoDato"] =
                        "N" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    configId++;
                    tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CDT";
                    tConfiguracionesRow["Valor"] = Properties.Settings.Default.DistribuidorasAnt + dist2Credito;
                    tConfiguracionesRow["TipoDato"] =
                        "N" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);
                    
                    Properties.Settings.Default.LastIdCliente = clienteID;
                    Properties.Settings.Default.isColocacionConfigured = true;
                    Properties.Settings.Default.ClientesNuevos = creditosCN;
                    Properties.Settings.Default.Clientes2Credito = creditosCA;
                    Properties.Settings.Default.ColocacionE = colocacion;

                    if (periodoActual == 1)
                    {
                        Properties.Settings.Default.ApCapital = double.Parse(txtMCapital.Text.Replace("$", "").Replace(",", ""));
                    }

                    Properties.Settings.Default.PerdidaE = double.Parse(txtPPerdida.Text);
                    Properties.Settings.Default.ComisionDistE = double.Parse(txtPComisionDist.Text);
                    Properties.Settings.Default.GastosFijosPROSAE = double.Parse(txtMGastosFijosPROSA.Text);
                    Properties.Settings.Default.GastosVarPROSAE = double.Parse(txtMGastosVarPROSA.Text);
                    Properties.Settings.Default.GastosFijosZafyE = double.Parse(txtMGastosFijosZafy.Text);
                    Properties.Settings.Default.GastosVarZafyE = double.Parse(txtMGastosVarZafy.Text);
                    Properties.Settings.Default.GastosXPublicidadE = double.Parse(txtMGastosXPublicidad.Text);
                    Properties.Settings.Default.BonosPremiosE = double.Parse(txtPBonosPremios.Text);
                    Properties.Settings.Default.RetirosE = double.Parse(txtPRetiros.Text);
                    //Verifica colocación y genera porcentajes en caso de variación
                    double dist = Properties.Settings.Default.DistribuidorasAnt;

                    double ctesTotales = creditosCN + creditosCA;

                    if (periodoActual > 1 && ctesTotales != clientesNuevos)
                    {
                        ctesDist = Math.Round((creditosCN - dist) * (Properties.Settings.Default.CtesDistPProd / 100));
                        ctesMM = Math.Round((creditosCN - dist) * (Properties.Settings.Default.CtesMMPProd / 100));
                        ctesCZ = Math.Round((creditosCN - dist) * (Properties.Settings.Default.CtesCZPProd / 100));
                        ctesMC = Math.Round((creditosCN - dist) * (Properties.Settings.Default.CtesMCPProd / 100));

                        double ctesDistPerm = Math.Round((creditosCA - dist2Credito) * clientesDP2C);
                        double ctesMMPerm = Math.Round((creditosCA - dist2Credito) * clientesMMP2C);
                        double ctesCZPerm = Math.Round((creditosCA - dist2Credito) * clientesZafyP2C);
                        double ctesMCPerm = Math.Round((creditosCA - dist2Credito) * clientesMCP2C);

                        Properties.Settings.Default.CtesDistPProd = (ctesDist / (creditosCN - dist)) * 100;
                        Properties.Settings.Default.CtesMMPProd = (ctesMM / (creditosCN - dist)) * 100;
                        Properties.Settings.Default.CtesCZPProd = (ctesCZ / (creditosCN - dist)) * 100;
                        Properties.Settings.Default.CtesMCPProd = (ctesMC / (creditosCN - dist)) * 100;

                        Properties.Settings.Default.CtesDistPPerm = (ctesDistPerm / (creditosCA - dist2Credito)) * 100;
                        Properties.Settings.Default.CtesMMPPerm = (ctesMMPerm / (creditosCA - dist2Credito)) * 100;
                        Properties.Settings.Default.CtesCZPPerm = (ctesCZPerm / (creditosCA - dist2Credito)) * 100;
                        Properties.Settings.Default.CtesMCPPerm = (ctesMCPerm / (creditosCA - dist2Credito)) * 100;

                        Properties.Settings.Default.ClientesDistP = 
                            ((ctesDist + ctesDistPerm + ctesDistAnt) / 
                            ((ctesTotales + (carteraTotal - clientes2Credito)) - distribuidoras)) * 100;
                        Properties.Settings.Default.ClientesMMP = 
                            ((ctesMM + ctesMMPerm + ctesMMAnt) / 
                            ((ctesTotales + (carteraTotal - clientes2Credito)) - distribuidoras)) * 100;
                        Properties.Settings.Default.ClientesZafyP = 
                            ((ctesCZ + ctesCZPerm + ctesCZAnt) / 
                            ((ctesTotales + (carteraTotal - clientes2Credito)) - distribuidoras)) * 100;
                        Properties.Settings.Default.ClientesMCP = 
                            ((ctesMC + ctesMCPerm + ctesMCAnt) / 
                            ((ctesTotales + (carteraTotal - clientes2Credito)) - distribuidoras)) * 100;
                    }

                    //Guarda las proporciones de los clientes para el período
                    //Clientes Distribuidoras
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesDP";
                    tConfiguracionesRow["Valor"] = Properties.Settings.Default.CtesDistPProd;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    //Clientes Medios Masivos
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesMMP";
                    tConfiguracionesRow["Valor"] = Properties.Settings.Default.CtesMMPProd;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    //Clientes Zafy
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesZP";
                    tConfiguracionesRow["Valor"] = Properties.Settings.Default.CtesCZPProd;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    //Clientes Miembros de Célula
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesMC";
                    tConfiguracionesRow["Valor"] = Properties.Settings.Default.CtesMCPProd;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    //Clientes Distribuidoras de permanencia
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesDPPerm";
                    tConfiguracionesRow["Valor"] = Properties.Settings.Default.CtesDistPPerm;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    //Clientes Medios Masivos de permanencia
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesMMPPerm";
                    tConfiguracionesRow["Valor"] = Properties.Settings.Default.CtesMMPPerm;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    //Clientes Zafy de permanencia
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesZPPerm";
                    tConfiguracionesRow["Valor"] = Properties.Settings.Default.CtesCZPPerm;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    //Clientes Miembros de Célula de permanencia
                    configId++;
                    tConfiguracionesRow = flujoDBDataSet1.T_Configuraciones.NewT_ConfiguracionesRow();
                    tConfiguracionesRow["Id"] = configId;
                    tConfiguracionesRow["SesionId"] = Properties.Settings.Default.SessionId.ToString().Trim();
                    tConfiguracionesRow["Campo"] = "CtesMCPerm";
                    tConfiguracionesRow["Valor"] = Properties.Settings.Default.CtesMCPPerm;
                    tConfiguracionesRow["TipoDato"] =
                        "P" +
                        (Properties.Settings.Default.PeriodoActual > 0 ?
                        Properties.Settings.Default.PeriodoActual.ToString().PadLeft(3, '0') : "000") +
                        "C";
                    tConfiguracionesRow["Estatus"] = "1";

                    flujoDBDataSet1.T_Configuraciones.AddT_ConfiguracionesRow(tConfiguracionesRow);

                    Properties.Settings.Default.Save();

                    t_ConfiguracionesTableAdapter1.ClearBeforeFill = true;
                    
                    this.SaveFinalConfig();

                    
                }
                //Fin de colocación
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Colocación no guardada");
            }
        }

        /// <summary>
        /// Cancela la colocación.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void btnCancelColocacion_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Genera las tablas de amortización.
        /// </summary>
        /// <param name="data">Los datos a guardar</param>
        private void generaTablaDeAmortizacion(Control data)
        {
            try
            {
                int c = int.Parse(data.Name.Substring(5, 2));
                int q = int.Parse(data.Name.Substring(8, 2));

                int periodoInicial = int.Parse(periodoActual.ToString());
                int periodoFinal = (periodoInicial + q) - 1;
                double montoCredito = 0;
                double interesTotal = 0;
                int cantidad = int.Parse(data.Text);
                double montoTotal = 0;
                double tasa = 0;
                int numeroPagos = q;
                double capital = 0;
                double interes = 0;
                double saldo = 0;

                switch (c)
                {
                    case 1:
                        montoCredito = c1;
                        break;
                    case 2:
                        montoCredito = c2;
                        break;
                    case 3:
                        montoCredito = c3;
                        break;
                    case 4:
                        montoCredito = c4;
                        break;
                    case 5:
                        montoCredito = c5;
                        break;
                    case 6:
                        montoCredito = c6;
                        break;
                }

                switch (q)
                {
                    case 2:
                        tasa = i02;
                        break;
                    case 4:
                        tasa = i04;
                        break;
                    case 6:
                        tasa = i06;
                        break;
                    case 8:
                        tasa = i08;
                        break;
                    case 10:
                        tasa = i10;
                        break;
                    case 12:
                        tasa = i12;
                        break;

                }
                interesTotal = montoCredito * tasa;
                montoTotal = montoCredito + interesTotal;
                capital = montoCredito / q;
                interes = interesTotal / q;
                
                for (int j = 0; j < cantidad; j++)
                {
                    clienteIDList.Add(clienteID);
                    numeroPagos = q;
                    saldo = montoTotal;
                    for (int i = 0; i < q; i++)
                    {

                        t_AmortizacionesTableAdapter1.InsertAmortizacion(
                            Properties.Settings.Default.SessionId,
                            periodoInicial.ToString(),
                            periodoFinal.ToString(),
                            montoCredito.ToString(),
                            cantidad,
                            montoTotal.ToString(),
                            int.Parse((tasa * 100).ToString()),
                            numeroPagos,
                            capital.ToString(),
                            interes.ToString(),
                            (saldo -= (capital + interes)).ToString(),
                            1, clienteID);
                            numeroPagos -= 1;
                    }
                        clienteID++;
                }

                flujoDBDataSet1.AcceptChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        /// <summary>
        /// Agrega la aportación al capital
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void btnAgregarCapital_Click(object sender, EventArgs e)
        {

            apCapital = double.Parse(txtMApCapital.Text);
            apCapitalAcum += apCapital;
            capital = double.Parse(txtMCapital.Text.Replace("$", "").Replace(",", "").Replace(".", ""));

            if (apCapital > 0)
            {

                Properties.Settings.Default.ApCapital = apCapitalAcum;

                capital = Math.Round(capital + apCapital);

                txtMApCapital.Text = "0";
                txtMCapital.Text = string.Format("{0:C0}", capital);

                txtMApCapital.Enabled = false;
                btnAgregarCapital.Enabled = false;
                isApCapCte = true;
            }
        }
        #endregion

        #region Metodos de Validación

        /// <summary>
        /// Cálcula los créditos y el monto de colocación.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void checkColocacion(object sender, EventArgs e)
        {
            TextBox t = (TextBox)sender;
            apCapital = Properties.Settings.Default.ApCapital;
            capital = Properties.Settings.Default.Capital;
            creditosCN = 0;
            creditosCA = 0;
            colocacion = 0;
            creditosI = t.Text != "" ? double.Parse(t.Text) : 0;
            creditosRestantesN = 0;
            creditosRestantesA = 0;
            capitalRestante = 0;
            apCapitalTemp = 0;

            if(isApCapCte)
            {
                capital = Math.Round(capital + apCapital);
            }
            //Suma los créditos y la colocación
            this.SumaCreditos(this);

            this.SumaColocacion(this);
            /*
            if (creditos2Q > 0)
            {
                if(creditosCN >= creditos2Q)
                {
                    creditosRestantesN = clientesNuevos - (creditosCN - creditos2Q);
                    creditosRestantesA = clientesPermanencia - (creditosCA + creditos2Q);
                }
                else
                {
                    creditosRestantesN = clientesNuevos - creditosCN;
                    creditosRestantesA = clientesPermanencia - creditosCA;
                }
            }
            else
            {
                
            }
            */
            creditosRestantesN = clientesNuevos - creditosCN;
            creditosRestantesA = clientesPermanencia - creditosCA;
            
            if (creditosRestantesN <= 0)
            {
                incremento = ((creditosCN / clientesPeriodoAnt) - 1);
            }
            else
            {
                incremento = (clientesNuevos / clientesPeriodoAnt);
            }

            if (incremento < 0)
            {
                incremento = 0;
            }
            
            if(periodoActual <= 2)
            {
                if (creditosRestantesN < 0)
                {
                    creditosRestantesN = 0;
                }
                
                if (periodoActual == 1)
                {
                    capitalRestante = colocacion;
                }
                else
                {
                    capitalRestante = capital - colocacion;

                    if (capitalRestante < 0)
                    {
                        apCapitalTemp = (capitalRestante * -1);

                        capitalRestante = capital;

                        txtMApCapital.Enabled = true;
                        btnAgregarCapital.Enabled = true;
                        isApCapCte = false;
                    }
                }
            }
            else
            {
                if (creditosRestantesA < 0)
                {
                    if (creditosRestantesN < 0)
                    {
                        creditosRestantesN = 0;
                    }
                    t.Text = "0";

                    creditosRestantesA += creditosI;

                    if (isApCapCte)
                    {
                        capital = Math.Round(capital + apCapital);
                    }

                    capitalRestante = capital - colocacion;

                    if (capitalRestante < 0)
                    {
                        apCapitalTemp = (capitalRestante * -1);

                        capitalRestante = capital;

                        txtMApCapital.Enabled = true;
                        btnAgregarCapital.Enabled = true;
                        isApCapCte = false;
                    }

                    MessageBox.Show(
                        "No se puede exceder el número de créditos para clientes candidatos a segundo crédito", 
                        "Clientes segundo crédito");
                }
                else
                {
                    if (creditosRestantesN < 0)
                    {
                        creditosRestantesN = 0;
                    }

                    capitalRestante = capital - colocacion;

                    if (capitalRestante < 0)
                    {
                        apCapitalTemp = (capitalRestante * -1);

                        capitalRestante = capital;

                        txtMApCapital.Enabled = true;
                        btnAgregarCapital.Enabled = true;
                        isApCapCte = false;
                    }
                }
            }

            txtNCreditosT.Text = (creditosCA + creditosCN).ToString();
            txtMColocacion.Text = string.Format("{0:C0}", colocacion);
            txtMApCapital.Text = apCapitalTemp.ToString();
            txtMCapital.Text = string.Format("{0:C0}", capitalRestante);
            txtNCtesNuevos.Text = string.Format("{0:N0}", creditosRestantesN);
            txtNCtesPermanencia.Text = string.Format("{0:N0}", creditosRestantesA);
            txtPIncremento.Text = string.Format("{0:P2}", incremento);
        }

        /// <summary>
        /// Valida que sean números
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        public void checkNumbers(object sender, KeyPressEventArgs e)
        {
            TextBox txtB = (TextBox)sender;

            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }

            if(txtB.Name == "txtMApCapital" && e.KeyChar == 13)
            {
                this.btnAgregarCapital_Click(sender, e);
            }
        }

        /// <summary>
        /// Valida que no este vacio el campo y que no exceda el 100%
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        public void checkEmpty(object sender, KeyEventArgs e)
        {
            TextBox txtB = (TextBox)sender;
            
            if (txtB.Text == "")
            {
                ValidaCampo(this, txtB);
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
                else if (control is TabControl || control is TabPage || control is GroupBox)
                {
                    ValidaCampo(control, txtB);
                }
            }
        }

        /// <summary>
        /// Suma el número de créditos colocados.
        /// </summary>
        /// <param name="ctrl">El campo a sumar</param>
        public void SumaCreditos(Control ctrl)
        {
            foreach (Control control in ctrl.Controls)
            {
                double CtesPermanencia = clientesPermanencia - creditosCA;

                if (control is TextBox && control.Text != "" && control.Name.Substring(0, 5) == "txtNC")
                {
                    if (int.Parse(control.Name.Substring(8, 2)) <= 4)
                    {
                        if (CtesPermanencia > 0)
                        {
                            int val = int.Parse(control.Text);

                            if (val > 0)
                            {
                                if (clientesPermanencia < creditosCA && clientesPermanencia >= val)
                                {
                                    creditosCA += val;
                                }
                                else 
                                {
                                    creditosCA = clientesPermanencia;
                                    creditosCN = val - clientesPermanencia;
                                }
                            }
                        }
                        else
                        {
                            creditosCN += int.Parse(control.Text);
                        }
                    }
                    else if (int.Parse(control.Name.Substring(8, 2)) > 4)
                    {
                        creditosCA += int.Parse(control.Text);
                    }
                }
                else if (control is TabControl || control is TabPage || control is GroupBox)
                {
                    SumaCreditos(control);
                }
            }
        }

        /// <summary>
        /// Suma los montos de los créditos colocados.
        /// </summary>
        /// <param name="ctrl">El campo a sumar</param>
        private void SumaColocacion(Control ctrl)
        {
            foreach (Control control in ctrl.Controls)
            {
                if (control is TextBox && control.Text != "" && control.Name.Substring(0, 5) == "txtNC")
                {
                    int c = int.Parse(control.Name.Substring(5, 2));

                    switch (c)
                    {
                        case 1:
                            colocacion += int.Parse(control.Text) * c1;
                            break;
                        case 2:
                            colocacion += int.Parse(control.Text) * c2;
                            break;
                        case 3:
                            colocacion += int.Parse(control.Text) * c3;
                            break;
                        case 4:
                            colocacion += int.Parse(control.Text) * c4;
                            break;
                        case 5:
                            colocacion += int.Parse(control.Text) * c5;
                            break;
                        case 6:
                            colocacion += int.Parse(control.Text) * c6;
                            break;
                    }
                }
                else if (control is TabControl || control is TabPage || control is GroupBox)
                {
                    SumaColocacion(control);
                }
            }
        }

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
        /// Asigna la cantidad de períodos a procesar.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void txtNCantPeriodos_ValueChanged(object sender, EventArgs e)
        {
            NumericUpDown nud = (NumericUpDown)sender;

            Properties.Settings.Default.CantPeriodos = double.Parse(nud.Value.ToString());
        }

        /// <summary>
        /// Asigna si será automático los siguientes períodos.
        /// </summary>
        /// <param name="sender">El objeto que llama la función</param>
        /// <param name="e">Los eventos</param>
        private void FormColocacion_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Properties.Settings.Default.CantPeriodos > 1)
            {
                Properties.Settings.Default.IsAutomatic = true;
            }
            else
            {
                if (!Properties.Settings.Default.IsToFinish)
                {
                    Properties.Settings.Default.IsAutomatic = false;
                }
            }
        }

        #endregion
    }
}
