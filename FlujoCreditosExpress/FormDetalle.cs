using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FlujoCreditosExpress
{
    public partial class FormDetalle : Form
    {
        public FormDetalle()
        {
            InitializeComponent();
        }

        private void FormDetalle_Load(object sender, EventArgs e)
        {
            try
            {
                Text = "Detalles del período " + (Properties.Settings.Default.PeriodoActual - 1);
                double carteraT = Properties.Settings.Default.CarteraTotal;
                double distribuidoras = Properties.Settings.Default.Distribuidoras;
                double clientesDP = Properties.Settings.Default.ClientesDistP / 100;
                double clientesMMP = Properties.Settings.Default.ClientesMMP / 100;
                double clientesZP = Properties.Settings.Default.ClientesZafyP / 100;
                double clientesMCP = Properties.Settings.Default.ClientesMCP / 100;
                double ctesDist = Math.Round((carteraT - distribuidoras) * clientesDP);
                double ctesMM = Math.Round((carteraT - distribuidoras) * clientesMMP);
                double ctesCZ = Math.Round((carteraT - distribuidoras) * clientesZP);
                double ctesMC = Math.Round((carteraT - distribuidoras) * clientesMCP);
                double hijasP = Properties.Settings.Default.HijasP / 100;
                double nietasP = Properties.Settings.Default.NietasP / 100;
                double bisnietasP = Properties.Settings.Default.BisnietasP / 100;
                double limiteH = Properties.Settings.Default.HijasC;
                double limiteN = Properties.Settings.Default.NietasC;
                double limiteB = Properties.Settings.Default.BisnietasC;
                double limiteRed = Properties.Settings.Default.MiembrosC;
                double progresoRed = limiteRed > 0 ?
                    (((ctesMC > 0 ? (ctesMC * hijasP) : 0) + (ctesMC > 0 ? (ctesMC * nietasP) : 0) + (ctesMC > 0 ? (ctesMC * bisnietasP) : 0)) /
                    limiteRed * 100) : 0;

                txtCarteraTotal.Text = string.Format("{0:N0}", carteraT);

                txtDistribuidorasDet.Text = string.Format("{0:N0}", distribuidoras);
                txtCtesDsitribuidorasDet.Text = string.Format("{0:N0}", ctesDist);
                txtCtesMediosMasivosDet.Text = string.Format("{0:N0}", ctesMM);
                txtCtesZafyDet.Text = string.Format("{0:N0}", ctesCZ);

                txtLideresIniciadorasDet.Text = string.Format("{0:N0}", Properties.Settings.Default.CantIniciadoras);
                txtLideresDet.Text = string.Format("{0:N0}", Properties.Settings.Default.CantLideresH);
                txtMiembrosCelulaDet.Text = string.Format("{0:N0}", ctesMC);

                pgrBarCT.Value = int.Parse((carteraT > 0 ? (carteraT / carteraT) * 100 : 0).ToString());
                pgrBarD.Value = int.Parse((carteraT > 0 ? Math.Round((distribuidoras / carteraT) * 100) : 0).ToString());
                pgrBarCD.Value = int.Parse((carteraT > 0 ? Math.Round((ctesDist / carteraT) * 100) : 0).ToString());
                pgrBarMM.Value = int.Parse((carteraT > 0 ? Math.Round((ctesMM / carteraT) * 100) : 0).ToString());
                pgrBarCZ.Value = int.Parse((carteraT > 0 ? Math.Round((ctesCZ / carteraT) * 100) : 0).ToString());
                pgrBarMC.Value = int.Parse((carteraT > 0 ? Math.Round((ctesMC / carteraT) * 100) : 0).ToString());

                txtMadresProd.Text = string.Format("{0:N0}", Properties.Settings.Default.CantIniciadoras);
                txtHijasProd.Text = string.Format("{0:N0}", ctesMC > 0 ? (ctesMC * hijasP) : 0);
                picBoxH.Visible = double.Parse(txtHijasProd.Text) == limiteH && limiteH > 0 ? true : false;
                txtNietasProd.Text = string.Format("{0:N0}", ctesMC > 0 ? (ctesMC * nietasP) : 0);
                picBoxN.Visible = double.Parse(txtNietasProd.Text) == limiteN && limiteN > 0 ? true : false;
                txtBisnietasProd.Text = string.Format("{0:N0}", ctesMC > 0 ? (ctesMC * bisnietasP) : 0);
                picBoxB.Visible = double.Parse(txtBisnietasProd.Text) == limiteB && limiteB > 0 ? true : false;

                lblProgresoRed.Text = string.Format("Progreso de la Red {0:P2}", progresoRed / 100);
                pgrBarRed.Value = int.Parse(Math.Round(progresoRed).ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error en el detalle");
            }

            
        }
        /// <summary>
        /// Cierra la ventana actual.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormDetalle_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar == Convert.ToChar(Keys.Escape))
            {
                this.Close();
            }
        }
    }
}
