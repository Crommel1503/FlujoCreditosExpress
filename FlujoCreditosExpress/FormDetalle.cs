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

            txtCarteraTotal.Text = string.Format("{0:N0}", carteraT);

            txtDistribuidorasDet.Text = string.Format("{0:N0}", distribuidoras);
            txtCtesDsitribuidorasDet.Text = string.Format("{0:N0}", ctesDist);
            txtCtesMediosMasivosDet.Text = string.Format("{0:N0}", ctesMM);
            txtCtesZafyDet.Text = string.Format("{0:N0}", ctesCZ);

            txtLideresIniciadorasDet.Text = string.Format("{0:N0}", Properties.Settings.Default.CantIniciadoras);
            txtLideresDet.Text = string.Format("{0:N0}", Properties.Settings.Default.CantLideresH);
            txtMiembrosCelulaDet.Text = string.Format("{0:N0}", ctesMC);

            lblPCT.Text = string.Format("{0:P2}", carteraT > 0 ? (carteraT / carteraT) : 0);
            lblPD.Text = string.Format("{0:P2}", distribuidoras > 0 ? (distribuidoras / carteraT) : 0);
            lblPCD.Text = string.Format("{0:P2}", ctesDist > 0 ? (ctesDist / carteraT) : 0);
            lblPMM.Text = string.Format("{0:P2}", ctesMM > 0 ? (ctesMM / carteraT) : 0);
            lblPCZ.Text = string.Format("{0:P2}", ctesCZ > 0 ? (ctesCZ / carteraT) : 0);
            lblPMC.Text = string.Format("{0:P2}", ctesMC > 0 ? (ctesMC / carteraT) : 0);

            txtMadresProd.Text = string.Format("{0:N0}", Properties.Settings.Default.CantIniciadoras);
            txtHijasProd.Text = string.Format("{0:N0}", ctesMC > 0 ? (ctesMC * hijasP) : 0);
            txtNietasProd.Text = string.Format("{0:N0}", ctesMC > 0 ? (ctesMC * nietasP) : 0);
            txtBisnietasProd.Text = string.Format("{0:N0}", ctesMC > 0 ? (ctesMC * bisnietasP) : 0);
        }
    }
}
