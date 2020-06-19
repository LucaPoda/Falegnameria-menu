using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Falegnameria_menu
{
    public partial class Form1 : Form
    {
        public bool ShowHeaderInFiles { get; set; }
        public Color Std { get; set; }
        public Cliente Cliente {get; set; }
        public string Nome { get; set; }

        public Form1()
        {
            InitializeComponent();
            /*index è il numero di fattura a cui si è arrivati per quel cliente*/
            
            
            //file_salvato = false;
            Cliente = new Cliente();
            timer1.Enabled = true;

            /*Imposto i bottoni dei costi totali come predefiniti, colorandone il contorno: */
            Std = btnTotaleTrasporto.FlatAppearance.BorderColor;
            /*Trasporto:*/
            btnTotaleTrasporto.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnTotaleTrasporto.FlatAppearance.BorderSize = 2;
            /*Accessori:*/
            btnTotaleAccessori.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnTotaleAccessori.FlatAppearance.BorderSize = 2;
            /*Posa:*/
            btnTotalePosa.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnTotalePosa.FlatAppearance.BorderSize = 2;
            /*Lavorazioni:*/
            btnTotaleLavorazioni.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnTotaleLavorazioni.FlatAppearance.BorderSize = 2;
        }

        private void SalvaCliente()
        { 
            Cliente.Nome = txtNome.Text;
            Cliente.Cognome = txtCognome.Text;

            Cliente.Numerofattura = int.Parse(txtNfattura.Text);
            Cliente.TotaleFattura = Convert.ToDouble(txtTotalefattura.Text);

            Cliente.ListaProdotti[Cliente.IndiceProdotto].NomeProdotto = txtProdotto.Text;
            Cliente.ListaProdotti[Cliente.IndiceProdotto].NumeroPezzi = System.Convert.ToInt32(txtNpezzi.Text);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].PrezzoListino = Convert.ToDouble(txtListino.Text);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].Sconto = Convert.ToDouble(txtSconto.Text);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].Costo = Convert.ToDouble(txtCostoParziale.Text);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].Ricarica = Convert.ToDouble(txtRicarica.Text);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].Trasporto = Convert.ToDouble(txtTrasporto.Text);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].Accessori = Convert.ToDouble(txtAccessori.Text);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].Posa = Convert.ToDouble(txtPosa.Text);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale = Convert.ToDouble(txtTotale.Text);
        }
        
        private bool IsDouble(string str)
        {
            foreach(var c in str)
            {
                if (!char.IsDigit(c) && c != ',')
                    return false;
            }
            return true;
        }
        
        private bool IsInteger(string str)
        {
            foreach (var c in str)
            {
                if (!char.IsDigit(c))
                    return false;
            }
            return true;
        }
        
        private void UpdateValues(object sender, EventArgs e)
        {
            Console.WriteLine(12.3);

            var s = (TextBox)sender;

            s.Text = s.Text.Replace('.', ',');

            if(s == txtNpezzi)
            {
                if (!IsInteger(s.Text))
                {
                    try
                    {
                        s.Text = ((int) double.Parse(s.Text)).ToString();
                    }
                    catch(Exception)
                    {
                        s.Text = "1";
                    }
                        
                }
            }                   
            if (!IsDouble(s.Text))
                s.Text = "0";

            double.TryParse(txtListino.Text, out double prezzoListino);
            if (prezzoListino != 0 && !string.IsNullOrWhiteSpace(txtListino.Text))
            {                
                Cliente.ListaProdotti[Cliente.IndiceProdotto].NumeroPezzi = int.Parse(txtNpezzi.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].PrezzoListino = double.Parse(txtListino.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Sconto = double.Parse(txtSconto.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Ricarica = double.Parse(txtRicarica.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Trasporto = double.Parse(txtTrasporto.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Accessori = double.Parse(txtAccessori.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Posa = double.Parse(txtPosa.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Lavorazioni = double.Parse(txtLavorazioni.Text);

                Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateCosto();
                txtCostoParziale.Text = Convert.ToString(Cliente.ListaProdotti[Cliente.IndiceProdotto].Costo);

                Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateTotale();
                txtTotale.Text = Convert.ToString(Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale);
                
                Cliente.UpdateTotaleFattura();                
                txtTotalefattura.Text = Convert.ToString(Cliente.TotaleFattura);
            }

            /*Se txtListino è nullo significa che devo utilizzare solo txtcosto (il prodotto non è stato acquistato da un rivenditore*/
            else if (prezzoListino != 0 && !string.IsNullOrWhiteSpace(txtListino.Text))
            {
                Cliente.ListaProdotti[Cliente.IndiceProdotto].NumeroPezzi = System.Convert.ToInt32(txtNpezzi.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Trasporto = Convert.ToDouble(txtTrasporto.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Accessori = Convert.ToDouble(txtAccessori.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Posa = Convert.ToDouble(txtPosa.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Lavorazioni = Convert.ToDouble(txtLavorazioni.Text);
                Cliente.ListaProdotti[Cliente.IndiceProdotto].Costo = Convert.ToDouble(txtCostoParziale.Text);

                Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateTotale();
                txtTotale.Text = Convert.ToString(Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale);
            }
            else
                txtTotale.Text = "0";
        }

        private void AggiornaBordi()
        {
            //
            // BtnTrasporto
            //
            if (Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaTrasporto)
            {
                btnUnitaTrasporto.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnUnitaTrasporto.FlatAppearance.BorderSize = 2;
                btnTotaleTrasporto.FlatAppearance.BorderColor = Std;
                btnTotaleTrasporto.FlatAppearance.BorderSize = 1;
            }
            else
            {
                btnTotaleTrasporto.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnTotaleTrasporto.FlatAppearance.BorderSize = 2;
                btnUnitaTrasporto.FlatAppearance.BorderColor = Std;
                btnUnitaTrasporto.FlatAppearance.BorderSize = 1;
            }
            //
            // BtnAccessori
            //
            if (Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaAccessori)
            {
                btnUnitaAccessori.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnUnitaAccessori.FlatAppearance.BorderSize = 2;
                btnTotaleAccessori.FlatAppearance.BorderColor = Std;
                btnTotaleAccessori.FlatAppearance.BorderSize = 1;
            }
            else
            {
                btnTotaleAccessori.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnTotaleAccessori.FlatAppearance.BorderSize = 2;
                btnUnitaAccessori.FlatAppearance.BorderColor = Std;
                btnUnitaAccessori.FlatAppearance.BorderSize = 1;
            }
            //
            // BtnPosa
            //
            if (Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaPosa)
            {
                btnUnitaPosa.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnUnitaPosa.FlatAppearance.BorderSize = 2;
                btnTotalePosa.FlatAppearance.BorderColor = Std;
                btnTotalePosa.FlatAppearance.BorderSize = 1;
            }
            else
            {
                btnTotalePosa.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnTotalePosa.FlatAppearance.BorderSize = 2;
                btnUnitaPosa.FlatAppearance.BorderColor = Std;
                btnUnitaPosa.FlatAppearance.BorderSize = 1;
            }
            //
            // BtnLavorazioni
            //
            if (Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaLavorazioni)
            {
                btnUnitaLavorazioni.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnUnitaLavorazioni.FlatAppearance.BorderSize = 2;
                btnTotaleLavorazioni.FlatAppearance.BorderColor = Std;
                btnTotaleLavorazioni.FlatAppearance.BorderSize = 1;
            }
            else
            {
                btnTotaleLavorazioni.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnTotaleLavorazioni.FlatAppearance.BorderSize = 2;
                btnUnitaLavorazioni.FlatAppearance.BorderColor = Std;
                btnUnitaLavorazioni.FlatAppearance.BorderSize = 1;
            }
        }
        //
        // Evento form_load
        private void Form1_Load(object sender, EventArgs e)
        {
            Cliente.ListaProdotti.Add(new Prodotto());
        }
        //
        //Eventi bottoni unità/totale:
        //
        private void btnUnitaTrasporto_Click(object sender, EventArgs e)
        {
            int a;
            int.TryParse(txtNpezzi.Text, out a);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaTrasporto = true;

            AggiornaBordi();
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateTotale();
            Cliente.UpdateTotaleFattura();

            txtTotale.Text = Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale.ToString();
            txtTotalefattura.Text = Cliente.TotaleFattura.ToString();
        }

        private void btnTotaleTrasporto_Click(object sender, EventArgs e)
        {
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaTrasporto = false;

            AggiornaBordi();
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateTotale();
            Cliente.UpdateTotaleFattura();

            txtTotale.Text = (Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale).ToString();
            txtTotalefattura.Text = (Cliente.TotaleFattura).ToString();
        }

        private void btnUnitaAccessori_Click(object sender, EventArgs e)
        {
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaAccessori = true;

            AggiornaBordi();
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateTotale();
            Cliente.UpdateTotaleFattura();

            txtTotale.Text = (Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale).ToString();
            txtTotalefattura.Text = (Cliente.TotaleFattura).ToString();
        }

        private void btnTotaleAccessori_Click(object sender, EventArgs e)
        {
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaAccessori = false;

            AggiornaBordi();
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateTotale();
            Cliente.UpdateTotaleFattura();

            txtTotale.Text = (Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale).ToString();
            txtTotalefattura.Text = (Cliente.TotaleFattura).ToString();
        }

        private void btnUnitaPosa_Click(object sender, EventArgs e)
        {
            int a;
            int.TryParse(txtNpezzi.Text, out a);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaPosa = true;

            AggiornaBordi();
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateTotale();
            Cliente.UpdateTotaleFattura();

            txtTotale.Text = (Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale).ToString();
            txtTotalefattura.Text = (Cliente.TotaleFattura).ToString();
        }

        private void btnTotalePosa_Click(object sender, EventArgs e)
        {
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaPosa = false;

            AggiornaBordi();
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateTotale();
            Cliente.UpdateTotaleFattura();

            txtTotale.Text = (Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale).ToString();
            txtTotalefattura.Text = (Cliente.TotaleFattura).ToString();
        }

        private void btnUnitaLavorazioni_Click(object sender, EventArgs e)
        {
            int a;
            int.TryParse(txtNpezzi.Text, out a);
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaLavorazioni = true;

            AggiornaBordi();
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateTotale();
            Cliente.UpdateTotaleFattura();

            txtTotale.Text = (Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale).ToString();
            txtTotalefattura.Text = (Cliente.TotaleFattura).ToString();
        }

        private void btnTotaleLavorazioni_Click(object sender, EventArgs e)
        {
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UnitaLavorazioni = false;

            btnTotaleLavorazioni.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnTotaleLavorazioni.FlatAppearance.BorderSize = 2;
            btnUnitaLavorazioni.FlatAppearance.BorderColor = Std;
            btnUnitaLavorazioni.FlatAppearance.BorderSize = 1;
            Cliente.ListaProdotti[Cliente.IndiceProdotto].UpdateTotale();
            Cliente.UpdateTotaleFattura();

            txtTotale.Text = (Cliente.ListaProdotti[Cliente.IndiceProdotto].Totale).ToString();
            txtTotalefattura.Text = (Cliente.TotaleFattura).ToString();
        }

        // Eventi bottoni in basso

        private void btnNuovoprodotto_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtNome.Text) || string.IsNullOrWhiteSpace(txtCognome.Text))
            {
                MessageBox.Show("Inserire nome e cognome!");
                return;
            }

            SalvaCliente();

            /*salvo il prodotto corrente nella combobox*/
            cmbNumeroprodotto.Items.RemoveAt(Cliente.IndiceProdotto);
            cmbNumeroprodotto.Items.Insert(Cliente.IndiceProdotto, Cliente.ListaProdotti[Cliente.IndiceProdotto].NomeProdotto);

            // aggiungo un'elemento vuoto alla combobox

            Cliente.ListaProdotti.Add(new Prodotto());
            if (!string.IsNullOrWhiteSpace(cmbNumeroprodotto.Items[cmbNumeroprodotto.Items.Count - 1].ToString()))
                cmbNumeroprodotto.Items.Add("");

            Cliente.IndiceProdotto = Cliente.ListaProdotti.Count - 1;
            cmbNumeroprodotto.SelectedIndex = Cliente.IndiceProdotto;
        }

        private void btnSovrascrivi_Click(object sender, EventArgs e)
        {
            if (File.Exists(Nome))
            {
                File.Delete(Nome);
                //scrittura();
            }
            else
            {
                btnSalva_Click(sender, e);
            }
        }

        private void btnSalva_Click(object sender, EventArgs e)
        {
            WordFile.scrittura(ShowHeaderInFiles, Cliente);
        }

        private void btnStampa_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txtNome.Text) && !string.IsNullOrWhiteSpace(txtCognome.Text))
            {
                SalvaCliente();
                WordFile.scrittura(ShowHeaderInFiles, Cliente);
                DataFile<Cliente> f = new DataFile<Cliente>("ArchivioFatture\\" + Nome + ".tnl");
                f.Reset();
                f.Scrvi(Cliente);
            }

            else
                MessageBox.Show("Inserire il nome ed il cognome del cliente!", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        // Eventi su menu strip

        private void aiutoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Chiama Marco!", "Sembra tu abbia una grave disabilità!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        private void nuovoProdottoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnNuovoprodotto_Click(sender,e);
        }
        
        private void nuovoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form1 Form = new Form1();
            Form.Show();
        }
        
        private void salvaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnSovrascrivi_Click(sender, e);
        }
        
        private void salvaConNomeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnSalva_Click(sender, e);
        }
        
        private void stampaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnStampa_Click(sender, e);
        }
        
        private void apriToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        // txt deselect

        private void txtNome_Deselect(object sender, EventArgs e)
        {
            Cliente.Nome = txtNome.Text;
            int index = 0;
            do
            {
                index++;
                Nome = txtNome.Text + txtCognome.Text + index;
            } while (File.Exists(Nome));
            txtNfattura.Text = index.ToString();
            Cliente.Indice = index;
        }

        private void txtCognome_Deselect(object sender, EventArgs e)
        {
            Cliente.Cognome = txtCognome.Text;
            int index = 0;
            do
            {
                index++;
                Nome = txtNome.Text + txtCognome.Text + index;
            } while (File.Exists(Nome));
            txtNfattura.Text = index.ToString();
            Cliente.Indice = index;
        }
        
        private void txtNomeProdotto_Deselect(object sender, EventArgs e)
        {
            Cliente.ListaProdotti[Cliente.IndiceProdotto].NomeProdotto = txtProdotto.Text;
            cmbNumeroprodotto.Items.RemoveAt(Cliente.IndiceProdotto);
            cmbNumeroprodotto.Items.Insert(Cliente.IndiceProdotto, txtProdotto.Text);
            cmbNumeroprodotto.SelectedIndex = Cliente.IndiceProdotto;
        }
        //
        // altro
        //
        private void cbxIntestazione_CheckedChanged(object sender, EventArgs e)
        {
            ShowHeaderInFiles = cbxIntestazione.Checked;
        }

        private void cmbNumeroprodotto_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cliente.IndiceProdotto = cmbNumeroprodotto.SelectedIndex;
            Prodotto p = Cliente.ListaProdotti[Cliente.IndiceProdotto];
            txtProdotto.Text = p.NomeProdotto;
            txtNpezzi.Text = p.NumeroPezzi.ToString();
            txtListino.Text = p.PrezzoListino.ToString();
            txtSconto.Text = p.Sconto.ToString();
            txtCostoParziale.Text = p.Costo.ToString();
            txtRicarica.Text = p.Ricarica.ToString();
            txtTrasporto.Text = p.Trasporto.ToString();
            txtAccessori.Text = p.Accessori.ToString();
            txtPosa.Text = p.Posa.ToString();
            txtLavorazioni.Text = p.Lavorazioni.ToString();
            txtTotale.Text = p.Totale.ToString();

            AggiornaBordi();

            if (cmbNumeroprodotto.SelectedIndex == cmbNumeroprodotto.Items.Count - 1)
                btnNuovoprodotto.Text = "Nuovo Prodotto";
            else
                btnNuovoprodotto.Text = "Aggiorna " + cmbNumeroprodotto.SelectedItem.ToString();
        }

        /*private void Form1_Closing(object sender, FormClosedEventArgs e)
        {
            MessageBoxButtons btns = MessageBoxButtons.YesNoCancel;
            DialogResult result = MessageBox.Show("Stai uscendo senza salvare. \nVuoi salvare le modifiche al file ...?", "ërror", btns);
            if (result == DialogResult.Yes)

            if (result == DialogResult.No)
                Close();
            //if (result == DialogResult.Cancel)


        }*/
    }
}
