using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Falegnameria_menu
{
    public partial class Form1 : Form
    {
        Color std;
        private Cliente Cliente1;
        private int index;
        private bool file_salvato;
        /*Costruttore vuoto*/
        public Form1()
        {
            InitializeComponent();
            /*index è il numero di fattura a cui si è arrivati per quel cliente*/
            index = 1;
            file_salvato = false;
            Cliente1 = new Cliente();
            timer1.Enabled = true;

            /*Imposto i bottoni dei costi totali come predefiniti, colorandone il contorno: */
            std = btnTotaleTrasporto.FlatAppearance.BorderColor;
                /*Trasporto:*/
                btnTotaleTrasporto.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnTotaleTrasporto.FlatAppearance.BorderSize = 2;
                /*Accessori:*/
                btnTotaleAccessori.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnTotaleAccessori.FlatAppearance.BorderSize = 2;
                /*Posa:*/
                btnTotalePosa.FlatAppearance.BorderColor = Color.RoyalBlue;
                btnTotalePosa.FlatAppearance.BorderSize = 2;
        }

        /*Aggiorno la form principale ogni 100ms...*/
        private void timer1_Tick(object sender, EventArgs e)
        {
            char[] numeri = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            int result;

            /*Evito bug sostituendo punti con le virgole nelle txtbox*/
            /*Se la txtbox non contiene un numero inserisco 0,0 :D*/
            txtListino.Text = txtListino.Text.Replace(".", ",");
            result = txtListino.Text.IndexOfAny(numeri);
            if (result == -1)
                txtListino.Text = "0,0";

            txtCostoParziale.Text = txtCostoParziale.Text.Replace(".", ",");
            result = txtCostoParziale.Text.IndexOfAny(numeri);
            if (result == -1)
                txtCostoParziale.Text = "0,0";

            txtSconto.Text = txtSconto.Text.Replace(".", ",");
            result = txtSconto.Text.IndexOfAny(numeri);
            if (result == -1)
                txtSconto.Text = "0,0";

            txtRicarica.Text = txtRicarica.Text.Replace(".", ",");
            result = txtRicarica.Text.IndexOfAny(numeri);
            if (result == -1)
                txtRicarica.Text = "0,0";

            txtTrasporto.Text = txtTrasporto.Text.Replace(".", ",");
            result = txtTrasporto.Text.IndexOfAny(numeri);
            if (result == -1)
                txtTrasporto.Text = "0,0";

            txtAccessori.Text = txtAccessori.Text.Replace(".", ",");
            result = txtAccessori.Text.IndexOfAny(numeri);
            if (result == -1)
                txtAccessori.Text = "0,0";

            txtPosa.Text = txtPosa.Text.Replace(".", ",");
            result = txtPosa.Text.IndexOfAny(numeri);
            if (result == -1)
                txtPosa.Text = "0,0";

            /*Se txtListino non è nullo*/
            if (txtListino.Text != "0,0" && txtListino.Text != "" && txtListino.Text != "0" && txtListino.Text != " ")
            {
                int a;
                int.TryParse(txtNpezzi.Text, out a);
                Cliente1.p[Cliente1.Indiceprod].Numeropezzi = a;
                Cliente1.p[Cliente1.Indiceprod].Prezzo_listino = Convert.ToDouble(txtListino.Text);
                Cliente1.p[Cliente1.Indiceprod].Sconto = Convert.ToDouble(txtSconto.Text);
                Cliente1.p[Cliente1.Indiceprod].Ricarica = Convert.ToDouble(txtRicarica.Text);
                Cliente1.p[Cliente1.Indiceprod].Trasporto = Convert.ToDouble(txtTrasporto.Text);
                Cliente1.p[Cliente1.Indiceprod].Accessori = Convert.ToDouble(txtAccessori.Text);
                Cliente1.p[Cliente1.Indiceprod].Posa = Convert.ToDouble(txtPosa.Text);

                Cliente1.p[Cliente1.Indiceprod].Costo = Cliente1.p[Cliente1.Indiceprod].Numeropezzi * (Cliente1.p[Cliente1.Indiceprod].Prezzo_listino * (100 - Cliente1.p[Cliente1.Indiceprod].Sconto) / 100);
                txtCostoParziale.Text = Convert.ToString(Cliente1.p[Cliente1.Indiceprod].Costo);
                if (Cliente1.p[Cliente1.Indiceprod].Ricarica > 0)
                    Cliente1.p[Cliente1.Indiceprod].Totale = Cliente1.p[Cliente1.Indiceprod].Costo + (Cliente1.p[Cliente1.Indiceprod].Costo) / 100 * Cliente1.p[Cliente1.Indiceprod].Ricarica + costi_aggiuntivi()/*trasporto + accessori + posa*/;
                else
                    Cliente1.p[Cliente1.Indiceprod].Totale = Cliente1.p[Cliente1.Indiceprod].Costo + Cliente1.p[Cliente1.Indiceprod].Trasporto + Cliente1.p[Cliente1.Indiceprod].Accessori + Cliente1.p[Cliente1.Indiceprod].Posa;

                txtTotale.Text = Convert.ToString(Cliente1.p[Cliente1.Indiceprod].Totale);
            }

            /*Se txtListino è nullo significa che devo utilizzare solo txtcosto (il prodotto non è stato acquistato da un rivenditore*/
            else if (txtCostoParziale.Text != "0,0" && txtCostoParziale.Text != "" && txtCostoParziale.Text != "0" && txtCostoParziale.Text != " ")
            {
                Cliente1.p[Cliente1.Indiceprod].Numeropezzi = System.Convert.ToInt32(txtNpezzi.Text);
                Cliente1.p[Cliente1.Indiceprod].Trasporto = Convert.ToDouble(txtTrasporto.Text);
                Cliente1.p[Cliente1.Indiceprod].Accessori = Convert.ToDouble(txtAccessori.Text);
                Cliente1.p[Cliente1.Indiceprod].Posa = Convert.ToDouble(txtPosa.Text);
                Cliente1.p[Cliente1.Indiceprod].Costo = Convert.ToDouble(txtCostoParziale.Text);
                //sì lo so quella sotto è una stringa del forza 4
                //MessageBox.Show(" hai vinto!", "Vittoria del giocatore 1!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Cliente1.p[Cliente1.Indiceprod].Totale = Cliente1.p[Cliente1.Indiceprod].Numeropezzi * Cliente1.p[Cliente1.Indiceprod].Costo + costi_aggiuntivi();
                txtTotale.Text = Convert.ToString(Cliente1.p[Cliente1.Indiceprod].Totale);
            }
        }

        /*Calcola costi per unità o totale per Trasporto, Accessori e Posa*/
        public double costi_aggiuntivi()
        {
            /*moltiplico il costo per il valore della var. Unitatotale... che può essere uguale ad 1 (totale) o uguale ad Npezzi*/
            return (Cliente1.p[Cliente1.Indiceprod].Trasporto * Cliente1.p[Cliente1.Indiceprod].Unitatotale_trasporto) +
                (Cliente1.p[Cliente1.Indiceprod].Accessori * Cliente1.p[Cliente1.Indiceprod].Unitatotale_accessori) +
                (Cliente1.p[Cliente1.Indiceprod].Posa * Cliente1.p[Cliente1.Indiceprod].Unitatotale_posa);
        }






        /*Eventi bottoni unità trasporto:*/
        /*Coloro di Blu il contorno del bottone selezionato e setto standard l'altro*/
        private void btnUnitaTrasporto_Click(object sender, EventArgs e)
        {
            Cliente1.p[Cliente1.Indiceprod].Unitatotale_trasporto = Cliente1.p[Cliente1.Indiceprod].Numeropezzi;

            btnUnitaTrasporto.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnUnitaTrasporto.FlatAppearance.BorderSize = 2;
            btnTotaleTrasporto.FlatAppearance.BorderColor = std;
            btnTotaleTrasporto.FlatAppearance.BorderSize = 1;
        }

        private void btnTotaleTrasporto_Click(object sender, EventArgs e)
        {
            Cliente1.p[Cliente1.Indiceprod].Unitatotale_trasporto = 1;

            btnTotaleTrasporto.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnTotaleTrasporto.FlatAppearance.BorderSize = 2;
            btnUnitaTrasporto.FlatAppearance.BorderColor = std;
            btnUnitaTrasporto.FlatAppearance.BorderSize = 1;
        }

        private void btnUnitaAccessori_Click(object sender, EventArgs e)
        {
            Cliente1.p[Cliente1.Indiceprod].Unitatotale_accessori = Cliente1.p[Cliente1.Indiceprod].Numeropezzi;

            btnUnitaAccessori.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnUnitaAccessori.FlatAppearance.BorderSize = 2;
            btnTotaleAccessori.FlatAppearance.BorderColor = std;
            btnTotaleAccessori.FlatAppearance.BorderSize = 1;
        }

        private void btnTotaleAccessori_Click(object sender, EventArgs e)
        {
            Cliente1.p[Cliente1.Indiceprod].Unitatotale_accessori = 1;

            btnTotaleAccessori.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnTotaleAccessori.FlatAppearance.BorderSize = 2;
            btnUnitaAccessori.FlatAppearance.BorderColor = std;
            btnUnitaAccessori.FlatAppearance.BorderSize = 1;
        }

        private void btnUnitaPosa_Click(object sender, EventArgs e)
        {
            Cliente1.p[Cliente1.Indiceprod].Unitatotale_posa = Cliente1.p[Cliente1.Indiceprod].Numeropezzi;

            btnUnitaPosa.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnUnitaPosa.FlatAppearance.BorderSize = 2;
            btnTotalePosa.FlatAppearance.BorderColor = std;
            btnTotalePosa.FlatAppearance.BorderSize = 1;
        }

        private void btnTotalePosa_Click(object sender, EventArgs e)
        {
            Cliente1.p[Cliente1.Indiceprod].Unitatotale_posa = 1;

            btnTotalePosa.FlatAppearance.BorderColor = Color.RoyalBlue;
            btnTotalePosa.FlatAppearance.BorderSize = 2;
            btnUnitaPosa.FlatAppearance.BorderColor = std;
            btnUnitaPosa.FlatAppearance.BorderSize = 1;
        }



        /*Click su menu strip*/
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
        /*Click su bottoni*/
        private void btnNuovoprodotto_Click(object sender, EventArgs e)
        {
            /*Salvo i dati dell'oggetto cliente:*/
            Cliente1.Totaleprodotti++;
            Cliente1.Nome = txtNome.Text;
            Cliente1.Cognome = txtCognome.Text;
            Cliente1.Nfattura = int.Parse(txtNfattura.Text);
            Cliente1.Totalefattura = Convert.ToDouble(txtTotalefattura.Text);
            Cliente1.p[Cliente1.Indiceprod].Nomeprodotto = txtProdotto.Text;
            Cliente1.p[Cliente1.Indiceprod].Numeropezzi = System.Convert.ToInt32(txtNpezzi.Text);
            Cliente1.p[Cliente1.Indiceprod].Prezzo_listino = Convert.ToDouble(txtListino.Text);
            Cliente1.p[Cliente1.Indiceprod].Sconto = Convert.ToDouble(txtSconto.Text);
            Cliente1.p[Cliente1.Indiceprod].Costo = Convert.ToDouble(txtCostoParziale.Text);
            Cliente1.p[Cliente1.Indiceprod].Ricarica = Convert.ToDouble(txtRicarica.Text);
            Cliente1.p[Cliente1.Indiceprod].Trasporto = Convert.ToDouble(txtTrasporto.Text);
            Cliente1.p[Cliente1.Indiceprod].Accessori = Convert.ToDouble(txtAccessori.Text);
            Cliente1.p[Cliente1.Indiceprod].Posa = Convert.ToDouble(txtPosa.Text);
            Cliente1.p[Cliente1.Indiceprod].Totale = Convert.ToDouble(txtTotale.Text);

            /*aggiungere alla combo box*/
            cmbNumeroprodotto.Items.Insert(Cliente1.Indiceprod, Cliente1.p[Cliente1.Indiceprod].Nomeprodotto);
            Cliente1.Indiceprod ++;



        }

        private void btnSovrascrivi_Click(object sender, EventArgs e)
        {

        }

        private void btnSalva_Click(object sender, EventArgs e)
        {

        }

        private void btnStampa_Click(object sender, EventArgs e)
        { /*
            //creo l'applicazione
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = true;
            wordApp.WindowState = Word.WdWindowState.wdWindowStateNormal;
            

            //creo il documento
            Word.Document wordDoc = wordApp.Documents.Add();
            Word.Range docRange = wordDoc.Range();
            
            string imagePath = @"C:\Users\Marco\Desktop\definitivo\image.jpg";

            // Create an InlineShape in the InlineShapes collection where the picture should be added later
            // It is used to get automatically scaled sizes.
            Word.InlineShape autoScaledInlineShape = docRange.InlineShapes.AddPicture(imagePath);
            float scaledWidth = autoScaledInlineShape.Width;
            float scaledHeight = autoScaledInlineShape.Height;
            autoScaledInlineShape.Delete();

            // Create a new Shape and fill it with the picture
            Word.Shape newShape = wordDoc.Shapes.AddShape(1, 0, 0, scaledWidth, scaledHeight);
            newShape.Fill.UserPicture(imagePath);

            // Convert the Shape to an InlineShape and optional disable Border
            Word.InlineShape finalInlineShape = newShape.ConvertToInlineShape();
            finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            // Cut the range of the InlineShape to clipboard
            finalInlineShape.Range.Cut();

            // And paste it to the target Range
            docRange.Paste();




 
            //aggiungo un paragrafo
            Word.Paragraph objPara;
            objPara = wordDoc.Paragraphs.Add();

            objPara.Range.Text = "\r\nSOOOOOOOOOOOOOOOOS qui\n\n";
            */

            private void ImageToDocx(List<string> Images)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document wordDoc = wordApp.Documents.Add();
            Range docRange = wordDoc.Range();

            float mHeight = 0;
            for (int i = 0; i <= Images.Count - 1; i++)
            {
                // Create an InlineShape in the InlineShapes collection where the picture should be added later
                // It is used to get automatically scaled sizes.
                InlineShape autoScaledInlineShape = docRange.InlineShapes.AddPicture(Images[i]);
                float scaledWidth = autoScaledInlineShape.Width;
                float scaledHeight = autoScaledInlineShape.Height;
                mHeight += scaledHeight;
                autoScaledInlineShape.Delete();

                // Create a new Shape and fill it with the picture
                Shape newShape = wordDoc.Shapes.AddShape(1, 0, 0, scaledWidth, mHeight);
                newShape.Fill.UserPicture(Images[i]);

                // Convert the Shape to an InlineShape and optional disable Border
                InlineShape finalInlineShape = newShape.ConvertToInlineShape();
                finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

                // Cut the range of the InlineShape to clipboard
                finalInlineShape.Range.Cut();

                // And paste it to the target Range
                docRange.Paste();
            }



            wordDoc.SaveAs("C:\\Users\\Marco\\Desktop\\definitivo\\Falegnameria menu\\Falegnameria menu\\ArchivioWord\\esempio.docx");
            wordDoc.Close();
            wordApp.Quit();
        }











        /*if(logoToolStripMenuItem.Checked==true && contattiToolStripMenuItem.Checked == true)*/
    }
}
