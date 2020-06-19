using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Falegnameria_menu
{
    [Serializable]
    public class Prodotto
    {
        public string NomeProdotto { get; set; }
        public int NumeroPezzi { get; set; }
        public double PrezzoListino { get; set; }
        public double Sconto { get; set; }
        public double Costo { get; set; }
        public double Ricarica { get; set; }
        public double Trasporto { get; set; }
        public double Accessori { get; set; }
        public double Posa { get; set; }
        public double Lavorazioni { get; set; }
        public double Totale { get; set; }
        public bool UnitaTrasporto { get; set; }
        public bool UnitaAccessori { get; set; }
        public bool UnitaPosa { get; set; }
        public bool UnitaLavorazioni { get; set; }

        public Prodotto()
        {
            NomeProdotto = "";
            NumeroPezzi = 1;
            PrezzoListino = 0;
            Sconto = 0;
            Costo = 0;
            Ricarica = 0;
            Trasporto = 0;
            Accessori = 0;
            Posa = 0;
            Lavorazioni = 0;
            Totale = 0;

            UnitaTrasporto = false;
            UnitaAccessori = false;
            UnitaPosa = false;
            UnitaLavorazioni = false;
        }

        public double UpdateCosto()
        {
            Costo = NumeroPezzi * (PrezzoListino * (100 - Sconto) / 100);
            return Costo;
        }

        public double UpdateTotale()
        {
            if (Ricarica > 0)
                Totale = Costo + (Costo) / 100 * Ricarica;
            else
                Totale = Costo;

            Totale += CalcolaCostiAggiuntivi();
            return Totale;
        }

        public double CalcolaCostiAggiuntivi()
        {
            double prezzo = 0;

            if (UnitaTrasporto)
                prezzo += Trasporto * NumeroPezzi;
            else
                prezzo += Trasporto;

            if (UnitaPosa)
                prezzo += Posa * NumeroPezzi;
            else
                prezzo += Posa;

            if (UnitaAccessori)
                prezzo += Accessori * NumeroPezzi;
            else
                prezzo += Accessori;

            if (UnitaLavorazioni)
                prezzo += Lavorazioni * NumeroPezzi;
            else
                prezzo += Lavorazioni;

            return prezzo;
        }
    }
}
