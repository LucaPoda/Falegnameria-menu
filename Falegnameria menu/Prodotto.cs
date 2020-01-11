using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Falegnameria_menu
{
    class Prodotto
    {
        private string nomeprodotto;
        private int numeropezzi;
        private double prezzo_listino;
        private double sconto;
        private double costo;
        private double ricarica;
        private double trasporto;
        private double accessori;
        private double posa;
        private double totale;

        private int unitatotale_trasporto;
        private int unitatotale_accessori;
        private int unitatotale_posa;

        public Prodotto()
        {
            Nomeprodotto = "";
            Numeropezzi = 1;
            Prezzo_listino = 0.0;
            Sconto = 0;
            Costo = 0.0;
            Ricarica = 0;
            Trasporto = 0.0;
            Accessori = 0.0;
            Posa = 0.0;
            Totale = 0;

            Unitatotale_trasporto = 1;
            Unitatotale_accessori = 1;
            Unitatotale_posa = 1;
        }

        public string Nomeprodotto
        {
            get
            {
                return nomeprodotto;
            }

            set
            {
                nomeprodotto = value;
            }
        }

        public int Numeropezzi
        {
            get
            {
                return numeropezzi;
            }

            set
            {
                numeropezzi = value;
            }
        }

        public double Prezzo_listino
        {
            get
            {
                return prezzo_listino;
            }

            set
            {
                prezzo_listino = value;
            }
        }

        public double Sconto
        {
            get
            {
                return sconto;
            }

            set
            {
                sconto = value;
            }
        }

        public double Costo
        {
            get
            {
                return costo;
            }

            set
            {
                costo = value;
            }
        }

        public double Ricarica
        {
            get
            {
                return ricarica;
            }

            set
            {
                ricarica = value;
            }
        }

        public double Trasporto
        {
            get
            {
                return trasporto;
            }

            set
            {
                trasporto = value;
            }
        }

        public double Accessori
        {
            get
            {
                return accessori;
            }

            set
            {
                accessori = value;
            }
        }

        public double Posa
        {
            get
            {
                return posa;
            }

            set
            {
                posa = value;
            }
        }

        public double Totale
        {
            get
            {
                return totale;
            }

            set
            {
                totale = value;
            }
        }

        public int Unitatotale_trasporto
        {
            get
            {
                return unitatotale_trasporto;
            }

            set
            {
                unitatotale_trasporto = value;
            }
        }

        public int Unitatotale_accessori
        {
            get
            {
                return unitatotale_accessori;
            }

            set
            {
                unitatotale_accessori = value;
            }
        }

        public int Unitatotale_posa
        {
            get
            {
                return unitatotale_posa;
            }

            set
            {
                unitatotale_posa = value;
            }
        }
    }
}
