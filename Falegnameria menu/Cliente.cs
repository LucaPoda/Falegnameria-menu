using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Falegnameria_menu
{
    class Cliente
    {
        private string nome;
        private string cognome;
        private int nfattura;
        private int indiceprod;
        private int totaleprodotti;
        public Prodotto[] p = new Prodotto[500];
        private double totalefattura;

        public Cliente()
        {
            this.Nome = " ";
            this.Cognome = " ";
            this.Indiceprod = 0;
            this.Nfattura = 1;
            this.Totaleprodotti = 1;
            //cliente1.p[1].nome 
            for (int i = 0; i < 500; i++)
            {
                p[i] = new Prodotto();
            }
        }

        public string Nome
        {
            get
            {
                return nome;
            }

            set
            {
                nome = value;
            }
        }

        public string Cognome
        {
            get
            {
                return cognome;
            }

            set
            {
                cognome = value;
            }
        }

        public int Indiceprod
        {
            get
            {
                return indiceprod;
            }
            set
            {
                indiceprod = value;
            }
        }

        public double Totalefattura
        {
            get
            {
                return totalefattura;
            }

            set
            {
                totalefattura = value;
            }
        }

        public int Totaleprodotti
        {
            get
            {
                return totaleprodotti;
            }

            set
            {
                totaleprodotti = value;
            }
        }

        public int Nfattura
        {
            get
            {
                return nfattura;
            }

            set
            {
                nfattura = value;
            }
        }
    }
}
