using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Falegnameria_menu
{
    [Serializable]
    public class Cliente
    {
        public string Nome { get; set; }
        public string Cognome { get; set; }
        public int Numerofattura { get; set; }
        public int IndiceProdotto { get; set; }
        public List<Prodotto> ListaProdotti { get; set; }
        public double TotaleFattura { get; set; }
        public int Indice { get; set; }

        public Cliente()
        {
            ListaProdotti = new List<Prodotto>();
            Nome = " ";
            Cognome = " ";
            IndiceProdotto = 0;
            Numerofattura = 1;
        }

        public void UpdateTotaleFattura()
        {
            double n = 0;
            foreach (var c in ListaProdotti)
            {
                n += c.Totale;
            }
            TotaleFattura = n;
        }
    }
}
