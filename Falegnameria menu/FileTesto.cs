using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Collections;

namespace Falegnameria_menu
{
    class FileTesto
    {
        private string nomeFile;
        public FileTesto(string nomeFile, Cliente Cliente1)
        {
            this.nomeFile = nomeFile;
        }

        public string NomeFile
        {
            set
            {
                nomeFile = value;
            }
            get
            {
                return nomeFile;
            }
        }

        // Se il file non esiste lo crea. Se il file esiste lo sovrascrive.
        public void Reset()
        {
            StreamWriter file = new StreamWriter(nomeFile, false);
            file.Close();
        }

        // Se il file non esiste lo crea. Se il file esiste non fa nulla.
        public void Crea(string nomeFile)
        {
            if (!File.Exists(nomeFile))
            {
                StreamWriter file = new StreamWriter(nomeFile, false);
                file.Close();
            }
        }

        // Aggiunge una nuova riga alla fine del file.
        public void Scrivi(Cliente Cliente1)
        {
            /////////////trasforma i costi(trasporto,accessori,posa) da unitari a totale per il NB
            StreamWriter file = new StreamWriter(nomeFile, true); // append
            file.WriteLine(Cliente1.Nome);
            file.WriteLine(Cliente1.Cognome);
            file.WriteLine(Cliente1.Totalefattura);
            file.WriteLine(Cliente1.Nfattura);
            file.WriteLine(Cliente1.Totaleprodotti);//da implementare ogni volta
            for (int i = 0; i < Cliente1.Totaleprodotti; i++)
            {
                file.WriteLine(Cliente1.p[i].Nomeprodotto);
                file.WriteLine(Cliente1.p[i].Numeropezzi);
                file.WriteLine(Cliente1.p[i].Prezzo_listino);
                file.WriteLine(Cliente1.p[i].Sconto);
                file.WriteLine(Cliente1.p[i].Costo);
                file.WriteLine(Cliente1.p[i].Ricarica);
                file.WriteLine(Cliente1.p[i].Trasporto);
                file.WriteLine(Cliente1.p[i].Accessori);
                file.WriteLine(Cliente1.p[i].Posa);
                file.WriteLine(Cliente1.p[i].Totale);
            }
            file.Close();
        }

        // Legge il file inserendo le righe in un ArrayList.
        public Cliente Leggi(Cliente Cliente1)
        {
            StreamReader file = new StreamReader(nomeFile);

            while (!file.EndOfStream)
            {
                Cliente1.Nome = file.ReadLine();
                Cliente1.Cognome = file.ReadLine();
                Cliente1.Totalefattura = Convert.ToDouble(file.ReadLine());
                Cliente1.Nfattura = System.Convert.ToInt32(file.ReadLine());
                Cliente1.Totaleprodotti = System.Convert.ToInt32(file.ReadLine());
                for (int i = 0; i < Cliente1.Totaleprodotti; i++)
                {
                    Cliente1.p[i].Nomeprodotto = file.ReadLine();
                    Cliente1.p[i].Numeropezzi = System.Convert.ToInt32(file.ReadLine());
                    Cliente1.p[i].Prezzo_listino = Convert.ToDouble(file.ReadLine());
                    Cliente1.p[i].Sconto = Convert.ToDouble(file.ReadLine());
                    Cliente1.p[i].Costo = Convert.ToDouble(file.ReadLine());
                    Cliente1.p[i].Ricarica = Convert.ToDouble(file.ReadLine());
                    Cliente1.p[i].Trasporto = Convert.ToDouble(file.ReadLine());
                    Cliente1.p[i].Accessori = Convert.ToDouble(file.ReadLine());
                    Cliente1.p[i].Posa = Convert.ToDouble(file.ReadLine());
                    Cliente1.p[i].Totale = Convert.ToDouble(file.ReadLine());
                }
            }
            file.Close();
            return Cliente1;
        }
    }
}
