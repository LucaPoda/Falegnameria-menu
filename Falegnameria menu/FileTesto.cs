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
            file.WriteLine(Cliente1.TotaleFattura);
            file.WriteLine(Cliente1.Numerofattura);
            file.WriteLine(Cliente1.TotaleProdotti);//da implementare ogni volta
            for (int i = 0; i < Cliente1.TotaleProdotti; i++)
            {
                file.WriteLine(Cliente1.ListaProdotti[i].NomeProdotto);
                file.WriteLine(Cliente1.ListaProdotti[i].NumeroPezzi);
                file.WriteLine(Cliente1.ListaProdotti[i].PrezzoListino);
                file.WriteLine(Cliente1.ListaProdotti[i].Sconto);
                file.WriteLine(Cliente1.ListaProdotti[i].Costo);
                file.WriteLine(Cliente1.ListaProdotti[i].Ricarica);
                file.WriteLine(Cliente1.ListaProdotti[i].Trasporto);
                file.WriteLine(Cliente1.ListaProdotti[i].Accessori);
                file.WriteLine(Cliente1.ListaProdotti[i].Posa);
                file.WriteLine(Cliente1.ListaProdotti[i].Totale);
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
                Cliente1.TotaleFattura = Convert.ToDouble(file.ReadLine());
                Cliente1.Numerofattura = System.Convert.ToInt32(file.ReadLine());
                Cliente1.TotaleProdotti = System.Convert.ToInt32(file.ReadLine());
                for (int i = 0; i < Cliente1.TotaleProdotti; i++)
                {
                    Cliente1.ListaProdotti[i].NomeProdotto = file.ReadLine();
                    Cliente1.ListaProdotti[i].NumeroPezzi = System.Convert.ToInt32(file.ReadLine());
                    Cliente1.ListaProdotti[i].PrezzoListino = Convert.ToDouble(file.ReadLine());
                    Cliente1.ListaProdotti[i].Sconto = Convert.ToDouble(file.ReadLine());
                    Cliente1.ListaProdotti[i].Costo = Convert.ToDouble(file.ReadLine());
                    Cliente1.ListaProdotti[i].Ricarica = Convert.ToDouble(file.ReadLine());
                    Cliente1.ListaProdotti[i].Trasporto = Convert.ToDouble(file.ReadLine());
                    Cliente1.ListaProdotti[i].Accessori = Convert.ToDouble(file.ReadLine());
                    Cliente1.ListaProdotti[i].Posa = Convert.ToDouble(file.ReadLine());
                    Cliente1.ListaProdotti[i].Totale = Convert.ToDouble(file.ReadLine());
                }
            }
            file.Close();
            return Cliente1;
        }
    }
}
