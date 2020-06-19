using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace Falegnameria_menu
{
    [Serializable]
    class DataFile<T>
    {
        private string Name { get; set; }

        public DataFile(string nome)
        {
            this.Name = nome;
        }
        //
        // se il file non esiste lo crea.
        // se il file esiste lo sovrascrive.
        //
        public void Reset()
        {
            FileStream file = new FileStream(Name, FileMode.Create);
            file.Close();
        }
        //
        // se il file non esiste lo crea.
        // se il file eesiste non fa nulla.
        //
        public void Crea()
        {
            if (!File.Exists(Name))
            {
                Reset();
            }
        }
        //
        // Aggiunge un nuovo conatto alla fine del file.
        //
        public void Scrvi(T c)
        {
            FileStream file = new FileStream(Name, FileMode.Append);
            BinaryFormatter bf = new BinaryFormatter();
            bf.Serialize(file, c);
            file.Close();
        }

        public List<T> Leggi()
        {
            FileStream file = new FileStream(Name, FileMode.Open);
            BinaryFormatter bf = new BinaryFormatter();
            List<T> list = new List<T>();
            T c;
            while (file.Position != file.Length)
            {
                c = (T)bf.Deserialize(file);
                list.Add(c);
            }
            file.Close();
            return list;
        }
    }
}
