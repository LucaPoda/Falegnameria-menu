using System;
using System.Collections;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace Falegnameria_menu
{
    class WordFile
    {
        /*Avendo già memmorizzato i dati ( Cliente1.Nome = txtNome.Text; e Cliente1.Cognome = txtCognome.Text;)*/
        public static void scrittura(bool intestazione, Cliente cliente)
        {
            object missing = System.Reflection.Missing.Value;

            Word.Application wordApp = new Word.Application();
            wordApp.Visible = false;
            wordApp.WindowState = Word.WdWindowState.wdWindowStateNormal;

            Word.Document wordDoc = wordApp.Documents.Add();

            var path = @"" + Path.GetFullPath("image.jpg");

            Word.Range range;

            //se la CheckBox dell'intestazione è checkata
            if (intestazione)
            {
                //incollo l'immagine come intestazione
                foreach (Microsoft.Office.Interop.Word.Section section in wordDoc.Sections)
                {
                    //Get the header range and add the header details.
                    Microsoft.Office.Interop.Word.Range docRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    docRange.Fields.Add(docRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);

                    // Create an InlineShape in the InlineShapes collection where the picture should be added later
                    // It is used to get automatically scaled sizes.
                    Word.InlineShape autoScaledInlineShape = docRange.InlineShapes.AddPicture(path);
                    float scaledWidth = autoScaledInlineShape.Width;
                    float scaledHeight = autoScaledInlineShape.Height;
                    autoScaledInlineShape.Delete();

                    // Create a new Shape and fill it with the picture
                    Word.Shape newShape = wordDoc.Shapes.AddShape(1, 0, 0, scaledWidth, scaledHeight);
                    newShape.Fill.UserPicture(path);

                    // Convert the Shape to an InlineShape and optional disable Border
                    Word.InlineShape finalInlineShape = newShape.ConvertToInlineShape();
                    finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

                    // Cut the range of the InlineShape to clipboard
                    finalInlineShape.Range.Cut();

                    // And paste it to the target Range
                    docRange.Paste();
                }
            }

            //formattazione testo (se la modifichi occhio alle tabulazioni sotto)
            wordDoc.Content.SetRange(0, 0);
            wordDoc.Content.Font.Size = 15;

            // Get the current date
            DateTime thisDay = DateTime.Today;

            // intestazione con nome cognome e data:
            wordDoc.Content.Text = "Preventivo\t\t\t\t\t\t\t\t\t\t" + thisDay.ToString("d") + "\nRiferimento: " + cliente.Nome + " " + cliente.Cognome + "\n" + Environment.NewLine;

            int i = 0;
            ArrayList totali_singoli_prodotti = new ArrayList(); //creo un array formato dal totale dei prodotti per riempire una tab alla fine

            //per ogni prodotto presente nell'array:
            foreach (Prodotto prodotto in cliente.ListaProdotti)
            {
                if (!string.IsNullOrWhiteSpace(cliente.ListaProdotti[i].NomeProdotto))
                {
                    range = GetRange(wordDoc);

                    //stampo la tabella con le caratteristiche del singolo prodotto
                    Word.Table tab = wordDoc.Tables.Add(range, 9, 3, ref missing, ref missing);
                    tab.Borders.Enable = 1;

                    // tabella prodotto n
                    foreach (Word.Row row in tab.Rows)
                    {
                        foreach (Word.Cell cell in row.Cells)
                        {
                            cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            //la colonna con la descrizione del prodotto la metto con sfondo grigio
                            if (cell.ColumnIndex == 1 || cell.RowIndex == 1)
                            {
                                cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                            }

                            //prima riga
                            if (cell.RowIndex == 1)
                            {
                                if (cell.ColumnIndex == 2)
                                    cell.Range.Text = "Sconto :";
                                if (cell.ColumnIndex == 3)
                                    cell.Range.Text = "Totale :";
                            }

                            //scrivo le descrizioni nella colonna a sinistra
                            if (cell.ColumnIndex == 1)
                            {
                                if (cell.RowIndex == 2)
                                {
                                    if (cliente.ListaProdotti[i].NumeroPezzi != 1)
                                        cell.Range.Text = cliente.ListaProdotti[i].NumeroPezzi + " x " + cliente.ListaProdotti[i].NomeProdotto + ":";
                                    else
                                    {
                                        cell.Range.Text = cliente.ListaProdotti[i].NomeProdotto + ":";
                                    }
                                }
                                if (cell.RowIndex == 3)
                                {
                                    cell.Range.Text = "Sconto :";
                                }
                                if (cell.RowIndex == 4)
                                {
                                    cell.Range.Text = "Costo :";
                                }
                                if (cell.RowIndex == 5)
                                {
                                    cell.Range.Text = "Ricarica :";
                                }
                                if (cell.RowIndex == 6)
                                {
                                    cell.Range.Text = "Posa :";
                                }
                                if (cell.RowIndex == 7)
                                {
                                    cell.Range.Text = "Trasporti :";
                                }
                                if (cell.RowIndex == 8)
                                {
                                    cell.Range.Text = "Accessori :";
                                }
                                if (cell.RowIndex == 9)
                                {
                                    cell.Range.Text = "Lavorazioni :";
                                }
                            }

                            //colonna in centro
                            if (cell.ColumnIndex == 2)
                            {
                                if (cell.RowIndex == 3)
                                {
                                    cell.Range.Text = cliente.ListaProdotti[i].Sconto + "%";
                                }
                                if (cell.RowIndex == 5)
                                {
                                    cell.Range.Text = cliente.ListaProdotti[i].Ricarica + "%";
                                }
                            }

                            //colonna destra
                            if (cell.ColumnIndex == 3)
                            {
                                if (cell.RowIndex == 2)
                                {
                                    cell.Range.Text = cliente.ListaProdotti[i].PrezzoListino.ToString();
                                }
                                if (cell.RowIndex == 4)
                                {
                                    cell.Range.Text = cliente.ListaProdotti[i].Costo.ToString();
                                }
                                if (cell.RowIndex == 6)
                                {
                                    if (cliente.ListaProdotti[i].UnitaPosa)
                                        cell.Range.Text = cliente.ListaProdotti[i].NumeroPezzi + " x " + cliente.ListaProdotti[i].Posa;
                                    else
                                    {
                                        cell.Range.Text = cliente.ListaProdotti[i].Posa.ToString();
                                    }
                                }
                                if (cell.RowIndex == 7)
                                {
                                    if (cliente.ListaProdotti[i].UnitaTrasporto)
                                        cell.Range.Text = cliente.ListaProdotti[i].NumeroPezzi + " x " + cliente.ListaProdotti[i].Trasporto;
                                    else
                                    {
                                        cell.Range.Text = cliente.ListaProdotti[i].Trasporto.ToString();
                                    }
                                }
                                if (cell.RowIndex == 8)
                                {
                                    if (cliente.ListaProdotti[i].UnitaAccessori)
                                        cell.Range.Text = cliente.ListaProdotti[i].NumeroPezzi + " x " + cliente.ListaProdotti[i].Accessori;
                                    else
                                    {
                                        cell.Range.Text = cliente.ListaProdotti[i].Accessori.ToString();
                                    }
                                }
                                if (cell.RowIndex == 9)
                                {
                                    if (cliente.ListaProdotti[i].UnitaLavorazioni)
                                        cell.Range.Text = cliente.ListaProdotti[i].NumeroPezzi + " x " + cliente.ListaProdotti[i].Lavorazioni;
                                    else
                                    {
                                        cell.Range.Text = cliente.ListaProdotti[i].Lavorazioni.ToString();
                                    }
                                }

                            }
                        }
                    }

                    range = GetRange(wordDoc);

                    //stampo la tabella con il totale de singolo Prodotto
                    Word.Table TabTotaleProd = wordDoc.Tables.Add(range, 1, 3, ref missing, ref missing);
                    TabTotaleProd.Borders.Enable = 1;
                    foreach (Word.Row row in TabTotaleProd.Rows)
                    {
                        foreach (Word.Cell cell in row.Cells)
                        {
                            //sistemo il background e allineo il testo

                            cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            //se è nella prima colonna (testo):
                            if (cell.ColumnIndex == 1)
                            {
                                cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                                cell.Range.Text = "Totale Prodotto:";
                                cell.Range.Font.Bold = 1;
                            }

                            //se è nella terza colonna:
                            else if (cell.ColumnIndex == 3)
                            {
                                cell.Range.Text = cliente.ListaProdotti[i].Totale + "";
                            }
                        }
                    }

                    //aggiungo il totale di questo prodotto all'array
                    totali_singoli_prodotti.Add(cliente.ListaProdotti[i].Totale);

                }
                i++;
            }
            i = 0;
            range = GetRange(wordDoc);

            //riassunto tutti totali
            Word.Table tabTotFattura = wordDoc.Tables.Add(range, cliente.ListaProdotti.Count + 2, 2, ref missing, ref missing);
            tabTotFattura.Borders.Enable = 1;
            //scorro tutti i prodotti
            foreach (Prodotto prodotto in cliente.ListaProdotti)
            {
                if (!string.IsNullOrWhiteSpace(cliente.ListaProdotti[i].NomeProdotto))
                {
                    foreach (Word.Row row in tabTotFattura.Rows)
                    {
                        foreach (Word.Cell cell in row.Cells)
                        {
                            cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            if (cell.ColumnIndex == 1)
                            {
                                cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                                if (cell.RowIndex == 1)
                                    cell.Range.Text = "Prodotti:";
                                if (cell.RowIndex == i + 2)
                                    cell.Range.Text = cliente.ListaProdotti[i].NomeProdotto;
                            }
                            if (cell.ColumnIndex == 2)
                            {                                
                                if (cell.RowIndex == 1)
                                {
                                    cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                                    cell.Range.Text = "Totale:";
                                }                                    
                                if (cell.RowIndex == i + 2)
                                    cell.Range.Text = cliente.ListaProdotti[i].Totale.ToString();
                                if (cell.RowIndex == cliente.ListaProdotti.Count + 2)
                                    cell.Range.Text = cliente.TotaleFattura + " €";
                            }
                        }
                    }
                }
                i++;
            }


            //tabella con il riassunto dei totali:
            //chiudo
            var savingPath = @"" + Path.GetFullPath("ArchivioWord") + "\\" + cliente.Nome + cliente.Cognome + cliente.Indice + ".docx";
            wordDoc.SaveAs(savingPath);
            wordDoc.Close();
            wordApp.Quit();
        }

        private static Word.Range GetRange(Word.Document wordDoc)
        {
            object oEndOfDoc = "\\endofdoc";

            Word.Paragraph objParagraph;
            object objRangePara;

            var range = wordDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            objRangePara = wordDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            objParagraph = wordDoc.Content.Paragraphs.Add(ref objRangePara);
            objParagraph.Range.Text = Environment.NewLine;

            return range;
        }
    }    
}
