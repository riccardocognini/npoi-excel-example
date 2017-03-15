using NPOI.Excel.Example.Modelli;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.Excel.Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Il path in cui andrò a salvare il file xlsx creato con NPOI
            string pathOutputXLSX = @"C:/Users/Riccardo/Desktop/Progetti Riccardo/RCSharp/Npoi-Excel-Example/TestFile1.xlsx";
            GeneraFileXLSX(pathOutputXLSX);

            string pathOutputXLSX2 = @"C:/Users/Riccardo/Desktop/Progetti Riccardo/RCSharp/Npoi-Excel-Example/TestFile2.xlsx";

            Persona[] ArrayPersone = new Persona[3];
            ArrayPersone[0] = new Persona() { IdPersona = 1, Nome = "Riccardo", Cognome = "Cognini", DataNascita = new DateTime(1987, 3, 12) };
            ArrayPersone[1] = new Persona() { IdPersona = 1, Nome = "Mario", Cognome = "Rossi", DataNascita = new DateTime(1968, 11, 3) };
            ArrayPersone[2] = new Persona() { IdPersona = 1, Nome = "Luigi", Cognome = "Bianchi", DataNascita = new DateTime(2000, 1, 7) };

            GeneraFileReportXLSXDaModello(pathOutputXLSX2, ArrayPersone);
        }

        public static bool GeneraFileXLSX(string outputFile)
        {
            try
            {
                // Inizializzo l'oggetto NPOI
                XSSFWorkbook wb1 = new XSSFWorkbook();

                // Creo il Foglio di calcolo
                wb1.CreateSheet("Nome del Foglio di Calcolo");

                // Creo una riga nel primo foglio di calcolo (la prima riga del foglio)
                IRow Riga1 = wb1.GetSheetAt(0).CreateRow(0);

                Riga1.CreateCell(0).SetCellValue("A1");
                Riga1.CreateCell(1).SetCellValue("B1");
                Riga1.CreateCell(2).SetCellValue("C1");
                Riga1.CreateCell(3).SetCellValue("D1");
                Riga1.CreateCell(4).SetCellValue("E1");

                // Creo una riga nel primo foglio di calcolo (la seconda riga del foglio)
                IRow Riga2 = wb1.GetSheetAt(0).CreateRow(1);

                Riga2.CreateCell(0).SetCellValue("A2");
                Riga2.CreateCell(1).SetCellValue("B2");
                Riga2.CreateCell(2).SetCellValue("C2");
                Riga2.CreateCell(3).SetCellValue("D2");
                Riga2.CreateCell(4).SetCellValue("E2");

                using (var file = new FileStream(outputFile, FileMode.Create, FileAccess.ReadWrite))
                {
                    wb1.Write(file);
                    file.Close();
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool GeneraFileReportXLSXDaModello(string outputFile, Persona[] ArrayPersone)
        {
            try
            {
                // Inizializzo l'oggetto NPOI
                XSSFWorkbook wb1 = new XSSFWorkbook();

                // Creo il Foglio di calcolo
                wb1.CreateSheet("Nome del Foglio di Calcolo");

                int numeroRiga = 0;

                // Creo una riga nel primo foglio di calcolo (la prima riga del foglio)
                IRow Riga1 = wb1.GetSheetAt(0).CreateRow(numeroRiga);
                numeroRiga++;

                Riga1.CreateCell(0).SetCellValue("IdPersona");
                Riga1.CreateCell(1).SetCellValue("Nome");
                Riga1.CreateCell(2).SetCellValue("Cognome");
                Riga1.CreateCell(3).SetCellValue("Data di Nascita");

                for(int i = 0; i < ArrayPersone.Length; i++)
                {
                    IRow Riga = wb1.GetSheetAt(0).CreateRow(numeroRiga);

                    Riga.CreateCell(0).SetCellValue(ArrayPersone[i].IdPersona);
                    Riga.CreateCell(1).SetCellValue(ArrayPersone[i].Nome);
                    Riga.CreateCell(2).SetCellValue(ArrayPersone[i].Cognome);
                    Riga.CreateCell(3).SetCellValue(ArrayPersone[i].DataNascita.ToString("dd/MM/yyyy"));

                    numeroRiga++;
                }

                using (var file = new FileStream(outputFile, FileMode.Create, FileAccess.ReadWrite))
                {
                    wb1.Write(file);
                    file.Close();
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}
