using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace DLMSExtended
{
    class Excel
    {
        string path = Path.GetFullPath("prog-feladat-jelsz.xlsx");
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel()
        {
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null) return ws.Cells[i, j].Value2;
            else return null;
        }
    }

    class User
    {
        protected int id;
        public string userName;
        protected string password;

        public User(int id, string userName, string password)
        {
            this.id = id;
            this.userName = userName;
            this.password = password;
        }

        public string ValidatePassword()
        {
            string text = id + ". " + userName;
            string passMessage = ": Eros jelszo!\n\n";
            string msgExtension = "";
            bool weakPass = false;
            int ucCount = 0;
            int lcCount = 0;
            int nrCount = 0;
            int spCount = 0;

            for (int j = 0; j < password.Length; j++)
            {
                if (char.IsLetter(password[j]))
                {
                    if (char.IsUpper(password[j])) ucCount++;
                    else lcCount++;
                }
                else if (char.IsNumber(password[j])) nrCount++;
                else spCount++;
            }

            if (password.Length < 8)
            {
                weakPass = true;
                msgExtension += "\nMin. 8 karakter! (Jelenleg: " + password.Length + ")";
            }

            if (ucCount < 2)
            {
                weakPass = true;
                msgExtension += "\nMin. 2 nagy betu! (Jelenleg: " + ucCount + ")";
            }

            if (lcCount < 2)
            {
                weakPass = true;
                msgExtension += "\nMin. 2 kis betu! (Jelenleg: " + lcCount + ")";
            }

            if (nrCount < 2)
            {
                weakPass = true;
                msgExtension += "\nMin. 2 szam! (Jelenleg: " + nrCount + ")";
            }

            if (spCount < 1)
            {
                weakPass = true;
                msgExtension += "\nMin. 1 szpecialis karakter! (Jelenleg: " + spCount + ")";
            }

            if (weakPass)
            {
                passMessage = ": Gyenge jelszo!" + msgExtension + "\n\n";
            }

            text += passMessage;

            return text;
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            Excel ex = new Excel();
            List<User> users = new List<User>(); //Habár a jelenlegi feladatban nincs szerepe a listának, ha bövítésre kerül, elönyössé válhat

            for (int i = 1; ex.ReadCell(i, 1) != null; i++)
            {
                users.Add(new User(i,ex.ReadCell(i, 0), ex.ReadCell(i,1)));
                Console.WriteLine(users[i-1].ValidatePassword());
            }

            Console.ReadKey();
        }
    }
}
