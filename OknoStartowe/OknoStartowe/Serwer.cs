using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OknoStartowe
{
    public static class Serwer
    {
        public static void START()
        {
            while (true)
            {
                start();
            }
            //Task.Run(() => START());
        }
        private static void start()
        {
            string ID_SesjiOdpowiedzi = "";
            string ID_Funkcji = "";
            using (NamedPipeServerStream serverStream = new NamedPipeServerStream("PodpisPDF"))
            {
                serverStream.WaitForConnection();
                using (StreamReader reader = new StreamReader(serverStream, Encoding.UTF8))
                {
                    string message = reader.ReadToEnd();
                    List<string> Odebrano = message.Split('\n').ToList();
                    if (Odebrano.Count > 1)
                    {
                        ID_Funkcji = Odebrano[0].Replace("\r", "");
                        ID_SesjiOdpowiedzi = Odebrano[1].Replace("\r", "");
                        if (ID_Funkcji == "SignPDF")
                        {
                            Task.Run(() => SignPDF(ID_SesjiOdpowiedzi, Odebrano[2].Replace("\r", ""), int.Parse(Odebrano[3].Replace("\r", ""))));
                        }
                        else if (ID_Funkcji == "ReadSign")
                        {
                            Task.Run(() => ReadSign(ID_SesjiOdpowiedzi, Odebrano[2].Replace("\r", "")));
                        }
                    }
                }
            }
        }
        private static void Nadawanie(string ID_SesjiOdpowiedzi, List<string> WierszeDoWysylki)
        {
            using (NamedPipeClientStream clientStream = new NamedPipeClientStream(".", ID_SesjiOdpowiedzi, PipeDirection.Out))
            {
                clientStream.Connect();
                using (StreamWriter writer = new StreamWriter(clientStream, Encoding.UTF8))
                {
                    foreach (string Wiersz in WierszeDoWysylki)
                    {
                        writer.WriteLine(Wiersz);
                    }
                }
            }
        }
        private static void SignPDF(string ID_SesjiOdpowiedzi, string PowodPodpisania, int Wysokosc)
        {

        }
        private static void ReadSign(string ID_SesjiOdpowiedzi, string SciezkaDoPDF)
        {
            List<OdczytPodpisow.Podpis> ListaPodpisow = OdczytPodpisow.PodpisanePrzez(SciezkaDoPDF);
            List<string> Przekonwertowana = new List<string>();
            ListaPodpisow.ForEach(poz => Przekonwertowana.Add(poz.ToString()));
            Nadawanie(ID_SesjiOdpowiedzi, Przekonwertowana);
        }
    }
}
