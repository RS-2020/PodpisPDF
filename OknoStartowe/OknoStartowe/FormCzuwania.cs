using iText.Kernel.Pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.X509;
using System.IO;
using iText.Signatures;
using System.Runtime.InteropServices;
using System.Net;
using System.Security.Principal;
using System.Net.Sockets;

namespace OknoStartowe
{
    public partial class FormCzuwania : Form
    {
        private TcpListener listener = null;
        private TcpClient klient = null;
        private bool czypolaczono
        {
            get
            {
                if (klient != null)
                {
                    return klient.Connected;
                }
                else
                {
                    return false;
                }
            }
        }
        private BinaryReader r = null;
        private BinaryWriter w = null;
        public FormCzuwania()
        {
            InitializeComponent();
            UstawSerwer();
        }
        public static void Czuwanie()
        {
            FormCzuwania formCzuwania = new FormCzuwania();
            Application.Run(formCzuwania);
        }
        private void FormCzuwania_Shown(object sender, EventArgs e)
        {
            Hide();
        }
        //private void StartCzuwana()
        //{
        //    string FolderDanych = @"C:\TMP\watch\";
        //    if (Directory.Exists(FolderDanych) == false)
        //    {
        //        Directory.CreateDirectory(FolderDanych);
        //    }
        //    FileSystemWatcher FSW = new FileSystemWatcher(FolderDanych, "*.txt");
        //    FSW.EnableRaisingEvents = true;
        //    FSW.IncludeSubdirectories = false;
        //    FSW.NotifyFilter = NotifyFilters.Attributes
        //                         | NotifyFilters.CreationTime
        //                         | NotifyFilters.DirectoryName
        //                         | NotifyFilters.FileName
        //                         | NotifyFilters.LastAccess
        //                         | NotifyFilters.LastWrite
        //                         | NotifyFilters.Security
        //                         | NotifyFilters.Size;
        //    //FSW.WaitForChanged(WatcherChangeTypes.Created);
        //    FSW.Created += FSW_Created;
        //}
        //private void FSW_Created(object sender, FileSystemEventArgs e)
        //{
        //    string TxtFFN = e.FullPath;
        //    try
        //    {
        //        List<string> TrescPliku = File.ReadAllLines(TxtFFN).ToList();
        //        if (TrescPliku.Count>0 && File.Exists(TrescPliku[0]))
        //        {
        //            string Podpisy = Program.OdczytPojedynczegoPDF(TrescPliku[0]);
        //            File.WriteAllText(TxtFFN, Podpisy);
        //        }
        //    }
        //    catch { }
        //}
        private void UstawSerwer()
        {
            Polaczenie.RunWorkerAsync();
            //MessageBox.Show("Serwer stoi");
        }
        private void Rozlacz()
        {
            listener.Stop();
            if (klient != null) klient.Close();
            Polaczenie.CancelAsync();
            Odbieranie.CancelAsync();
            //czypolaczono = false;
            Polaczenie.RunWorkerAsync();
        }
        private void WyslijDane(string DoWyslania)
        {
            w.Write(DoWyslania);
        }
        private void Polaczenie_DoWork(object sender, DoWorkEventArgs e)
        {
            IPAddress AdresIP = null;
            IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (IPAddress ip in host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    AdresIP = ip;
                    break;
                }
            }
            listener = new TcpListener(AdresIP, 8000);

            listener.Start();
            while (!listener.Pending())
            {
                if (this.Polaczenie.CancellationPending)
                {
                    if (klient != null) klient.Close();
                    listener.Stop();
                    //czypolaczono = false;
                    return;
                }
            }
            klient = listener.AcceptTcpClient();
            NetworkStream stream = klient.GetStream();
            w = new BinaryWriter(stream);
            r = new BinaryReader(stream);
            if (r.ReadString() != "DISCONNECT")
            {
                //czypolaczono = true;
                w.Write("CONNECTED");
                Odbieranie.RunWorkerAsync();
            }
            else
            {
                if (klient != null) klient.Close();
                listener.Stop();
                //czypolaczono = false;
            }
        }
        private void Odbieranie_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                string tekst;
                while ((tekst = r.ReadString()) != "DISCONNECT")
                {
                    //MessageBox.Show(tekst, "SERWER");
                    if (tekst == "KILL") { this.Close(); }
                    PytanieOPDF(tekst);
                }
                Rozlacz();
                //czypolaczono = false;
                //klient.Close();
                //listener.Stop();
            }
            catch (Exception ex)
            {
                Rozlacz();
            }
        }
        private void PytanieOPDF(string Zapytanie)
        {
            Zapytanie = Zapytanie.Replace("read ", "");
            string Odpowiedz = Program.OdczytPojedynczegoPDF(Zapytanie);
            WyslijDane(Odpowiedz);
        }
    }
}
