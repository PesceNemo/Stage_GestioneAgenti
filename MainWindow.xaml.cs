using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;

namespace GestioneAgenti
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<Clienti> dati { get; set; }
        public List<StatoManualeEnum> StatiManualiDisponibili { get; set; }

        public MainWindow()
        {
            InitializeComponent();

            // Imposto DataContext per binding XAML
            this.DataContext = this;

            // Lista dei valori disponibili per ComboBox
            StatiManualiDisponibili = new List<StatoManualeEnum>()
            {
                StatoManualeEnum.Completato,
                StatoManualeEnum.InLavorazione,
                StatoManualeEnum.DaGestire
            };

            // Dati di esempio 
            dati = new ObservableCollection<Clienti>()
            {
                //new Clienti
                //{
                //    Agente = "Mario Rossi",
                //    Cash = 1200.50,
                //    OrdineOrigine = "ORD001",
                //    Pagato = false,
                //    DataFatturaSaldata = new DateTime(2025, 7, 22), 
                //    StatoManuale = null
                //},
                //new Clienti
                //{
                //    Agente = "Luca Bianchi",
                //    Cash = 850.00,
                //    OrdineOrigine = "ORD002",
                //    Pagato = false,
                //    DataFatturaSaldata = new DateTime(2025, 4, 30), 
                //    StatoManuale = null
                //},
                //new Clienti
                //{
                //    Agente = "Luca Bianchi",
                //    Cash = 850.00,
                //    OrdineOrigine = "ORD002",
                //    Pagato = false,
                //    DataFatturaSaldata = new DateTime(2025, 4, 30),
                //    StatoManuale = null
                //},
                //new Clienti
                //{
                //    Agente = "Luca Bianchi",
                //    Cash = 850.00,
                //    OrdineOrigine = "ORD002",
                //    Pagato = false,
                //    DataFatturaSaldata = new DateTime(2025, 4, 30),
                //    StatoManuale = null
                //}
            };


            dati = RecuperoDatiDB();





        }

        private ObservableCollection<Clienti> RecuperoDatiDB()
        {

            ObservableCollection<Clienti> clienti = new ObservableCollection<Clienti>();


            string connectionString = "Server=192.168.1.99;Database=GV_NOEMI;User ID=sa;Password=GvC0nsulting!";
            using (SqlConnection sql = new SqlConnection(connectionString))
            {
                sql.Open();

                var numero = 3;

                var query = $@"
                SELECT AGENTE,CASH,ORIGINE,PAGATO,DATA_FATTURA_SALDATA,STATO
                FROM AGENTI
";

                using (SqlDataReader reader = new SqlCommand(query, sql).ExecuteReader())
                {
                    while (reader.Read())
                    {

                        clienti.Add(new Clienti()
                        {
                            Agente = reader.GetString(0),
                            Cash = reader.GetDecimal(1),
                            OrdineOrigine = reader.GetString(2),
                            Pagato = reader.GetBoolean(3),
                            DataFatturaSaldata = reader.GetDateTime(4),
                            StatoManuale = reader.IsDBNull(5) ? StatoManualeEnum.DaGestire : (StatoManualeEnum)reader.GetInt32(5),
                        });
                    }
                }
            }
            return clienti;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = sender as ComboBox;
            if (comboBox?.DataContext is Clienti clienti && comboBox.SelectedItem is StatoManualeEnum nuovoValore)
            {
                clienti.StatoManuale = nuovoValore;
                AggiornaStatoNelDB(clienti);
            }
        }


        private void AggiornaStatoNelDB(Clienti clienti)
        {
            string connectionString = "Server=192.168.1.99;Database=GV_NOEMI;User ID=sa;Password=GvC0nsulting!";

            using (SqlConnection sql = new SqlConnection(connectionString))
            {
                sql.Open();

                string query = @"
                UPDATE AGENTI
                SET STATO = @stato
                WHERE AGENTE = @agente AND ORIGINE = @origine";

                using (SqlCommand cmd = new SqlCommand(query, sql))
                {
                    cmd.Parameters.AddWithValue("@stato", (int)clienti.StatoManuale.GetValueOrDefault());
                    cmd.Parameters.AddWithValue("@agente", clienti.Agente);
                    cmd.Parameters.AddWithValue("@origine", clienti.OrdineOrigine);
                    cmd.ExecuteNonQuery();
                }
            }
        }

    }
}