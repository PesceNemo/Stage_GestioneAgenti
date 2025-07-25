using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.ComponentModel;
using System.Windows.Data;


namespace GestioneAgenti
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<Clienti> dati { get; set; }
        public List<StatoManualeEnum> StatiManualiDisponibili { get; set; }

        private ICollectionView datiView;

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

            dati = new ObservableCollection<Clienti>(){};

            dati = RecuperoDatiDB();
            datiView = CollectionViewSource.GetDefaultView(dati);
            dataGrid.ItemsSource = datiView;
        }

        private ObservableCollection<Clienti> RecuperoDatiDB()
        {

            ObservableCollection<Clienti> clienti = new ObservableCollection<Clienti>();


            string connectionString = "Server=192.168.1.200;Database=Concept;User ID=sa;Password=Gv2021!;TrustServerCertificate=True";
            using (SqlConnection sql = new SqlConnection(connectionString))
            {
                sql.Open();

                var numero = 3;

                var  = $@"
               
                    WITH scadenze
                    AS (
                        SELECT DISTINCT DSPAG AS PAGATO
                            , DSKEY
                        FROM Concept.dbo.A01_SCA_MOV
                        )
                        , documenti
                    AS (
                        SELECT (
                                SELECT TOP 1 IDTES
                                FROM Concept.dbo.A01_DOC_VER
                                WHERE DRORI IN (
                                        SELECT RECORD_ID
                                        FROM Concept.dbo.A01_doc_ver
                                        WHERE IDTES = ven.RECORD_ID
                                        )
                                ) IDFAT
                            , ven.RECORD_ID
                            , CONCAT (
                                DTCAU
                                , ' '
                                , DTANN
                                , '/'
                                , DTNUM
                                , ' '
                                , DTSER, ' ',
                                DTRAS
                                ) ORIGINE
                            , CONCAT (
                                AGNOM
                                , ' '
                                , AGCOG
                                , ' - '
                                , CDAGE
                                ) AS AGENTE
                            , FLVAL CASH
                            , DTCAU CAUSALE
                            , DTTOT TOTALE_DOCUMENTO
                            , DTIMP IMPONIBILE
                            , DTUSA RIFERIMENTO_INTERNO
                        FROM Concept.dbo.A01_DOC_VEN ven
                        LEFT JOIN Concept.dbo.A01_FLD_VAL
                            ON FLTAB = 'DOC_VEN'
                                AND FLKEY = ven.RECORD_ID
                                AND flFLD = 'GV_PRXE'
                        INNER JOIN Concept.dbo.A01_AGE_NTI
                            ON DTAGE = CDAGE
                        WHERE DTAGE <> ''
                            AND DTTPD = 'ORD'
                            AND (
                                SELECT TOP 1 vv.DTTPD
                                FROM Concept.dbo.A01_DOC_VER rr
                                INNER JOIN Concept.dbo.A01_DOC_VEN vv
                                    ON vv.RECORD_ID = rr.IDTES
                                WHERE DRORI IN (
                                        SELECT RECORD_ID
                                        FROM Concept.dbo.A01_doc_ver
                                        WHERE IDTES = ven.RECORD_ID
                                        )
                                ) = 'FAT'
                        )
                    SELECT AGENTE,CASH,ORIGINE, (
                            SELECT CASE
                                    WHEN EXISTS(
                                            SELECT 1
                                            FROM scadenze
                                            WHERE PAGATO = 0
                                                AND DSKEY = documenti.IDFAT
                                            )
                                        THEN 0
                                    ELSE 1
                                    END AS RESULT
                            ) PAGATO, GETDATE(), IDFAT, (SELECT FLVAL FROM A01_FLD_VAL WHERE FLFLD = 'GV_STAT' AND FLKEY = IDFAT),CAUSALE,TOTALE_DOCUMENTO,IMPONIBILE,RIFERIMENTO_INTERNO
                    FROM documenti
                    WHERE idfat IS NOT NULL
                    ";

                using (SqlDataReader reader = new SqlCommand(query, sql).ExecuteReader())
                {
                    while (reader.Read())
                    {
                        try
                        {
                            // Handle possible DBNull values and parsing errors gracefully
                            var cashValue = reader[1] != DBNull.Value ? int.Parse(reader[1].ToString()) : 0; // Default to 0 if null
                            var ordineOrigine = reader[2] != DBNull.Value ? reader.GetString(2) : string.Empty; // Default to empty string if null
                            var pagato = reader[3].ToString() == "1" ? true : false; // Default to empty string if null
                            var stato = reader.IsDBNull(6) ? StatoManualeEnum.DaGestire : (StatoManualeEnum)int.Parse(reader.GetString(6));
                            clienti.Add(new Clienti()
                            {
                                Agente = reader[0].ToString(), // Assuming reader[0] is always non-null
                                Cash = cashValue,
                                OrdineOrigine = ordineOrigine,
                                DataFatturaSaldata = DateTime.Now,
                                Pagato = pagato,
                                IdFat = reader.GetInt32(5),    
                                StatoManuale = stato,
                            });
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing record: {reader[2].ToString()}, Exception: {ex.Message}");
                        }
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
            string connectionString = "Server=192.168.1.200;Database=Concept;User ID=sa;Password=Gv2021!;TrustServerCertificate=True";

            using (SqlConnection sql = new SqlConnection(connectionString))
            {
                sql.Open();

                string query = @"
                UPDATE A01_FLD_VAL
                SET FLVAL = @stato
                WHERE FLKEY = @VALUE AND FLFLD = 'GV_STAT'
";

                using (SqlCommand cmd = new SqlCommand(query, sql))
                {
                    cmd.Parameters.AddWithValue("@STATO", clienti.StatoManuale.GetValueOrDefault());
                    cmd.Parameters.AddWithValue("@VALUE", clienti.IdFat);
                    cmd.ExecuteNonQuery();
                }
            }
        }
        private void BtnEsportaExcel_Click(object sender, RoutedEventArgs e)
        {
            EsportaInExcel();
        }

        private void EsportaInExcel()
        {
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = "EsportazioneClienti.xlsx"
            };

            if (saveDialog.ShowDialog() == true)
            {
                using (var workbook = new ClosedXML.Excel.XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Clienti");

                    // Intestazioni
                    ws.Cell(1, 1).Value = "Agente";
                    ws.Cell(1, 2).Value = "Cash";
                    ws.Cell(1, 3).Value = "Ordine Origine";
                    ws.Cell(1, 4).Value = "Pagato";
                    ws.Cell(1, 5).Value = "Data Fattura Saldata";
                    ws.Cell(1, 6).Value = "Stato Manuale";

                    // Dati
                    for (int i = 0; i < dati.Count; i++)
                    {
                        var c = dati[i];
                        ws.Cell(i + 2, 1).Value = c.Agente;
                        ws.Cell(i + 2, 2).Value = c.Cash;
                        ws.Cell(i + 2, 3).Value = c.OrdineOrigine;
                        ws.Cell(i + 2, 4).Value = c.Pagato ? "Sì" : "No";
                        ws.Cell(i + 2, 5).Value = c.DataFatturaSaldata.ToShortDateString();
                        ws.Cell(i + 2, 6).Value = c.StatoManuale?.ToString() ?? "DaGestire";
                    }

                    ws.Columns().AdjustToContents();

                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Esportazione completata!", "Excel", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }
        private void FiltroTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filtro = FiltroTextBox.Text?.ToLower() ?? string.Empty;

            datiView.Filter = obj =>
            {
                if (obj is Clienti cliente)
                {
                    return (cliente.Agente?.ToLower().Contains(filtro) ?? false) ||
                           (cliente.OrdineOrigine?.ToLower().Contains(filtro) ?? false);
                }
                return false;
            };
        }

    }
}