using GestioneAgenti;
using System;

public class Clienti
{
    public string Agente { get; set; }
    public decimal Cash { get; set; }
    public string OrdineOrigine { get; set; }
    public bool Pagato { get; set; }
    public DateTime DataFatturaSaldata { get; set; }
    public int IdFat { get; set; }
    public StatoManualeEnum? StatoManuale { get; set; } = StatoManualeEnum.DaGestire;
}
