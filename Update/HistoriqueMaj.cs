using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace Metrologo;

/// <summary>Une ligne de l'historique des mises a jour de CE poste.</summary>
public sealed class EntreeHistoriqueMaj
{
    public DateTime Date { get; set; }
    /// <summary>"Installation", "Mise a jour", "Retour arriere" ou "Echec".</summary>
    public string Type { get; set; } = "";
    public string? De { get; set; }
    public string Vers { get; set; } = "";
    public string? Description { get; set; }
}

/// <summary>Operation lancee juste avant un redemarrage Velopack (MAJ ou retour arriere).</summary>
public sealed class OperationMajEnAttente
{
    public string Version { get; set; } = "";
    public string Type { get; set; } = "";
    public string? Description { get; set; }
}

/// <summary>Copie locale du paquet de la version precedente (un seul retour arriere possible).</summary>
public sealed class VersionSauvegardee
{
    public string Version { get; set; } = "";
    public string Fichier { get; set; } = "";
    public string? Description { get; set; }
}

/// <summary>Etat persistant des mises a jour du poste (MAJ\etat-maj.json).</summary>
public sealed class EtatMaj
{
    /// <summary>Version vue au dernier demarrage : un ecart = une MAJ/un retour a eu lieu.</summary>
    public string? VersionConnue { get; set; }
    public OperationMajEnAttente? EnAttente { get; set; }
    /// <summary>Apres un retour arriere : on n'installe plus cette version (ni une plus ancienne).
    /// Levee automatiquement quand une version plus recente est publiee.</summary>
    public string? VersionBloquee { get; set; }
    /// <summary>Anti-boucle : MAJ deja tentee et qui n'a pas pris sur ce poste.</summary>
    public string? TentativeEchouee { get; set; }
    public VersionSauvegardee? Precedente { get; set; }
    public List<EntreeHistoriqueMaj> Historique { get; set; } = new();
}

/// <summary>
/// Historique des mises a jour + sauvegarde de la version precedente, stockes dans
/// <c>%LocalAppData%\Metrologo\MAJ\</c>. Aucune methode ne leve d'exception : un
/// probleme de fichier ne doit jamais empecher l'application de demarrer.
/// </summary>
public static class HistoriqueMaj
{
    public const string TypeInstallation = "Installation";
    public const string TypeMiseAJour = "Mise a jour";
    public const string TypeRetourArriere = "Retour arriere";
    public const string TypeEchec = "Echec";

    public static readonly string Dossier = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Metrologo", "MAJ");

    /// <summary>Dossier qui contient l'unique paquet de la version precedente.</summary>
    public static readonly string DossierPrecedente = Path.Combine(Dossier, "precedente");

    private static readonly string Fichier = Path.Combine(Dossier, "etat-maj.json");

    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

    public static EtatMaj Charger()
    {
        try
        {
            if (File.Exists(Fichier))
                return JsonSerializer.Deserialize<EtatMaj>(File.ReadAllText(Fichier)) ?? new EtatMaj();
        }
        catch { /* fichier abime : on repart d'un etat vide */ }
        return new EtatMaj();
    }

    public static void Sauver(EtatMaj etat)
    {
        try
        {
            Directory.CreateDirectory(Dossier);
            string tmp = Fichier + ".tmp";
            File.WriteAllText(tmp, JsonSerializer.Serialize(etat, Json));
            File.Move(tmp, Fichier, overwrite: true);
        }
        catch { /* sans importance pour le demarrage */ }
    }

    /// <summary>
    /// A appeler a chaque demarrage (version installee connue). Compare a la version
    /// du demarrage precedent pour ajouter la ligne d'historique correspondante, et
    /// detecte une operation lancee qui n'a pas abouti (version inchangee).
    /// </summary>
    public static EtatMaj Synchroniser(string versionCourante)
    {
        var etat = Charger();
        var attente = etat.EnAttente;

        if (etat.VersionConnue != versionCourante)
        {
            string type;
            string? description = null;
            if (attente != null && attente.Version == versionCourante)
            {
                type = attente.Type;
                description = attente.Description;
            }
            else if (etat.VersionConnue == null)
                type = TypeInstallation;
            else
                type = Comparer(versionCourante, etat.VersionConnue) > 0 ? TypeMiseAJour : TypeRetourArriere;

            etat.Historique.Add(new EntreeHistoriqueMaj
            {
                Date = DateTime.Now,
                Type = type,
                De = etat.VersionConnue,
                Vers = versionCourante,
                Description = description,
            });

            // Un seul retour arriere : la sauvegarde est consommee.
            if (type == TypeRetourArriere)
                SupprimerPrecedente(etat);

            etat.VersionConnue = versionCourante;
            etat.TentativeEchouee = null;
        }
        else if (attente != null)
        {
            // Redemarrage effectue mais toujours la meme version : l'operation a echoue.
            etat.Historique.Add(new EntreeHistoriqueMaj
            {
                Date = DateTime.Now,
                Type = TypeEchec,
                De = versionCourante,
                Vers = attente.Version,
                Description = attente.Type == TypeRetourArriere
                    ? "Le retour à la version précédente n'a pas pu être appliqué."
                    : "La mise à jour n'a pas pu être installée sur ce poste.",
            });
            if (attente.Type == TypeMiseAJour)
                etat.TentativeEchouee = attente.Version;
            else
                etat.VersionBloquee = null; // retour rate : on reste sur la version actuelle, sans blocage
        }

        etat.EnAttente = null;
        Sauver(etat);
        return etat;
    }

    /// <summary>Chemin du paquet de la version precedente, ou null s'il n'est plus sur le disque.</summary>
    public static string? CheminPrecedente(EtatMaj etat)
    {
        if (etat.Precedente == null) return null;
        string chemin = Path.Combine(DossierPrecedente, etat.Precedente.Fichier);
        return File.Exists(chemin) ? chemin : null;
    }

    public static void SupprimerPrecedente(EtatMaj etat)
    {
        etat.Precedente = null;
        try
        {
            if (Directory.Exists(DossierPrecedente))
                Directory.Delete(DossierPrecedente, recursive: true);
        }
        catch { /* sera ecrase a la prochaine sauvegarde */ }
    }

    /// <summary>Compare deux numeros de version (numeriques). &lt;0 si a &lt; b.</summary>
    public static int Comparer(string? a, string? b)
    {
        if (Version.TryParse(a, out var va) && Version.TryParse(b, out var vb))
            return va.CompareTo(vb);
        return string.CompareOrdinal(a, b);
    }

    /// <summary>"2.1.0" -> "2.1" ; les anciens numeros dates (1.260706.1451) restent tels quels.</summary>
    public static string Affichage(string? version)
    {
        if (string.IsNullOrEmpty(version)) return "?";
        if (Version.TryParse(version, out var v) && v.Build <= 0 && v.Revision <= 0)
            return $"{v.Major}.{v.Minor}";
        return version;
    }
}
