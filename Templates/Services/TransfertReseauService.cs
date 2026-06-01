using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Journal;

namespace Metrologo.Services
{
    /// <summary>
    /// Transfert du dossier FI complet (rapport .xlsx + Journal_FI.txt + profilings + tout
    /// le contenu de <c>C:\Users\…\Desktop\Metrologo\&lt;FI&gt;\</c>) vers le partage réseau
    /// (<c>CheminsMetrologo.MesuresLocal</c>, typiquement <c>M:\…\Mesures\&lt;FI&gt;\</c>).
    ///
    /// Stratégie :
    /// <list type="bullet">
    ///   <item>Pendant la mesure : aucune I/O réseau (tout reste en local).</item>
    ///   <item>À la fin de chaque mesure : tentative de transfert complet du dossier FI.</item>
    ///   <item>Si transfert échoue (M:\ down, latence, etc.) : on enregistre la FI dans
    ///         <c>%LocalAppData%\Metrologo\Configuration\transferts_en_attente.json</c>
    ///         pour reprise automatique au prochain démarrage de l'app.</item>
    ///   <item>Au démarrage : on rejoue tous les transferts en attente si M:\ accessible.</item>
    /// </list>
    /// </summary>
    public static class TransfertReseauService
    {
        private static readonly object _sync = new();

        /// <summary>Chemin du fichier qui mémorise les FI en attente de transfert.</summary>
        public static string CheminFichierEnAttente =>
            Path.Combine(CheminsMetrologo.Configuration, "transferts_en_attente.json");

        /// <summary>
        /// Transfère le dossier local d'une FI (<c>Desktop\Metrologo\&lt;FI&gt;\</c>) vers
        /// son équivalent réseau. Si le transfert échoue ou que M:\ n'est pas configuré,
        /// la FI est ajoutée à la liste en attente.
        ///
        /// Retourne <c>true</c> si le transfert s'est terminé avec succès, <c>false</c> sinon.
        /// </summary>
        public static async Task<bool> TransfererDossierFIAsync(string numFI)
        {
            if (string.IsNullOrWhiteSpace(numFI)) return false;

            return await Task.Run(() =>
            {
                string numFISafe = SanitizerNomFichier(numFI);
                string dossierLocal = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "Metrologo", numFISafe);

                if (!Directory.Exists(dossierLocal))
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_RESEAU_DOSSIER_INTROUVABLE",
                        $"Dossier local de la FI {numFI} introuvable : {dossierLocal}");
                    return false;
                }

                if (!CheminsMetrologo.MesuresLocalConfigure)
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_RESEAU_PAS_CONFIGURE",
                        "Aucun chemin réseau configuré (Admin > Chemins d'accès) — "
                        + "le dossier reste en local.");
                    AjouterFIEnAttente(numFI);
                    return false;
                }

                try
                {
                    string dossierCible = Path.Combine(CheminsMetrologo.MesuresLocal, numFISafe);
                    Directory.CreateDirectory(dossierCible);
                    CopierDossierRecursif(dossierLocal, dossierCible);
                    RetirerFIEnAttente(numFI);
                    Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_RESEAU_OK",
                        $"Dossier FI {numFI} transféré sur le réseau : {dossierCible}");
                    return true;
                }
                catch (Exception ex)
                {
                    AjouterFIEnAttente(numFI);
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_RESEAU_KO",
                        $"Transfert FI {numFI} vers le réseau échoué : {ex.Message} — "
                        + "FI ajoutée à la liste de reprise au prochain démarrage.");
                    return false;
                }
            });
        }

        /// <summary>
        /// Au démarrage de l'application, tente de rejouer tous les transferts en attente
        /// (FI qui n'ont pas pu être copiées sur le réseau lors des sessions précédentes).
        /// Best-effort : si le réseau est toujours indisponible, les FI restent dans la liste
        /// pour la prochaine tentative.
        /// </summary>
        public static async Task TenterTransfertsEnAttenteAsync()
        {
            var enAttente = LireFIEnAttente();
            if (enAttente.Count == 0) return;

            if (!CheminsMetrologo.MesuresLocalConfigure)
            {
                Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_REPRISE_PAS_CONFIGURE",
                    $"{enAttente.Count} FI en attente de transfert, mais aucun chemin réseau "
                    + "configuré — report à la prochaine session.");
                return;
            }

            Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_REPRISE_DEBUT",
                $"Reprise auto de {enAttente.Count} transfert(s) FI en attente : "
                + string.Join(", ", enAttente));

            int nbOk = 0, nbKo = 0;
            foreach (var fi in enAttente.ToList())
            {
                bool ok = await TransfererDossierFIAsync(fi);
                if (ok) nbOk++; else nbKo++;
            }

            Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_REPRISE_FIN",
                $"Reprise terminée : {nbOk} OK, {nbKo} KO (restent en attente).");
        }

        /// <summary>Retourne la liste des FI actuellement en attente de transfert.</summary>
        public static List<string> LireFIEnAttente()
        {
            lock (_sync)
            {
                try
                {
                    if (!File.Exists(CheminFichierEnAttente)) return new List<string>();
                    string json = File.ReadAllText(CheminFichierEnAttente);
                    var liste = JsonSerializer.Deserialize<List<string>>(json);
                    return liste ?? new List<string>();
                }
                catch (Exception ex)
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_LISTE_LECTURE_KO",
                        $"Lecture de la liste des transferts en attente échouée : {ex.Message}");
                    return new List<string>();
                }
            }
        }

        private static void AjouterFIEnAttente(string numFI)
        {
            lock (_sync)
            {
                try
                {
                    var liste = LireFIEnAttenteSansLock();
                    if (!liste.Contains(numFI, StringComparer.OrdinalIgnoreCase))
                    {
                        liste.Add(numFI);
                        SauvegarderListe(liste);
                    }
                }
                catch (Exception ex)
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_LISTE_AJOUT_KO",
                        $"Ajout FI {numFI} en attente échoué : {ex.Message}");
                }
            }
        }

        private static void RetirerFIEnAttente(string numFI)
        {
            lock (_sync)
            {
                try
                {
                    var liste = LireFIEnAttenteSansLock();
                    int avant = liste.Count;
                    liste.RemoveAll(f => string.Equals(f, numFI, StringComparison.OrdinalIgnoreCase));
                    if (liste.Count != avant) SauvegarderListe(liste);
                }
                catch { /* best-effort */ }
            }
        }

        private static List<string> LireFIEnAttenteSansLock()
        {
            if (!File.Exists(CheminFichierEnAttente)) return new List<string>();
            string json = File.ReadAllText(CheminFichierEnAttente);
            var liste = JsonSerializer.Deserialize<List<string>>(json);
            return liste ?? new List<string>();
        }

        private static void SauvegarderListe(List<string> liste)
        {
            string dossier = Path.GetDirectoryName(CheminFichierEnAttente)!;
            Directory.CreateDirectory(dossier);
            string json = JsonSerializer.Serialize(liste,
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(CheminFichierEnAttente, json);
        }

        private static void CopierDossierRecursif(string source, string cible)
        {
            Directory.CreateDirectory(cible);
            foreach (var fichier in Directory.EnumerateFiles(source))
            {
                string nomFichier = Path.GetFileName(fichier);
                string fichierCible = Path.Combine(cible, nomFichier);
                // overwrite=true : on synchronise en écrasant côté réseau (le local fait foi).
                File.Copy(fichier, fichierCible, overwrite: true);
            }
            // Récursion sur les sous-dossiers (si jamais un jour il y en a — ex: dossier
            // d'archives à l'intérieur d'une FI). Aujourd'hui le dossier FI est plat.
            foreach (var sousDossier in Directory.EnumerateDirectories(source))
            {
                string nomSousDossier = Path.GetFileName(sousDossier);
                CopierDossierRecursif(sousDossier, Path.Combine(cible, nomSousDossier));
            }
        }

        private static string SanitizerNomFichier(string nom)
        {
            if (string.IsNullOrWhiteSpace(nom)) return "sans-nom";
            var invalides = new HashSet<char>(Path.GetInvalidFileNameChars());
            var sb = new System.Text.StringBuilder(nom.Length);
            foreach (var c in nom)
            {
                sb.Append(invalides.Contains(c) ? '_' : c);
            }
            string resultat = sb.ToString().Trim(' ', '.');
            return string.IsNullOrEmpty(resultat) ? "sans-nom" : resultat;
        }
    }
}
