using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Services.Besancon;
using System;
using System.Globalization;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Paramètres de la récupération automatique Besançon (fichier partagé
    /// <c>besancon.ftp.json</c>) : activation de la tâche quotidienne, heure de déclenchement et
    /// identifiants FTP. Évite d'éditer le JSON à la main. L'enregistrement applique aussitôt la
    /// nouvelle config au planificateur (<see cref="BesanconScheduler.Reconfigurer"/>).
    /// </summary>
    public partial class BesanconParametresViewModel : ObservableObject
    {
        [ObservableProperty] private bool _active;
        [ObservableProperty] private string _heureTexte = "09:38";
        [ObservableProperty] private string _ftpHote = "";
        [ObservableProperty] private int _ftpPort = 21;
        [ObservableProperty] private string _ftpUtilisateur = "";
        [ObservableProperty] private string _ftpMotDePasse = "";
        [ObservableProperty] private bool _ftpSsl;
        [ObservableProperty] private string _fichierDistant = "ef_utcop";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = "";

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        /// <summary>Ferme la fenêtre ; <c>true</c> = enregistré, <c>false</c> = annulé.</summary>
        public Action<bool>? CloseAction { get; set; }

        public BesanconParametresViewModel()
        {
            var cfg = BesanconConfig.Charger();
            Active = cfg.Active;
            HeureTexte = cfg.HeureDeclenchement;
            FtpHote = cfg.FtpHote;
            FtpPort = cfg.FtpPort;
            FtpUtilisateur = cfg.FtpUtilisateur;
            FtpMotDePasse = cfg.FtpMotDePasse;
            FtpSsl = cfg.FtpSsl;
            FichierDistant = cfg.FichierDistant;
        }

        [RelayCommand]
        private void Enregistrer()
        {
            string h = (HeureTexte ?? "").Trim();
            if (!TimeSpan.TryParseExact(h, "hh\\:mm", CultureInfo.InvariantCulture, out _))
            {
                MessageErreur = "Heure invalide. Format attendu : HH:mm (ex. 09:38).";
                return;
            }
            if (FtpPort < 1 || FtpPort > 65535)
            {
                MessageErreur = "Port FTP invalide (1 – 65535).";
                return;
            }

            // On recharge la config partagée pour ne pas écraser d'éventuels champs non exposés
            // ici (ex. SupprimerApresTelechargement, flags de migration), puis on applique nos champs.
            var cfg = BesanconConfig.Charger();
            cfg.Active = Active;
            cfg.HeureDeclenchement = h;
            cfg.FtpHote = (FtpHote ?? "").Trim();
            cfg.FtpPort = FtpPort;
            cfg.FtpUtilisateur = (FtpUtilisateur ?? "").Trim();
            cfg.FtpMotDePasse = FtpMotDePasse ?? "";
            cfg.FtpSsl = FtpSsl;
            cfg.FichierDistant = string.IsNullOrWhiteSpace(FichierDistant) ? "ef_utcop" : FichierDistant.Trim();
            cfg.Sauvegarder();

            // Applique aussitôt : (re)programme à la nouvelle heure ou arrête si désactivé.
            BesanconScheduler.Reconfigurer();

            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
